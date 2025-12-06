import os
import re
try:
    import pythoncom
    import win32com.client as win32
except Exception as e:
    raise ImportError(
        "The module 'pythoncom' (part of pywin32) is required on Windows. "
        "Install it into your Python environment with:\n\n"
        "    python -m pip install pywin32\n\n"
        "After installing, if the import still fails, ensure you're using the same Python interpreter.\n"
        f"Original error: {e}"
    )

class ReferenceValidator:
    def __init__(self, filepath):
        self.filepath = os.path.abspath(filepath)
        self.word = None
        self.doc = None
        self.results = {
            'total_references': 0,
            'total_citations': 0,
            'missing_references': set(),
            'unused_references': set(),
            'sequence_issues': [],
            'citation_sequence': []
        }

    def __enter__(self):
        pythoncom.CoInitialize()
        self.word = win32.Dispatch("Word.Application")
        self.word.Visible = False
        self.word.ScreenUpdating = False
        # Open read-only by default for validation. Renumbering will reopen writable when needed.
        self.doc = self.word.Documents.Open(self.filepath, ReadOnly=True)
        return self

    def __exit__(self, exc_type, exc_val, exc_tb):
        if self.doc:
            self.doc.Close(SaveChanges=False)
        if self.word:
            self.word.Quit()
        pythoncom.CoUninitialize()

    def validate(self, auto_renumber=False, save_path=None):
        ref_numbers = self._get_reference_numbers()
        self.results['total_references'] = len(ref_numbers)

        citations = self._get_citations()
        self.results['total_citations'] = len(citations)

        cited_numbers = set()
        for citation in citations:
            cited_numbers.update(citation['numbers'])
            self.results['citation_sequence'].extend(citation['numbers'])

        self.results['missing_references'] = sorted(cited_numbers - ref_numbers)
        self.results['unused_references'] = sorted(ref_numbers - cited_numbers)
        self._check_citation_sequence()
        # Optionally auto-renumber when sequence issues are found
        if auto_renumber and 'NOT in sequence' in self.results.get('sequence_message', ''):
            ren = self.renumber_if_needed(save_path=save_path)
            self.results['renumber_attempt'] = ren
            # If renumbering happened, refresh results by reopening the saved/original document
            if ren.get('renumbered'):
                reopen_path = save_path if save_path else self.filepath
                try:
                    if self.doc:
                        self.doc.Close(SaveChanges=False)
                except Exception:
                    pass
                try:
                    self.doc = self.word.Documents.Open(os.path.abspath(reopen_path), ReadOnly=True)
                except Exception:
                    try:
                        self.doc = self.word.Documents.Open(reopen_path)
                    except Exception:
                        # couldn't reopen, return current results with renumber info
                        return self.results

                # recompute results
                ref_numbers = self._get_reference_numbers()
                self.results['total_references'] = len(ref_numbers)
                citations = self._get_citations()
                self.results['total_citations'] = len(citations)
                cited_numbers = set()
                self.results['citation_sequence'] = []
                for citation in citations:
                    cited_numbers.update(citation['numbers'])
                    self.results['citation_sequence'].extend(citation['numbers'])
                self.results['missing_references'] = sorted(cited_numbers - ref_numbers)
                self.results['unused_references'] = sorted(ref_numbers - cited_numbers)
                self._check_citation_sequence()

        return self.results

    def renumber_if_needed(self, save_path=None):
        """
        Reorder references based on citation appearance order and renumber everything.
        This will physically move reference paragraphs to match the citation sequence.
        If citation sequence is not ordered, renumber citations and reference list.
        Writes changes back to the document (overwrites original unless save_path provided).
        """
        # Ensure we have the latest sequence data
        sequence = self.results.get('citation_sequence', [])
        if not sequence:
            return {'renumbered': False, 'message': 'No citations found.'}

        seen = set()
        unique_sequence = []
        for num in sequence:
            if num not in seen:
                unique_sequence.append(num)
                seen.add(num)

        is_ordered = unique_sequence == sorted(unique_sequence)
        if is_ordered:
            return {'renumbered': False, 'message': 'Citations already in sequence.'}

        # Build renumber map: first-appearance order becomes 1..n. Include any reference-only numbers afterwards.
        ref_numbers = sorted(list(self._get_reference_numbers()))
        ordered_all = list(unique_sequence)
        
        # Add any references that weren't cited (append them at the end in their original relative order)
        for n in ref_numbers:
            if n not in seen:
                ordered_all.append(n)

        renumber_map = {old: new for new, old in enumerate(ordered_all, start=1)}

        # Reopen document writable
        try:
            if self.doc:
                self.doc.Close(SaveChanges=False)
            self.doc = self.word.Documents.Open(self.filepath, ReadOnly=False)
        except Exception:
            # try opening without ReadOnly named arg
            self.doc = self.word.Documents.Open(self.filepath)

        # Update citations in text
        self._update_citations(renumber_map)
        
        # Reorder and renumber reference list
        self._reorder_references(renumber_map, ordered_all)

        # Save document
        try:
            if save_path:
                self.doc.SaveAs(FileName=os.path.abspath(save_path))
            else:
                self.doc.Save()
        except Exception:
            # Some Word versions use SaveAs2
            try:
                if save_path:
                    self.doc.SaveAs2(FileName=os.path.abspath(save_path))
                else:
                    self.doc.Save()
            except Exception as e:
                return {'renumbered': False, 'message': f'Failed to save document: {e}'}

        return {'renumbered': True, 'map': renumber_map}

    def _update_citations(self, renumber_map):
        """Update all citation numbers in the text."""
        try:
            cite_style = self.doc.Styles("cite_bib")
        except Exception:
            return

        rng = self.doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Style = cite_style
        rng.Find.Text = ""

        while rng.Find.Execute():
            text = rng.Text.strip()
            if not text:
                rng.Collapse(0)
                continue
            
            nums = self._extract_numbers(text)
            if not nums:
                rng.Collapse(0)
                continue

            # Map numbers and create new display text
            new_display = re.sub(
                r'\d+(-\d+)?', 
                lambda m: self._mapped_segment(m.group(0), renumber_map), 
                text
            )
            
            if not re.search(r'\d', new_display):
                mapped = [renumber_map.get(n, n) for n in nums]
                new_display = self._numbers_to_string(mapped)

            try:
                rng.Text = new_display
            except Exception:
                pass
            rng.Collapse(0)

    def _reorder_references(self, renumber_map, ordered_all):
        """
        Physically reorder reference list paragraphs based on the new citation order.
        Strategy: Extract all REF-N paragraphs, map them to old numbers, then insert them
        back in the correct order.
        """
        try:
            refpara_style = self.doc.Styles("REF-N")
        except Exception:
            # If REF-N style not found, try updating bib_number style only
            self._update_bib_numbers_only(renumber_map)
            return

        try:
            bib_char_style = self.doc.Styles("bib_number")
        except Exception:
            bib_char_style = None

        # Step 1: Collect all reference paragraphs with their old numbers
        ref_paras = []  # List of dicts with old_num, text, para reference
        
        for para in self.doc.Paragraphs:
            try:
                if para.Range.Style == refpara_style or getattr(para.Range.Style, 'NameLocal', '') == 'REF-N':
                    # Extract the old reference number
                    txt = para.Range.Text
                    old_num = self._extract_first_number_from_para(para, bib_char_style)
                    
                    if old_num is not None:
                        ref_paras.append({
                            'old_num': old_num,
                            'text': txt,
                            'para': para
                        })
            except Exception:
                continue

        if not ref_paras:
            return

        # Step 2: Find the position where references start
        first_ref_para = ref_paras[0]['para']
        insert_position = first_ref_para.Range.Start

        # Step 3: Delete all existing reference paragraphs (in reverse to avoid index shifts)
        for ref_data in reversed(ref_paras):
            try:
                ref_data['para'].Range.Delete()
            except Exception:
                pass

        # Step 4: Create a mapping from old number to reference data
        ref_map = {item['old_num']: item for item in ref_paras}

        # Step 5: Insert references in new order
        insertion_range = self.doc.Range(insert_position, insert_position)
        
        for new_num, old_num in enumerate(ordered_all, start=1):
            if old_num not in ref_map:
                continue
            
            ref_data = ref_map[old_num]
            old_text = ref_data['text']
            
            # Update the reference number in the text
            new_text = self._update_reference_number(old_text, old_num, new_num)
            
            # Insert the paragraph
            try:
                insertion_range.InsertAfter(new_text)
                insertion_range.InsertParagraphAfter()
                
                # Apply REF-N style to the newly inserted paragraph
                new_para_range = self.doc.Range(
                    insertion_range.Start, 
                    insertion_range.End - 1  # Exclude the paragraph mark
                )
                new_para_range.Style = refpara_style
                
                # Update insertion point for next reference
                insertion_range = self.doc.Range(insertion_range.End, insertion_range.End)
            except Exception as e:
                print(f"Error inserting reference {new_num}: {e}")
                continue

    def _extract_first_number_from_para(self, para, bib_char_style=None):
        """Extract the first number from a paragraph, preferring bib_number style."""
        if bib_char_style:
            try:
                p_rng = para.Range
                p_rng.Find.ClearFormatting()
                p_rng.Find.Style = bib_char_style
                p_rng.Find.Text = ""
                if p_rng.Find.Execute():
                    bib_text = p_rng.Text.strip()
                    match = re.search(r'\d+', bib_text)
                    if match:
                        return int(match.group())
            except Exception:
                pass
        
        # Fallback: extract first number from text
        text = para.Range.Text.strip()
        match = re.search(r'\b\d+\b', text)
        return int(match.group()) if match else None

    def _update_reference_number(self, text, old_num, new_num):
        """Update the reference number in a reference paragraph text."""
        # Replace first occurrence of old number with new number
        pattern = r'\b' + str(old_num) + r'\b'
        new_text = re.sub(pattern, str(new_num), text, count=1)
        return new_text

    def _update_bib_numbers_only(self, renumber_map):
        """Fallback method to update bib_number style when REF-N is not available."""
        try:
            bib_style = self.doc.Styles("bib_number")
        except Exception:
            return
        
        rng = self.doc.Content
        rng.Find.ClearFormatting()
        rng.Find.Style = bib_style
        rng.Find.Text = ""
        
        while rng.Find.Execute():
            txt = rng.Text.strip()
            if txt:
                match = re.search(r'\d+', txt)
                if match:
                    old_num = int(match.group())
                    new_num = renumber_map.get(old_num, old_num)
                    new_txt = re.sub(r'\d+', str(new_num), txt)
                    try:
                        rng.Text = new_txt
                    except Exception:
                        pass
            rng.Collapse(0)

    def _mapped_segment(self, seg, renumber_map):
        """Map a segment like '2' or '2-4' using renumber_map and return a string representation."""
        if '-' in seg:
            a, b = seg.split('-', 1)
            a_n = renumber_map.get(int(a), int(a))
            b_n = renumber_map.get(int(b), int(b))
            # If mapped numbers are contiguous, show as range
            if b_n - a_n >= 1 and self._is_contiguous_mapping(int(a), int(b), renumber_map):
                return f"{a_n}-{b_n}"
            else:
                # Return comma-separated mapped numbers
                return ','.join(str(renumber_map.get(int(x), int(x))) for x in range(int(a), int(b)+1))
        else:
            n = int(seg)
            return str(renumber_map.get(n, n))

    def _is_contiguous_mapping(self, a, b, renumber_map):
        vals = [renumber_map.get(i, i) for i in range(a, b+1)]
        return vals == list(range(min(vals), max(vals)+1))

    def _numbers_to_string(self, numbers):
        """Convert list of integers into a compact string like '1,3-5,7'."""
        if not numbers:
            return ''
        nums = sorted(numbers)
        ranges = []
        start = prev = nums[0]
        for n in nums[1:]:
            if n == prev + 1:
                prev = n
                continue
            else:
                if start == prev:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{prev}")
                start = prev = n
        # finalize
        if start == prev:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{prev}")
        return ','.join(ranges)

    def _get_reference_numbers(self):
        numbers = set()

        # Strategy: Only extract the reference LIST number (e.g., "1", "2", "3"), 
        # NOT all numbers in the reference text (years, pages, volumes, etc.)
        
        # First try to find reference paragraphs styled as 'REF-N' (paragraph style for the whole ref)
        try:
            refpara_style = self.doc.Styles("REF-N")
        except Exception:
            refpara_style = None

        # Also get bib_number style for extracting just the number
        try:
            bib_char_style = self.doc.Styles("bib_number")
        except Exception:
            bib_char_style = None

        if refpara_style:
            for para in self.doc.Paragraphs:
                try:
                    if para.Range.Style == refpara_style or getattr(para.Range.Style, 'NameLocal', '') == 'REF-N':
                        # Look for bib_number style within this paragraph
                        found_number = False
                        if bib_char_style:
                            # Search for bib_number style in this paragraph
                            p_rng = para.Range
                            p_rng.Find.ClearFormatting()
                            p_rng.Find.Style = bib_char_style
                            p_rng.Find.Text = ""
                            if p_rng.Find.Execute():
                                # Extract only the first number from the bib_number styled text
                                bib_text = p_rng.Text.strip()
                                # Extract just the first number (ignore ranges, just get the leading number)
                                match = re.search(r'\d+', bib_text)
                                if match:
                                    numbers.add(int(match.group()))
                                    found_number = True
                        
                        # Fallback: if no bib_number found, extract first number from paragraph
                        if not found_number:
                            text = para.Range.Text.strip()
                            match = re.search(r'\d+', text)
                            if match:
                                numbers.add(int(match.group()))
                except Exception:
                    continue

        # Also check for standalone bib_number occurrences (if not already found via REF-N)
        # This handles cases where bib_number might be used without REF-N paragraph style
        if bib_char_style and not refpara_style:
            for item in self._find_style_ranges("bib_number"):
                bib_text = item['text'].strip()
                match = re.search(r'\d+', bib_text)
                if match:
                    numbers.add(int(match.group()))

        return numbers

    def _get_citations(self):
        # Find citations that use 'cite_bib' character style anywhere in the text
        try:
            _ = self.doc.Styles("cite_bib")
        except Exception:
            raise ValueError("'cite_bib' style not found")

        citations = []
        for item in self._find_style_ranges("cite_bib"):
            text = item['text'].strip()
            if text:
                numbers = self._extract_numbers(text)
                if numbers:
                    citations.append({
                        'text': text,
                        'numbers': numbers,
                        'range_start': item.get('range_start'),
                        'range_end': item.get('range_end')
                    })

        return citations

    def _find_style_ranges(self, style_name):
        """Return list of dicts {'text','range_start','range_end'} for occurrences of a style.
        Supports both paragraph and character styles by running a paragraph scan and a Find on the document content.
        """
        results = []
        # Try to get the style object (may be paragraph or character style)
        try:
            style_obj = self.doc.Styles(style_name)
        except Exception:
            style_obj = None

        # 1) Paragraph-level scan: find paragraphs whose paragraph style matches style_name
        if style_obj:
            try:
                for para in self.doc.Paragraphs:
                    try:
                        if para.Range.Style == style_obj or getattr(para.Range.Style, 'NameLocal', '') == style_name:
                            results.append({'text': para.Range.Text, 'range_start': para.Range.Start, 'range_end': para.Range.End})
                    except Exception:
                        continue
            except Exception:
                pass

        # 2) Character-style Find across the document content (this will also find character style runs)
        try:
            rng = self.doc.Content
            rng.Find.ClearFormatting()
            if style_obj:
                rng.Find.Style = style_obj
            rng.Find.Text = ""
            rng.Find.Format = True
            # Execute Find: when Style is a character style this finds character runs; for paragraph styles it may re-find paragraphs too
            while rng.Find.Execute():
                results.append({'text': rng.Text, 'range_start': rng.Start, 'range_end': rng.End})
                rng.Collapse(0)
        except Exception:
            pass

        return results

    def _extract_numbers(self, text):
        numbers = []

        # Handle ranges first
        for match in re.finditer(r'(\d+)-(\d+)', text):
            start, end = int(match.group(1)), int(match.group(2))
            
            # Safety check for massive ranges (e.g. typos like 1-1000000 or phone numbers)
            if end - start > 999:
                # If range is too large, treat as separate numbers to avoid memory issues
                numbers.append(start)
                numbers.append(end)
            elif end >= start:
                numbers.extend(range(start, end + 1))

        # Get individual numbers (excluding those that were part of ranges)
        text_no_ranges = re.sub(r'\d+-\d+', '', text)
        numbers.extend([int(match.group()) for match in re.finditer(r'\b\d+\b', text_no_ranges)])

        return numbers

    def _check_citation_sequence(self):
        sequence = self.results['citation_sequence']
        if len(sequence) < 2:
            self.results['sequence_message'] = "Citations are in proper sequence."
            return

        seen = set()
        unique_sequence = []
        for num in sequence:
            if num not in seen:
                unique_sequence.append(num)
                seen.add(num)

        is_ordered = unique_sequence == sorted(unique_sequence)

        if not is_ordered:
            self.results['sequence_issues'].append(sequence)
            self.results['sequence_message'] = "Citations are NOT in sequence."
        else:
            self.results['sequence_message'] = "Citations are in proper sequence."