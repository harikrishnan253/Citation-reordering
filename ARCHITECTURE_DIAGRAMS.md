# Physical Reference Reordering: Visual Flow

## Architecture Diagram

```
┌─────────────────────────────────────────────────────────────────────┐
│                    Word Document (Input)                            │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  BODY TEXT:                                                         │
│  "Recent research [5] demonstrates..."                              │
│  "Earlier work [2] suggested..."                                    │
│  "Current studies [3] indicate..."                                 │
│                                                                     │
│  REFERENCES:                                                        │
│  [1] Author A. Paper A.                                             │
│  [2] Author B. Paper B.                                             │
│  [3] Author C. Paper C.                                             │
│  [4] Author D. Paper D.                                             │
│  [5] Author E. Paper E.                                             │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│              ReferenceValidator.validate()                          │
│              (with auto_renumber=True)                              │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  1. Extract citations: [5, 2, 3]                                   │
│  2. Check sequence: NOT in sequence                                │
│  3. Build renumber map:                                            │
│     {5→1, 2→2, 3→3, unused→4,5}                                    │
│  4. Reopen document writable                                       │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│       _extract_references_with_numbers()                            │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Scan all REF-N paragraphs and extract:                            │
│                                                                     │
│  {                                                                  │
│    1: {text: "[1] Author A...", paragraphs: [...]},              │
│    2: {text: "[2] Author B...", paragraphs: [...]},              │
│    3: {text: "[3] Author C...", paragraphs: [...]},              │
│    4: {text: "[4] Author D...", paragraphs: [...]},              │
│    5: {text: "[5] Author E...", paragraphs: [...]},              │
│  }                                                                  │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│  Update Citations in Body Text                                      │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Find all cite_bib styled text:                                    │
│  [5] → [1]                                                          │
│  [2] → [2]                                                          │
│  [3] → [3]                                                          │
│                                                                     │
│  Result:                                                            │
│  "Recent research [1] demonstrates..."  ← Updated                 │
│  "Earlier work [2] suggested..."        ← Unchanged               │
│  "Current studies [3] indicate..."      ← Unchanged               │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│       _reorder_references_physically()   ⭐ NEW ⭐                  │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Step 1: Delete ALL existing REF-N paragraphs                      │
│  ┌─────────────────────────────┐                                   │
│  │ Doc now has NO references   │                                   │
│  └─────────────────────────────┘                                   │
│                                                                     │
│  Step 2: Re-insert in citation order                               │
│  ┌────────────────────────────────────────────────────┐           │
│  │ Position 1: [1] Author E. Paper E.  (was Ref 5)   │           │
│  │ Position 2: [2] Author B. Paper B.  (was Ref 2)   │           │
│  │ Position 3: [3] Author C. Paper C.  (was Ref 3)   │           │
│  │ Position 4: [4] Author A. Paper A.  (was Ref 1)   │           │
│  │ Position 5: [5] Author D. Paper D.  (was Ref 4)   │           │
│  └────────────────────────────────────────────────────┘           │
│                                                                     │
│  Uses Temp Document for safe clipboard operations                  │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│                   Save Document                                     │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  Document.Save() or Document.SaveAs(output_path)                   │
│                                                                     │
│  Return: {renumbered: True, map: {5→1, 2→2, 3→3, ...}}           │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
                              │
                              ↓
┌─────────────────────────────────────────────────────────────────────┐
│                  Word Document (Output)                             │
├─────────────────────────────────────────────────────────────────────┤
│                                                                     │
│  BODY TEXT:                                                         │
│  "Recent research [1] demonstrates..."     ← Updated              │
│  "Earlier work [2] suggested..."           ← Updated              │
│  "Current studies [3] indicate..."         ← Updated              │
│                                                                     │
│  REFERENCES:                                                        │
│  [1] Author E. Paper E.          ← Moved from position 5          │
│  [2] Author B. Paper B.          ← Unchanged                       │
│  [3] Author C. Paper C.          ← Unchanged                       │
│  [4] Author A. Paper A.          ← Moved from position 1          │
│  [5] Author D. Paper D.          ← Moved from position 4          │
│                                                                     │
│  ✓ Citations & References ALIGNED                                  │
│  ✓ Citation [1] → Reference [1] (Author E)                        │
│  ✓ All numbers synchronized                                        │
│                                                                     │
└─────────────────────────────────────────────────────────────────────┘
```

## Data Flow: References Dictionary

```
Step A: EXTRACT (before modifications)
┌─────────────────────────────────────────┐
│ references_dict = {                     │
│   1: {                                  │
│     'text': '[1] Author A. Paper A',   │
│     'paragraphs': [Paragraph(1)]        │
│   },                                    │
│   2: {                                  │
│     'text': '[2] Author B. Paper B',   │
│     'paragraphs': [Paragraph(2)]        │
│   },                                    │
│   ...                                   │
│   5: {                                  │
│     'text': '[5] Author E. Paper E',   │
│     'paragraphs': [Paragraph(5)]        │
│   }                                     │
│ }                                       │
└─────────────────────────────────────────┘

Step B: CREATE NEW ORDER MAP
┌──────────────────────────────────────────┐
│ new_order_map = {                        │
│   1: 5,  # Position 1 gets old Ref 5    │
│   2: 2,  # Position 2 gets old Ref 2    │
│   3: 3,  # Position 3 gets old Ref 3    │
│   4: 1,  # Position 4 gets old Ref 1    │
│   5: 4   # Position 5 gets old Ref 4    │
│ }                                        │
└──────────────────────────────────────────┘

Step C: REORDER & UPDATE NUMBERS
┌──────────────────────────────────────────┐
│ For each new_position in [1,2,3,4,5]:   │
│   old_ref = new_order_map[new_position]  │
│   content = references_dict[old_ref]     │
│   text = content['text']                 │
│   # Replace "[5] " with "[1] "          │
│   new_text = re.sub('^\d+', '1', text)  │
│   # Insert at new position               │
│   doc.InsertParagraph(new_text)          │
└──────────────────────────────────────────┘
```

## Method Call Sequence

```
validate(auto_renumber=True, save_path="output.docx")
  ├─ _get_reference_numbers()          → Get all ref numbers
  ├─ _get_citations()                   → Get all citations
  ├─ _check_citation_sequence()         → Check if ordered
  │
  └─ renumber_if_needed(save_path)      → If sequence issues found
      ├─ _extract_references_with_numbers()    → Capture refs BEFORE changes
      │  ├─ For each paragraph:
      │  │  ├─ Check if REF-N style
      │  │  └─ _extract_ref_number_from_para()
      │  │     ├─ Look for bib_number style
      │  │     └─ Extract number with regex
      │
      ├─ Update citations in body text (cite_bib)
      │  ├─ Find all cite_bib styled runs
      │  ├─ _extract_numbers()           → Parse citation numbers
      │  ├─ _mapped_segment()            → Remap each segment
      │  └─ Replace with new numbers
      │
      ├─ _reorder_references_physically()       → ⭐ NEW ⭐
      │  ├─ Build new_order_map (new_pos → old_ref)
      │  ├─ Find first REF-N paragraph
      │  ├─ Delete all REF-N paragraphs (reverse order)
      │  ├─ Create temp_doc
      │  ├─ For each position in sorted order:
      │  │  ├─ Get old reference content
      │  │  ├─ Update number prefix
      │  │  ├─ Copy to temp_doc
      │  │  ├─ Paste to main doc
      │  │  └─ Move insertion point
      │  └─ Close temp_doc
      │
      ├─ Update reference prefixes (bib_number)
      │  ├─ For each REF-N paragraph:
      │  │  ├─ Find bib_number style
      │  │  └─ Replace number sequentially
      │
      └─ Save document
         ├─ document.SaveAs(output_path)    if save_path
         └─ document.Save()                 otherwise
```

## State Changes During Processing

```
INITIAL STATE
──────────────────────────────
Citations in order: [5, 2, 3]
 References in order: [1, 2, 3, 4, 5]
 Status: ❌ OUT OF SEQUENCE

AFTER Citation Update
──────────────────────────────
Citations in order: [1, 2, 3]
References in order: [1, 2, 3, 4, 5] (unchanged)
Status: ⚠️  Numbers mismatch - Ref 1 not cited first

AFTER Physical Reordering
──────────────────────────────
Citations in order: [1, 2, 3]
References in order: [5, 2, 3, 1, 4] → [1, 2, 3, 4, 5] (REORDERED)
Status: ✅ ALIGNED

FINAL STATE (Saved)
──────────────────────────────
Citations: [1, 2, 3]
References: [1] Author E, [2] Author B, [3] Author C, [4] Author A, [5] Author D
Status: ✅ Perfect Alignment
```

## Error Handling Flow

```
┌─ ReferenceValidator(filepath)
│
├─ __enter__: Open document Read-Only
│  ├─ pythoncom.CoInitialize()
│  ├─ Create Word.Application
│  └─ Open document
│
├─ validate()
│  └─ If sequence issues + auto_renumber:
│
├─ renumber_if_needed()
│  ├─ Reopen document Read-Write  ← Critical point
│  │  ├─ If fails: Try without ReadOnly parameter
│  │  └─ If fails again: Return error
│  │
│  ├─ Try extract references
│  │  └─ Exception: Log, return empty dict (graceful)
│  │
│  ├─ Try update citations
│  │  └─ Exception: Log, continue (other updates might work)
│  │
│  ├─ Try physical reordering
│  │  ├─ If no REF-N style found: Skip gracefully
│  │  ├─ If deletion fails: Continue (might have refs left)
│  │  └─ Ensure temp_doc always closes (finally block)
│  │
│  └─ Try save
│     ├─ Try .SaveAs() or .Save()
│     ├─ If fails: Try .SaveAs2() (Word 2013+)
│     └─ If fails: Return error message
│
└─ __exit__: Close document and cleanup
   ├─ doc.Close(SaveChanges=False)
   ├─ word.Quit()
   └─ pythoncom.CoUninitialize()
```

## Performance Timeline

```
For typical 50-ref scientific paper:

Step                        Time        Notes
─────────────────────────────────────────────────────
Open document              ~0.5 sec
Extract references         ~0.3 sec    Scan all paragraphs
Update citations           ~0.5 sec    Find/replace in text
Delete refs                ~0.3 sec    Remove ~50 paragraphs
Re-insert refs             ~2.0 sec    Clipboard operations
Update prefixes            ~0.3 sec    Sequential renumbering
Save document              ~0.5 sec
─────────────────────────────────────────────────────
TOTAL                      ~4.4 sec

Larger documents scale linearly.
Clipboard operations are the bottleneck.
```

## Key Success Indicators

```
✅ References extracted correctly
   └─ Length of references_dict == total references

✅ Physical reordering executed
   └─ _reorder_references_physically() returns True

✅ Citations & references aligned
   └─ After validation: citation_sequence == [1, 2, 3, ...]

✅ Document saved
   └─ renumber_attempt['renumbered'] == True
   └─ Output file exists and opens without errors

✅ No data loss
   └─ Reference content preserved (text matches original)
   └─ Only numbers changed
```
