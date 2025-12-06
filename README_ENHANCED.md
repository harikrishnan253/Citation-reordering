# Citation-Reordering: Physical Reference Reordering Enhancement

## Overview

This project validates and auto-corrects citation sequences in Microsoft Word DOCX documents. The **enhanced version** includes **physical reordering of reference paragraphs** to match citation appearance order in the document body.

## Problem Statement

**Before Enhancement**: When citations appear out of order (e.g., [5, 2, 3]):
- Only citation numbers were remapped
- Reference prefixes were updated ("5." → "1.", "2." → "2.", etc.)
- But reference paragraphs remained in original positions
- Result: Misalignment between citation order and reference content

**After Enhancement**: 
- Reference paragraphs are **physically moved** to match citation order
- All numbers (citations and references) are synchronized
- Document follows academic standards: citations point to references in order

## Key Features

### 1. Comprehensive Validation
- ✅ Detects out-of-sequence citations
- ✅ Finds missing references (cited but not defined)
- ✅ Identifies unused references (defined but not cited)
- ✅ Reports citation appearance sequence

### 2. Automatic Renumbering with Physical Reordering
- ✅ Maps old citation numbers to new (based on appearance order)
- ✅ Updates all cite_bib styled citations in text
- ✅ **Physically deletes and re-inserts reference paragraphs** in new order
- ✅ Updates bib_number prefixes sequentially
- ✅ Saves corrected document (overwrites or custom path)

### 3. Style-Aware Processing
- Recognizes Word paragraph styles: `REF-N` (reference paragraph)
- Recognizes character styles: `cite_bib` (citation), `bib_number` (reference number)
- Handles range citations: "1-5" → properly renumbered as "1-3", etc.

## File Structure

```
Citation-reordering/
├── validator_enhanced.py           # Enhanced validator with physical reordering
├── validator.py                    # Original validator (legacy)
├── manual_validate_enhanced.py     # Updated demo script for enhanced validator
├── manual_validate.py              # Original demo script (legacy)
├── requirements.txt                # Python dependencies
└── S4C-Processed-Documents/        # Sample documents
    ├── Abuhamad9781975242831-ch002.docx              # Original
    └── Abuhamad9781975242831-ch002-renumbered.docx  # Output
```

## Installation & Setup

### Windows Only (requires pywin32)

```bash
# Clone repository
git clone https://github.com/harikrishnan253/Citation-reordering.git
cd Citation-reordering

# Install dependencies
pip install -r requirements.txt
pip install pywin32

# Register COM interface (one-time setup)
python -m pywin32_postinstall -install
```

## Usage

### Quick Start

```python
from validator_enhanced import ReferenceValidator

file_path = "path/to/document.docx"
output_path = "path/to/document-renumbered.docx"

with ReferenceValidator(file_path) as validator:
    results = validator.validate(auto_renumber=True, save_path=output_path)

print(f"Renumbered: {results['renumber_attempt']['renumbered']}")
print(f"Map: {results['renumber_attempt']['map']}")
```

### Command-Line Demo

```bash
python manual_validate_enhanced.py
```

**Sample Output**:
```
============================================================
VALIDATION REPORT
============================================================
Total References: 42
Total Citations: 38

Sequence Status: Citations are NOT in sequence.

Missing References: []
Unused References: [2, 15, 41]

Renumber Status: True

✓ References have been PHYSICALLY REORDERED to match citation order
✓ Renumbered file saved to: .../Abuhamad9781975242831-ch002-renumbered.docx

Renumber Map (Old → New):
  5 → 1
  2 → 2
  3 → 3
  ...
============================================================
```

## How Physical Reordering Works

### Step 1: Extract References
```python
references_dict = validator._extract_references_with_numbers()
# Returns: {1: {text: "...", paragraphs: [...]}, ...}
```

### Step 2: Build Renumber Map
```python
# Citation sequence: [5, 2, 3, ...]
# Renumber map: {5→1, 2→2, 3→3, ...}
```

### Step 3: Delete & Re-insert
1. Delete all existing REF-N paragraphs
2. For each position (1, 2, 3, ...):
   - Get old reference content (e.g., Ref 5 → becomes position 1)
   - Update number prefix ("5. Author..." → "1. Author...")
   - Copy to temp document
   - Paste at insertion point
3. Save with all references in correct order

## Core Methods

### `_extract_references_with_numbers()`
Maps old reference numbers to their paragraph content before any modifications.

**Returns**: `{old_num: {text: str, paragraphs: [Range]}}`

### `_extract_ref_number_from_para(para)`
Extracts reference number from a single paragraph's bib_number style.

### `_reorder_references_physically(renumber_map, references_dict)`
Physically deletes and re-inserts reference paragraphs in citation order.

**Logic**:
1. Collect all REF-N paragraph indices
2. Delete in reverse order (prevents index shift)
3. Re-insert from temp document in new order
4. Returns: `True` if successful

### `renumber_if_needed(save_path=None)`
Main orchestration method (enhanced to include physical reordering):
1. Validate citation sequence
2. Build renumber map
3. Update citations in text
4. **Call `_reorder_references_physically()`** ← NEW
5. Update reference prefixes
6. Save document

## Example Transformation

### Before Renumbering
```
BODY:
"...recent research [5] shows improvement [2]. Earlier studies [3] suggest..."

REFERENCES:
[1] Author A. Paper A. 2020.
[2] Author B. Paper B. 2021.
[3] Author C. Paper C. 2022.
[4] Author D. Paper D. 2023.
[5] Author E. Paper E. 2024.
```

### After Physical Reordering
```
BODY:
"...recent research [1] shows improvement [2]. Earlier studies [3] suggest..."

REFERENCES:
[1] Author E. Paper E. 2024.        ← Was Ref 5, now Ref 1 (physically moved)
[2] Author B. Paper B. 2021.        ← Already Ref 2 (no physical move)
[3] Author C. Paper C. 2022.        ← Already Ref 3 (no physical move)
[4] Author D. Paper D. 2023.        ← Unused
[5] Author A. Paper A. 2020.        ← Moved to end
```

## Validation Results Structure

```python
results = {
    'total_references': int,
    'total_citations': int,
    'citation_sequence': [int],                    # Order of appearance
    'missing_references': set,                    # Cited but not defined
    'unused_references': set,                     # Defined but not cited
    'sequence_message': str,                      # 'IN sequence' or 'NOT in sequence'
    'sequence_issues': list,                      # Details if out of order
    'renumber_attempt': {
        'renumbered': bool,
        'map': {old_num: new_num},               # Renumber mapping
        'message': str                            # Status message
    }
}
```

## Limitations & Notes

### Current Limitations
- **Windows-only**: Requires Microsoft Word COM API (pywin32)
- **One reference per number**: Assumes 1-to-1 mapping (standard scientific writing)
- **Clipboard-based insertion**: Uses temp documents; may fail in restricted environments
- **No rollback**: Overwrites original unless save_path specified

### Handling Edge Cases
- **Multi-paragraph references**: Grouped by first paragraph's bib_number
- **References without numbers**: Skipped (warns in logs)
- **Massive ranges** (e.g., "1-9999"): Treated as separate numbers (safety check)
- **Duplicate citations**: All instances remapped

## Cross-Platform Roadmap (python-docx)

For Linux/macOS/cloud deployment:

```python
# Proposed future enhancement
from docx import Document
from copy import deepcopy

def _reorder_references_physically_docx(self, renumber_map, references_dict):
    doc = Document(self.filepath)
    parent = doc._element.body
    
    # Delete existing REF-N paragraphs
    # Re-insert via XML element manipulation
    # No clipboard needed
```

## Testing

Create unit tests in `tests/test_reordering.py`:

```python
def test_citation_reference_alignment():
    """Test physical reordering aligns citations with references."""
    with ReferenceValidator(test_docx) as validator:
        results = validator.validate(auto_renumber=True)
        
        # Verify citation sequence matches new reference order
        assert results['citation_sequence'] == [1, 2, 3]
        
        # Verify references are in correct order
        ref_numbers = validator._get_reference_numbers()
        assert list(ref_numbers) == [1, 2, 3]
```

Run with:
```bash
pytest tests/ -v
```

## S4C AI e-Publishing Integration

For the S4C platform:

1. **Batch Processing**: Process multiple chapters in sequence
2. **Metrics Dashboard**: Track fix rates, missing refs per book
3. **API Endpoint**: `/validate-docx` endpoint for web uploads
4. **Database Logging**: Store results in PostgreSQL for quality audits

## Future Enhancements

- [ ] Cross-platform support (python-docx for Linux/AWS Lambda)
- [ ] CLI with `--batch-dir`, `--dry-run`, `--report-json` flags
- [ ] Web API (FastAPI) for cloud deployment
- [ ] HTML/PDF reports with visualization
- [ ] Support for other citation styles (IEEE, APA, Chicago)
- [ ] Integration with Zotero/Mendeley metadata
- [ ] Git-based diff viewer for before/after

## Contributing

Contributions welcome! Please:
1. Test with multiple document styles
2. Document any style assumptions
3. Add unit tests for new features
4. Update requirements.txt if adding dependencies

## License

Internal S4C project. See LICENSE file.

## Support

For issues or questions:
- Check existing GitHub issues
- Review validator_enhanced.py code comments
- Test with sample documents in S4C-Processed-Documents/
