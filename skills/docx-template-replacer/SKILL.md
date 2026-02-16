---
name: docx-template-replacer
description: |
  This skill should be used when the user needs to create a new Word document based on an existing template,
  while preserving all formatting details (page margins, paragraph spacing, headers/footers, fonts, etc.)
  and only replacing specific text content.
  
  Use this skill when:
  - User provides a .docx template and wants to replace text content while keeping exact formatting
  - User asks to "copy format" or "same format" for a new document
  - User wants to preserve page layout, headers, footers, line spacing, indentation
  - User needs batch text replacement in docx with format preservation
---

# DOCX Template Text Replacer

Generate new Word documents by replacing text content in an existing template while preserving 100% of the original formatting.

## Overview

This skill provides a workflow and script for creating new Word documents based on templates. Unlike generating documents from scratch (which often loses subtle formatting), this approach:

1. Unpacks the template DOCX file (which is a ZIP archive)
2. Replaces specified text content in the XML
3. Repacks into a new DOCX file

This ensures **all formatting is preserved**: page margins, paragraph spacing, line height, headers/footers, fonts, underlines, borders, etc.

## Workflow

### Step 1: Identify Replacements

Ask the user to provide:
- Path to the template DOCX file
- List of text replacements (original text → new text)
- Path for the output file

### Step 2: Execute Replacement

Use the provided script to perform the replacement:

```bash
python .codebuddy/skills/docx-template-replacer/scripts/replace_docx.py \
  <template.docx> \
  <output.docx> \
  <replacement1_old> <replacement1_new> \
  [<replacement2_old> <replacement2_new> ...]
```

Or use the Python API directly:

```python
from skills.docx-template-replacer.scripts.replace_docx import replace_in_docx

replacements = [
    ("old text 1", "new text 1"),
    ("old text 2", "new text 2"),
    # ... more pairs
]

replace_in_docx(
    template_path="template.docx",
    output_path="output.docx",
    replacements=replacements
)
```

### Step 3: Verify Output

The replacement script automatically verifies the output:
- ✓ Checks all files from template exist in output
- ✓ Verifies critical XML files (document.xml, headers, footers, styles, relationships)
- ✓ Confirms formatting elements (page margins, size, spacing, header/footer references)

Manual verification checklist:
- [ ] Document opens in Word/WPS without errors
- [ ] Page margins match template
- [ ] Headers/footers display correctly
- [ ] Paragraph spacing matches template
- [ ] All specified text was replaced

## Important Notes

### Text Matching

- Text must match **exactly** (including spaces, punctuation)
- The script automatically handles HTML entity encoding/decoding (e.g., `&#22522;&#20110;` → `基于`)
- Multi-paragraph text replacements are supported

### Formatting Preservation

The following are preserved automatically:
- Page margins (`<w:pgMar>`)
- Page size (`<w:pgSz>`)
- Paragraph spacing (`<w:spacing>`)
- Line height (`w:line`, `w:lineRule`)
- Indentation (`<w:ind>`)
- Alignment (`<w:jc>`)
- Fonts and sizes (`<w:rFonts>`, `<w:sz>`)
- Headers and footers
- Underlines, borders, colors

### Limitations

- Does NOT modify document structure (add/remove paragraphs)
- Does NOT handle tracked changes or comments
- Best for text content replacement only

## Example Usage

```python
# Example: Creating a new task assignment based on template
replacements = [
    ("基于三维激光雷达的智能移动充电桩的自动驾驶方法研究", 
     "基于ESP32的无感FOC驱动设计"),
    
    ("随着电动汽车和自动驾驶技术的广泛推广...",
     "随着电动汽车和工业自动化的快速发展..."),
     
    ("（1）系统性地调研和掌握自动驾驶底盘技术...",
     "（1）系统性地调研无感FOC控制技术..."),
]

replace_in_docx(
    "测试用例/模板/任务书.docx",
    "基于ESP32的无感FOC驱动设计_任务书.docx",
    replacements
)
```

## Script Reference

### `scripts/replace_docx.py`

Core replacement script. Functions:

- `unpack_docx(docx_path, output_dir)` - Unpack DOCX to directory
- `pack_docx(input_dir, docx_path)` - Pack directory to DOCX
- `replace_in_docx(template_path, output_path, replacements, verify=True)` - Full workflow with verification
- `verify_docx(template_path, output_path)` - Verify output preserves template structure

### `scripts/verify_format.py`

Detailed format verification tool. Checks:
- Page margins, size, orientation
- Header/footer references and content
- Paragraph spacing counts
- Critical file presence

```bash
python verify_format.py template.docx output.docx
```

Example output:
```
=== 格式元素检查 ===
✓ 页边距 (w:pgMar)
   模板: <w:pgMar w:top="1474" w:right="1531" ...>
   输出: <w:pgMar w:top="1474" w:right="1531" ...>
✓ 行距520段落数
   模板: 19
   输出: 19
```

### `scripts/compare_docx.py`

File-by-file comparison tool. Shows:
- Missing or extra files
- Files with different content

```bash
# Basic comparison
python compare_docx.py template.docx output.docx

# Show content differences
python compare_docx.py template.docx output.docx --show-content
```

## Testing

Test files are located in `assets/test_cases/`:
- `template.docx` - Sample template
- `expected_output.docx` - Expected result after replacement

Run tests with:
```bash
python scripts/test_replace.py
```
