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
  - User needs to convert Markdown to DOCX following a template format
---

# DOCX Template Text Replacer

Generate new Word documents by replacing text content in an existing template while preserving 100% of the original formatting.

## Overview

This skill provides a workflow and script for creating new Word documents based on templates. Unlike generating documents from scratch (which often loses subtle formatting), this approach:

1. Unpacks the template DOCX file (which is a ZIP archive)
2. Replaces specified text content in the XML
3. Repacks into a new DOCX file

This ensures **all formatting is preserved**: page margins, paragraph spacing, line height, headers/footers, fonts, underlines, borders, etc.

---

# Markdown to DOCX Converter

Convert Markdown format thesis content to Word documents that exactly match the Southwest Jiaotong University thesis template format.

## Overview

This converter transforms Markdown text into DOCX format while preserving all template formatting:

- Extracts template structure and styles
- Converts Markdown elements to corresponding Word XML
- Properly handles LaTeX math formulas (inline `$...$` and block `$$...$$`)
- Generates OMML (Office Math Markup Language) for formulas

## Workflow

### Step 1: Prepare Files

Ensure you have:
- Template DOCX file (e.g., `正文.docx`)
- Markdown content file (e.g., `正文内容.md`)

### Step 2: Execute Conversion

Use the provided converter script:

```bash
python skills/docx-template-replacer/scripts/md_to_docx_converter.py \
  --template <template.docx> \
  --markdown <content.md> \
  --output <output.docx>
```

Or use the Python API directly:

```python
from skills.docx-template-replacer.scripts.md_to_docx_converter import MarkdownToDocxConverter

converter = MarkdownToDocxConverter("template.docx")
with open("content.md", "r", encoding="utf-8") as f:
    markdown_content = f.read()
converter.convert(markdown_content, "output.docx")
```

### Step 3: Verify Output

Check that the output document:
- Opens correctly in Word/WPS
- Formulas display properly (especially matrices and cases)
- All heading styles match the template
- Page layout and margins are preserved

## Supported Markdown Elements

| Markdown | Output |
|----------|--------|
| `# Heading 1` | Heading 1 style |
| `## Heading 2` | Heading 2 style (with numbering) |
| `### Heading 3` | Heading 3 style (with numbering) |
| `**bold text**` | Bold formatting |
| `$...$` | Inline math formula |
| `$$...$$` | Display math formula |
| `\| table \|` | Table with borders |
| `` `code` `` | Code block |

## Formula Support

### LaTeX to OMML Conversion

The converter handles these LaTeX constructs:

- **Greek letters**: `\alpha`, `\beta`, `\omega`, etc. → α, β, ω
- **Subscripts**: `x_d`, `x_{abc}`, `i_\alpha` → x_d, x_abc, i_α
- **Superscripts**: `x^2`, `x^{n}`, `\omega^*` → x², xⁿ, ω*
- **Combined sub/sup**: `u_d^*`, `\omega_e^*`, `i_\alpha^*` → u_d^*, ω_e^*, i_α^*
- **Fractions**: `\frac{a}{b}` → a/b
- **Matrices**: `\begin{bmatrix}...\end{bmatrix}` → [matrix] with brackets
- **Cases**: `\begin{cases}...\end{cases}` → {cases with left brace
- **Functions**: `\sin()`, `\cos()`, `\tan()` → sin(), cos(), tan()

### Important OMML Structure Notes

Critical elements for formula display:

1. **m:sty element**: Every `<m:r>` must contain `<m:rPr><m:sty m:val="p"/></m:rPr>`
2. **Namespace prefixes**: Must use `w:`, `w14:`, `m:`, `r:` not `ns0:`, `ns1:` etc.
3. **Matrix brackets**: Matrices need `<m:d>` delimiter with `begChr="["` and `endChr="]"`
4. **Matrix columns**: Must include `<m:mcs>` with `<m:mc>` column definitions

## Example Usage

```python
# Example: Converting thesis content
converter = MarkdownToDocxConverter(
    "test_cases/Southwest Jiaotong University Thesis Template/正文.docx"
)

with open("test_cases/正文内容.md", "r", encoding="utf-8") as f:
    markdown_content = f.read()

converter.convert(markdown_content, "test_cases/output.docx")
```

## Script Reference

### `scripts/md_to_docx_converter.py`

Main conversion script with classes:

- `LatexToOmmlConverter` - Converts LaTeX math to OMML
  - `convert(latex)` - Convert LaTeX string to OMML Element
  - `_parse_expr(parent, expr)` - Recursive expression parser
  
- `MarkdownToDocxConverter` - Main converter class
  - `convert(markdown_content, output_path)` - Full conversion workflow
  - `_generate_document_xml(markdown)` - Generate Word XML from Markdown

### `scripts/replace_docx.py`

Template-based text replacement (original skill):

- `unpack_docx(docx_path, output_dir)` - Unpack DOCX to directory
- `pack_docx(input_dir, docx_path)` - Pack directory to DOCX
- `replace_in_docx(template_path, output_path, replacements, verify=True)` - Full workflow

### `scripts/verify_format.py`

Detailed format verification tool.

### `scripts/compare_docx.py`

File-by-file comparison tool.

## Testing

Test files are located in `assets/test_cases/`:
- `template.docx` - Sample template
- `正文内容.md` - Sample Markdown content

Run conversion test:
```bash
python scripts/md_to_docx_converter.py
```

Verify formulas in output:
```bash
# Check sSubSup elements (combined subscript/superscript)
unzip -p output.docx word/document.xml | grep -o 'sSubSup' | wc -l

# Check matrix delimiter brackets
unzip -p output.docx word/document.xml | grep -o 'begChr="\["' | wc -l
```

## Troubleshooting

### Formulas not displaying
- Check namespace prefixes are correct (w:, m:, not ns0:, ns1:)
- Verify `<m:sty m:val="p"/>` exists in all `<m:r>` elements

### Matrix brackets missing
- Ensure `<m:d>` wrapper with `begChr="["` and `endChr="]"` around matrix

### Greek letters not showing
- Verify Greek letter mapping in `greek_map` dictionary
- Check `eastAsia` hint for Greek characters

### Cases environment not displaying
- Cases converted to single delimiter with left brace `{`
- Each case row parsed separately
