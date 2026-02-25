---
name: docx-essay-writer
description: |
  Use when user provides a .docx template together with specific writing requirements,
  and needs to generate a new document that follows the same writing structure, formatting,
  and layout as the template. Covers the full workflow: template analysis → Markdown generation
  → user review → format-preserving DOCX conversion.
---

# DOCX Essay Writer

Generate a new Word document from a reference .docx template: analyze the template's writing structure, produce Markdown content matching user requirements, and convert to a format-identical DOCX.

## Overview

This skill provides an end-to-end workflow:

1. **Template Analysis** — extract heading hierarchy, styles, page layout, and special elements
2. **Content Generation** — AI writes Markdown following the template's outline
3. **User Review** — iterate until the user confirms
4. **DOCX Conversion** — convert the confirmed Markdown back to DOCX with 100% format fidelity

All scripts are self-contained under `skills/docx-essay-writer/scripts/`.

## IMPORTANT RULES

### Rule 0: References Are Sacred (参考文献神圣不可侵犯)

**ABSOLUTELY PROHIBITED** under any circumstances:
- Never fabricate any bibliographic information
- Never guess or "reasonably infer" reference details
- Never use format examples as real references
- Never duplicate existing references with modified details

**MANDATORY**:
- Only present verified data from OpenAlex API responses
- Report search results accurately; clearly state when results are insufficient
- Direct users to other databases (CNKI, Wanfang, Google Scholar) when needed

### Rule 1: Analyze Before Writing

Always run `analyze_template.py` first. Never start writing content without understanding the template structure.

### Rule 2: Match the Outline

The Markdown's heading hierarchy must correspond 1-to-1 with the template's outline. Do not add or remove sections unless the user explicitly requests it.

### Rule 3: Use Built-in Reference Search

When references are needed, use `scripts/search_references.py` to search via OpenAlex. Never write reference entries by hand.

## Workflow

### Phase 1: Analyze Template

```bash
python3 skills/docx-essay-writer/scripts/analyze_template.py <template.docx>
```

Output is JSON containing:
- `outline` — heading hierarchy with section titles and special elements
- `style_map` — all style names → styleId mappings
- `page_layout` — margins, page size, header/footer references
- `special_elements` — counts of tables, formulas, images, references

Review the outline carefully before proceeding.

### Phase 2: Generate Markdown

Based on the analysis output, write Markdown content that:

- Follows the **exact same heading hierarchy** as the template
- Uses `# Heading 1`, `## Heading 2`, `### Heading 3` for section titles
- Uses **standard Markdown** for all elements:

| Template Element | Markdown Syntax |
|-----------------|-----------------|
| Body text | Plain paragraphs |
| Bold | `**bold text**` |
| Italic | `*italic text*` |
| Inline math | `$formula$` |
| Display math | `$$formula$$` |
| Table | `\| col1 \| col2 \|` |
| Image | `![caption](path/to/image.png)` |
| Unordered list | `- item` |
| Ordered list | `1. item` |
| Reference | `[N] Author. Title[J]. Journal, Year.` |

For **formulas**, use LaTeX syntax. Supported constructs:
- Greek letters: `\alpha`, `\omega`, etc.
- Subscripts/superscripts: `x_d`, `x^2`, `x_d^*`
- Fractions: `\frac{a}{b}`
- Matrices: `\begin{bmatrix}...\end{bmatrix}`
- Cases: `\begin{cases}...\end{cases}`
- Functions: `\cos()`, `\sin()`, etc.

For **images**, generate them using available tools (matplotlib, mermaid, etc.) and reference the saved file path.

For **references**, search using the built-in script:

```bash
python3 skills/docx-essay-writer/scripts/search_references.py \
  --query "search topic" \
  --min-year 2015 \
  --min-results 5 \
  --language en
```

Or use the Python API:

```python
import sys
sys.path.insert(0, 'skills/docx-essay-writer/scripts')
from search_references import ReferenceSearcher

searcher = ReferenceSearcher(api_key="YOUR_KEY")  # optional
refs = searcher.search("ESP32 FOC motor control", min_year=2015, min_results=5)
for ref in refs:
    print(ref)
```

### Phase 3: User Review

Present the complete Markdown to the user. Wait for confirmation or modification requests. Iterate until the user is satisfied.

### Phase 4: Convert to DOCX

```bash
python3 skills/docx-essay-writer/scripts/md_to_docx.py \
  --template <template.docx> \
  --markdown <content.md> \
  --output <output.docx>
```

Or via Python:

```python
sys.path.insert(0, 'skills/docx-essay-writer/scripts')
from md_to_docx import MarkdownToDocxConverter

converter = MarkdownToDocxConverter("template.docx")
with open("content.md", "r", encoding="utf-8") as f:
    md = f.read()
converter.convert(md, "output.docx")
```

### Phase 4.5: Verify Formatting

```bash
python3 skills/docx-essay-writer/scripts/verify_format.py <template.docx> <output.docx>
```

Checks: page margins, page size, headers/footers, styles, spacing.

## Supported Markdown Elements

| Markdown | DOCX Output |
|----------|-------------|
| `# Heading 1` | Heading 1 style (dynamic from template) |
| `## Heading 2` | Heading 2 style (dynamic from template) |
| `### Heading 3` | Heading 3 style (dynamic from template) |
| `**bold**` | Bold formatting |
| `*italic*` | Italic formatting |
| `$...$` | Inline OMML formula |
| `$$...$$` | Display OMML formula paragraph |
| `\| table \|` | Table (border style from template, e.g. three-line) |
| `![cap](path)` | Embedded image |
| `- item` | Bullet list item |
| `1. item` | Numbered list item |
| `[N] Author...` | Reference paragraph |

## Script Reference

### `scripts/analyze_template.py`

- `TemplateAnalyzer(docx_path)` — main class
  - `analyze()` → full JSON analysis
  - `extract_outline()` → heading hierarchy
  - `extract_style_map()` → style name/ID map
  - `extract_page_layout()` → margins, size, header/footer refs
  - `extract_special_elements()` → table/formula/image/reference counts

### `scripts/search_references.py`

- `ReferenceSearcher(api_key=None)` — main class
  - `search(query, min_year, max_year, min_results, language, min_citations, open_access)` → list of formatted refs
  - `search_multi(topics, ...)` → merged deduplicated refs

### `scripts/md_to_docx.py`

- `MarkdownToDocxConverter(template_path)` — main class
  - `convert(markdown_content, output_path)` — full conversion

### `scripts/verify_format.py`

- `verify_docx(template_path, output_path)` → `(passed, messages)`

## Example Usage

```
参考 test_cases/Southwest Jiaotong University Thesis Template/正文.docx 的格式，
编写一篇关于"基于ESP32的无感FOC驱动设计"的绪论，要求有参考文献、公式、表格。
```

The AI will:
1. Run `analyze_template.py` on `正文.docx`
2. Generate Markdown with matching structure, including LaTeX formulas, Markdown tables, and real references from OpenAlex
3. Present for review
4. Convert to DOCX and verify formatting

## Troubleshooting

### Formulas not displaying
- Ensure OMML namespace prefixes are `w:`, `m:`, not `ns0:`, `ns1:`
- Every `<m:r>` must have `<m:rPr><m:sty m:val="p"/></m:rPr>`

### Styles not matching
- Run `analyze_template.py` to check the template's actual style IDs
- The converter dynamically extracts styles — check stderr for warnings

### Images not showing
- Ensure Pillow is installed: `pip install Pillow`
- Image path in Markdown must be accessible from the working directory

### References empty
- OpenAlex may require an API key (free from openalex.org)
- Try broader search terms or remove language/year filters
