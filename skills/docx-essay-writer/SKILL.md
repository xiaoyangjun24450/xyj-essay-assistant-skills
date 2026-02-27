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

### Rule 2: Match the Template Structure

- **Essay templates** (with heading hierarchy): Markdown heading levels must correspond 1-to-1 with the template's outline.
- **Form templates** (no heading hierarchy): Markdown section labels must match the template's numbered sections exactly. Do not repeat section header text as content.

Do not add or remove sections unless the user explicitly requests it.

### Rule 3: Use Built-in Reference Search

When references are needed, use `scripts/search_references.py` to search via OpenAlex. Never write reference entries by hand.

## Workflow

### Phase 1: Analyze Template

```bash
python3 skills/docx-essay-writer/scripts/analyze_template.py <template.docx>
```

Output is JSON containing:
- `document_type` — `"essay"` (heading-based), `"form"` (numbered sections), or `"complex"` (multi-section/cover pages)
- `outline` — heading hierarchy with section titles and special elements (essay type)
- `content_structure` — numbered sections with labels and sample content (form type)
- `style_map` — all style names → styleId mappings
- `page_layout` — margins, page size, header/footer references
- `special_elements` — counts of tables, formulas, images, references

**Check `document_type` first** — it determines which Phase 2 variant to use:
- `essay` → Use heading hierarchy (Phase 2A)
- `form` → Use numbered sections (Phase 2B)  
- `complex` → Use placeholder-based content (Phase 2C)

### Phase 2: Generate Markdown

#### Phase 2A: Essay Templates (`document_type: "essay"`)

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

#### Phase 2B: Form Templates (`document_type: "form"`)

The `content_structure` in the analysis output shows numbered sections with their labels and sample content. Write Markdown where:

- Each section uses `## N、section_label` as a delimiter (matching `content_structure.sections[].label`)
- Content under each section replaces the template's sample content
- Do **NOT** repeat the section header text or question text as content
- Do **NOT** use `---` horizontal rules or Markdown table syntax
- Write each sub-item (e.g. `（1）...`, `（2）...`) as a separate line
- For time allocation sections, write each item as a plain line: `第一部分 内容描述  ( N 周)`

Example form Markdown:

```markdown
## 1、本设计（论文）的目的、意义

随着新能源汽车的快速发展，FOC控制技术...本毕业设计的目的是...

## 2、学生应完成的任务

（1）系统性地调研FOC矢量控制技术...
（2）深入研究电机参数辨识方法...
（3）完成ESP32硬件电路的设计...
（4）研究并实现FOC核心算法...
（5）开发电机控制固件...
（6）撰写论文，制作答辩PPT，完成本科答辩。

## 3、本设计（论文）与本专业的毕业要求达成度如何？

（1）培养学生综合运用学科基础理论和专业知识...
（2）能够针对任务书要求设计复杂工程问题的解决方案...

## 4、本设计（论文）各部分内容及时间分配

第一部分 收集资料与阅读，撰写开题报告；  ( 2 周)
第二部分 完成算法理论学习与仿真验证；  ( 2 周)
```

#### Phase 2C: Complex Templates (`document_type: "complex"`)

Complex templates include cover pages, multi-section documents with multiple headers/footers, or documents with complex table layouts. The converter automatically preserves all document structure (sectPr, tables, images) and performs text replacement based on placeholder matching.

Write Markdown using descriptive headers that map to template placeholders:

| Markdown Header | Maps To Template Field |
|----------------|------------------------|
| `## 题目` / `## 标题` | Document title |
| `## 英文标题` | English title |
| `## 学院` / `## 系别` | School/Department |
| `## 专业` | Major |
| `## 年级` | Grade/Year |
| `## 学号` | Student ID |
| `## 姓名` | Student name |
| `## 指导教师` | Advisor |
| `## 日期` | Date |
| `## 摘要` | Abstract |
| `## 关键词` | Keywords |

Example complex template Markdown:

```markdown
## 题目

基于ESP32的无刷电机FOC控制系统设计

## 英文标题

Design of FOC Control System for Brushless Motor Based on ESP32

## 学院

物理与信息工程学院

## 专业

物联网工程

## 姓名

张三

## 学号

20210001

## 指导教师

李四 教授

## 日期

二〇二五年五月
```

### Phase 3: User Review

Present the complete Markdown to the user. Wait for confirmation or modification requests. Iterate until the user is satisfied.

### Phase 4: Convert to DOCX

The converter **automatically detects** the template type and uses the appropriate strategy:

- **Essay mode**: generates new `document.xml` from Markdown (heading-based templates)
- **Form-fill mode**: modifies the template's existing XML in-place, preserving all paragraph/run formatting (form-based templates with numbered sections)
- **Complex mode**: preserves entire document structure including multiple sections (sectPr), tables, headers/footers; performs placeholder-based text replacement (cover pages, complex layouts)

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
  - `analyze()` → full JSON analysis (includes `document_type` and `content_structure` for forms)
  - `detect_document_type()` → `"essay"` or `"form"`
  - `extract_outline()` → heading hierarchy (essay templates)
  - `extract_content_structure()` → numbered sections with labels, indices, sample content (form templates)
  - `extract_style_map()` → style name/ID map
  - `extract_page_layout()` → margins, size, header/footer refs
  - `extract_special_elements()` → table/formula/image/reference counts

### `scripts/search_references.py`

- `ReferenceSearcher(api_key=None)` — main class
  - `search(query, min_year, max_year, min_results, language, min_citations, open_access)` → list of formatted refs
  - `search_multi(topics, ...)` → merged deduplicated refs

### `scripts/md_to_docx.py`

- `MarkdownToDocxConverter(template_path)` — main class
  - `convert(markdown_content, output_path)` — auto-detects template type, uses essay or form-fill mode

### `scripts/verify_format.py`

- `verify_docx(template_path, output_path)` → `(passed, messages)`

## Example Usage

### Essay Example

```
参考 test_cases/Southwest Jiaotong University Thesis Template/正文.docx 的格式，
编写一篇关于"基于ESP32的无感FOC驱动设计"的绪论，要求有参考文献、公式、表格。
```

The AI will:
1. Run `analyze_template.py` → detects `document_type: "essay"`, gets heading outline
2. Generate Markdown with matching heading structure (Phase 2A)
3. Present for review
4. Convert to DOCX (essay mode) and verify formatting

### Form Example

```
参考 test_cases/Southwest Jiaotong University Thesis Template/任务书.docx 的格式，
编写"基于ESP32的FOC控制"的毕业设计任务书。
```

The AI will:
1. Run `analyze_template.py` → detects `document_type: "form"`, gets `content_structure` with sections
2. Generate Markdown with matching section labels (Phase 2B)
3. Present for review
4. Convert to DOCX (form-fill mode, preserving all original formatting)

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
