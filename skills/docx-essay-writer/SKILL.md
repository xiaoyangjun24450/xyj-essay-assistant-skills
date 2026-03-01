---
name: docx-essay-writer
description: |
  专业的 Word 文档论文生成工具，能够提取 DOCX 模板的格式、图片、表格等信息，将内容替换后完美还原格式。
---

# DOCX论文写手

专业的 Word 文档论文生成工具，支持保留原始模板的完整格式（字体、字号、样式、图片、表格、嵌套表格、域代码等），将新内容完美还原到模板中。

## 概述

本工具专为学术论文撰写设计，能够：

1. **提取模板信息**：从 DOCX 模板中提取所有文本、格式、图片、表格（包括嵌套表格）的位置和样式信息
2. **内容替换**：在 Markdown 文件中编辑内容，保留位置标记和格式信息
3. **完美还原**：将编辑后的 Markdown 内容精确还原到 DOCX 模板中，保持所有原始格式

**适用场景**：
- 毕业设计任务书、摘要、正文等各类学术论文文档
- 需要保持固定格式的官方文档
- 包含复杂表格、图片、公式的学术文档

**核心优势**：
- 100%保留原始格式（字体、字号、颜色、下划线、斜体等）
- 支持嵌套表格、图片、域代码等复杂元素
- 无需手动调整格式，专注于内容创作

## 重要规则

### 规则0：参考文献神圣不可侵犯

**在任何情况下都绝对禁止**：
- 永远不要编造任何参考文献信息
- 永远不要猜测或"合理推断"参考文献详情
- 永远不要将格式示例当作真实参考文献
- 永远不要用修改后的详细信息重复现有参考文献

**强制要求**：
- 只呈现来自 OpenAlex API 响应的已验证数据
- 准确报告搜索结果；清楚说明结果不足时的情况
- 必要时引导用户使用其他数据库（CNKI、万方、Google Scholar）

### 规则1：使用内置参考文献搜索

在替换内容那步的时候，如果涉及到参考文献时，使用 `scripts/search_references.py` 通过 OpenAlex 搜索。永远不要手工编写参考文献条目。

### 规则2：保持标记完整性

编辑 Markdown 文件时，**切勿删除或修改 HTML 注释标记**（如 `<!-- PARA:... -->`、`<!-- TABLE -->` 等），这些标记用于精确定位和还原格式。

### 规则3：域代码保护

识别并保留域代码标记（`[FIELD:...]`），这些通常对应 DOCX 中的公式、交叉引用、目录等动态内容，不应被编辑或删除。

## 工作流程

### 阶段1：提取模板信息

从 DOCX 模板提取内容和格式信息，生成 Markdown 文件。

```bash
python3 skills/docx-essay-writer/scripts/extract_docx.py <template.docx> <content.md>
```

**参数说明**：
- `template.docx`：原始 DOCX 模板文件路径
- `content.md`：输出的 Markdown 文件路径（可选，默认与模板同名）

**输出格式**：
Markdown 文件中每个文本段都带有位置和格式标记：
```
<!-- PARA:10,RUN:3,FORMAT:font_name=黑体|font_size=16.0|bold=True -->基于ESP32的FOC控制器设计
```
- `PARA:10`：第 10 个段落
- `RUN:3`：该段落的第 3 个文本段
- `FORMAT:`：格式信息（字体、字号、加粗等）

### 阶段2：编辑内容

#### 阶段2A：处理参考文献

如果文档中包含参考文献部分，使用内置搜索工具查找真实学术文献：

```bash
python3 skills/docx-essay-writer/scripts/search_references.py \
  --query "搜索主题" \
  --min-year 2015 \
  --max-year 2025 \
  --min-results 5 \
  --language en
```

**参数说明**：
- `--query, -q`：搜索关键词（必填）
- `--min-year`：最小年份
- `--max-year`：最大年份（默认当前年份）
- `--min-results`：最少返回结果数（默认 5）
- `--language`：语言代码（如 `en`、`zh`）
- `--min-citations`：最少引用次数
- `--open-access`：仅开放获取论文
- `--output, -o`：输出到文件

**输出格式**：GB/T 7714-2015 标准格式
```
[1] Smith J, Doe A. FOC Control Design[J]. IEEE Trans. Ind. Appl., 2020, 56(3): 1234-1245.
[2] Zhang L, Wang M. ESP32 Motor Control[C]. 2021.
```

#### 阶段2B：替换文本内容

根据用户要求，直接在 Markdown 文件中替换标记后的文本内容：

1. **找到需要修改的标记行**（如 `<!-- PARA:10,RUN:3,... -->旧内容`）
2. **保留完整的标记部分**（`<!-- PARA:10,RUN:3,FORMAT:... -->`）
3. **替换标记后的文本内容**
4. **不要添加或删除任何标记行**

**示例**：
```markdown
<!-- PARA:10,RUN:3,FORMAT:font_name=黑体|font_size=16.0|bold=True -->基于深度学习的图像识别系统
```

### 阶段3：还原为 DOCX

将编辑后的 Markdown 文件还原为 DOCX 文档，保持所有原始格式。

```bash
python3 skills/docx-essay-writer/scripts/restore_docx.py <content.md> <template.docx> <output.docx>
```

**参数说明**：
- `content.md`：编辑后的 Markdown 文件
- `template.docx`：原始 DOCX 模板文件
- `output.docx`：输出文件路径（可选，默认为 `content_restored.docx`）

**还原特性**：
- 保留所有格式信息（字体、字号、颜色、下划线等）
- 保留所有图片
- 保留所有表格结构（包括嵌套表格）
- 保留域代码（公式、交叉引用、目录等）
- 精确还原文本位置

## 支持的文档元素

### 文本格式
- ✅ 字体名称（如宋体、黑体、Times New Roman）
- ✅ 字号
- ✅ 加粗
- ✅ 斜体
- ✅ 下划线
- ✅ 删除线
- ✅ 上标/下标
- ✅ 文字颜色

### 段落结构
- ✅ 普通段落
- ✅ 标题段落
- ✅ 空行

### 表格
- ✅ 普通表格
- ✅ 嵌套表格（表格中套表格）
- ✅ 单元格格式

### 其他元素
- ✅ 图片（保留原始图片）
- ✅ 域代码（公式、交叉引用、目录等）
- ✅ 制表符、换行符

## 完整使用示例

### 示例1：生成毕业设计任务书

```bash
# 1. 提取模板
python3 skills/docx-essay-writer/scripts/extract_docx.py \
  test_cases/case1/带边框的任务书.docx \
  content.md

# 2. 编辑 content.md，修改内容（保留所有标记）

# 3. 还原为 DOCX
python3 skills/docx-essay-writer/scripts/restore_docx.py \
  content.md \
  test_cases/case1/带边框的任务书.docx \
  output.docx
```

### 示例2：带参考文献的论文正文

```bash
# 1. 提取模板
python3 skills/docx-essay-writer/scripts/extract_docx.py \
  template.docx content.md

# 2. 搜索参考文献
python3 skills/docx-essay-writer/scripts/search_references.py \
  --query "FOC control ESP32" \
  --min-year 2018 \
  --min-results 8 \
  --language en \
  -o references.txt

# 3. 将参考文献粘贴到 content.md 的参考文献部分

# 4. 还原为 DOCX
python3 skills/docx-essay-writer/scripts/restore_docx.py \
  content.md template.docx output.docx
```

### 示例3：包含表格和公式的复杂文档

```bash
# 1. 提取模板（自动处理表格和域代码）
python3 skills/docx-essay-writer/scripts/extract_docx.py \
  complex_template.docx content.md

# 2. 编辑内容，注意：
#    - 不要修改 [FIELD:...] 标记（这些是公式）
#    - 保留 <!-- TABLE --> 和 <!-- TABLE_END --> 标记
#    - 表格单元格标记格式：<!-- TABLE:0,ROW:0,CELL:0,PARA:0,RUN:0,FORMAT:... -->

# 3. 还原（自动保留表格结构和公式）
python3 skills/docx-essay-writer/scripts/restore_docx.py \
  content.md complex_template.docx output.docx
```

## 故障排除

### 问题1：还原后格式不正确

**原因**：Markdown 文件中标记被修改或删除

**解决方案**：
- 检查 `<!-- PARA:... -->` 等标记是否完整
- 不要删除或修改 HTML 注释标记
- 确保 `FORMAT:` 字段未被更改

### 问题2：参考文献搜索失败

**原因**：OpenAlex API 需要认证或网络问题

**解决方案**：
- 获取免费 API Key：https://openalex.org/settings/api
- 添加参数 `--api-key YOUR_KEY`
- 或使用 `--min-year` 缩小搜索范围
- 如果仍无法满足需求，引导用户使用 CNKI、万方等中文数据库

### 问题3：图片或公式丢失

**原因**：被错误识别为普通文本并清空

**解决方案**：
- 检查是否有 `[FIELD:...]` 标记被删除
- `restore_docx.py` 会自动保留非文本元素，确保不要手动干预

### 问题4：嵌套表格内容错位

**原因**：表格标记格式不正确

**解决方案**：
- 检查嵌套表格标记格式：`TABLE:0,ROW:0,CELL:0,NESTED:0,ROW:0,CELL:0,PARA:0,RUN:0`
- 确保外层表格和内层表格的标记都存在

## 技术细节

### Markdown 标记格式

**段落标记**：
```
<!-- PARA:段落索引,RUN:文本段索引,FORMAT:格式信息 -->文本内容
```

**表格标记**：
```
<!-- TABLE:表格索引,ROW:行索引,CELL:单元格索引,PARA:段落索引,RUN:文本段索引,FORMAT:格式信息 -->文本内容
```

**嵌套表格标记**：
```
<!-- TABLE:表格索引,ROW:外层行索引,CELL:外层单元格索引,NESTED:嵌套表格索引,ROW:嵌套行索引,CELL:嵌套单元格索引,PARA:段落索引,RUN:文本段索引,FORMAT:格式信息 -->文本内容
```

**格式字段**：
- `font_name`：字体名称
- `font_size`：字号（磅值）
- `bold`：加粗（True/False）
- `italic`：斜体（True/False）
- `underline`：下划线（True/False）
- `color`：颜色（RGB 值）
- `strike`：删除线（True/False）
- `superscript`：上标（True/False）
- `subscript`：下标（True/False）

## 依赖项

```txt
python-docx
pyalex
```

安装命令：
```bash
pip install python-docx pyalex
```

## 目录结构

```
skills/docx-essay-writer/
├── SKILL.md                      # 本文档
├── scripts/
│   ├── extract_docx.py          # 提取 DOCX 为 Markdown
│   ├── restore_docx.py          # 从 Markdown 还原为 DOCX
│   └── search_references.py     # 参考文献搜索工具
```

## 版本历史

- **v1.0**：初始版本，支持基本文档转换和格式保留
- **v1.1**：增加嵌套表格支持
- **v1.2**：增加域代码保护和参考文献搜索功能
