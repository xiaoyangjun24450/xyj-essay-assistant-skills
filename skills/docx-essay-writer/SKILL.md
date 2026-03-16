---
name: docx-essay-writer
description: 给定一份 Word 模板文件和新主题，生成格式完全相同但内容不同的新 Word 文档。
  TRIGGER when: 用户提供 .docx 模板路径和新主题描述，要求生成格式相同内容不同的新文档。
  示例触发词："参考模板生成新文档"、"格式不变换主题"、"按模板写一篇关于X的文章"。
---

# docx-essay-writer

## 使用方式

```
参考 [模板文件路径] 的 docx 模板，生成格式完全相同但主题为「[新主题]」的文档
```

## 工作流程

此 skill 包含 4 个阶段。**在开始前，先确认用户提供了**：
1. 模板 `.docx` 文件路径（`TEMPLATE_DOCX`）
2. 新文档主题描述（`NEW_TOPIC`）

以下所有脚本路径相对于本 skill 目录（`skills/docx-essay-writer/`）。

---

### 准备工作

设定工作目录变量：

```bash
SKILL_DIR="skills/docx-essay-writer"
WORK_DIR="output/$(basename $TEMPLATE_DOCX .docx)_$(date +%Y%m%d_%H%M%S)"
mkdir -p "$WORK_DIR"
```

---

### 阶段 1：提取文本和格式

使用预处理器提取 docx 中的文本、格式信息和文档结构：

```bash
python3 "$SKILL_DIR/scripts/docx_preprocessor.py" "$TEMPLATE_DOCX" "$WORK_DIR"
```

**期望输出：**
- `$WORK_DIR/unzipped/` — docx 解压后的原始结构（还原用）
- `$WORK_DIR/chunks/chunk_0.md`, `chunk_1.md`, ... — 带 `[PARA_ID]` 标记的段落文本（含格式标记）
- `$WORK_DIR/origin_chunks/` — chunks 的原始备份
- `$WORK_DIR/format.xml` — 格式化后的 document.xml

**格式标记说明：**
- 每行格式：`[PARA_ID] 段落文字`
- 格式差异用 `<f=字体,sz=字号,b=加粗,...>文本</...>` 标记包裹
- 公式保留为 `$...$` LaTeX 格式

如果脚本报错，停止并告知用户错误信息。

---

### 阶段 2：重写 chunks

读取 `$WORK_DIR/chunks/` 下的所有 chunk 文件，根据新主题重写内容。

**约束（必须严格遵守）：**
1. **行数不变** — 每个 `[PARA_ID]` 对应一行，不能增删行
2. **标记不变** — `[PARA_ID]` 标记必须原样保留
3. **格式标记保留** — 如果原文有 `<...>...</...>` 格式标记，新内容也要保留相同标记结构
4. **公式保留** — 包含 `$...$` 的行：保留公式定界符和 LaTeX 语法；如需替换，换成新主题相关的合法 LaTeX 公式，不得改为普通文本
5. **结构保持** — 标题依然是标题，正文依然是正文

**处理流程：**

1. 列出所有 chunk 文件：
```bash
ls -la "$WORK_DIR/chunks/"
```

2. 逐个读取 chunk 文件内容，根据 `NEW_TOPIC` 重写每行文字

3. 将重写后的内容写入 `$WORK_DIR/chunks_new/` 目录（保持相同文件名）

---

### 阶段 3：校验与修复

使用校验脚本验证重写后的 chunks：

```bash
python3 "$SKILL_DIR/scripts/verify_chunks.py" \
    "$WORK_DIR/chunks" \
    "$WORK_DIR/chunks_new"
```

**校验内容：**
- 文件数量是否一致
- 每个 chunk 行数是否与原 chunk 相同
- `[PARA_ID]` 标记是否完整保留
- 格式标记 `<...>` 是否正确闭合

**结果处理：**

| 退出码 | 含义 | 处理方式 |
|--------|------|----------|
| 0 | 校验通过 | 进入阶段 4 |
| 1 | 校验失败 | 回退到阶段 2 修复 |

**修复流程（校验失败时）：**
1. 根据错误信息定位问题 chunk 和具体行
2. 仅修改有问题的行，其他行保持不变
3. 修复后重新运行校验脚本
4. 最多重试 3 次，仍失败则标记该 chunk 需人工介入

---

### 阶段 4：还原为 docx

校验通过后，将重写后的 chunks 还原为 Word 文档：

```bash
OUTPUT_DOCX="output/$(basename $TEMPLATE_DOCX .docx)_${NEW_TOPIC}.docx"
python3 "$SKILL_DIR/scripts/docx_chunks_restorer.py" \
    "$WORK_DIR/unzipped" \
    "$WORK_DIR/chunks_new" \
    "$OUTPUT_DOCX"
```

**期望输出：**
- `$OUTPUT_DOCX` — 格式与模板完全相同、内容为新主题的 Word 文档

**完成！** 告知用户输出文件路径：`$OUTPUT_DOCX`

---

## 错误处理汇总

| 情况 | 处理方式 |
|------|----------|
| 阶段 1 脚本报错 | 停止，展示错误给用户 |
| 阶段 2 重写后格式标记丢失 | 阶段 3 校验会捕获，回退修复 |
| 阶段 3 校验失败 | 回退到阶段 2 修复，最多重试 3 次 |
| 阶段 3 反复校验失败 | 标记该 chunk 需人工介入，询问是否继续或中止 |
| 阶段 4 还原失败 | 检查 chunks_new 是否完整，必要时重新执行阶段 2-3 |

---

## 文件结构说明

```
$WORK_DIR/
├── unzipped/           # docx 解压后的原始结构
│   ├── word/
│   │   └── document.xml
│   └── ...
├── chunks/             # 原始 chunks（阶段1生成）
│   ├── chunk_0.md
│   ├── chunk_1.md
│   └── ...
├── origin_chunks/      # chunks 备份
│   └── ...
├── chunks_new/         # 重写后的 chunks（阶段2生成）
│   ├── chunk_0.md
│   ├── chunk_1.md
│   └── ...
└── format.xml          # 格式化的 document.xml
```
