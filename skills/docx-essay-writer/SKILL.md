---
name: docx-essay-writer
description: 专业的 Word 文档论文生成工具，提取 DOCX 模板格式，替换内容后完美还原。
---

# DOCX论文写手

## ⚠️ 关键认知（执行前必读）

**本技能的核心价值是格式保真度，不是执行速度。**

DOCX中的每个文本段（Run）都有独立的格式属性（字体、字号、加粗、颜色等），这些信息存储在`<!-- PARA:X,RUN:Y,FORMAT:... -->`标记中。

**两种编辑方式**：
1. **整体重写（阶段2B）**：可以重写整个markdown，但必须通过阶段2C校验确保格式标记完整
2. **逐行替换（阶段2D）**：逐个替换标记后的文本，最安全但效率低

**关键原则**：
- 阶段1生成的原始markdown（`content.md`）永远不能被修改
- 格式标记（`<!--...-->`）和域代码（`[FIELD:...]`）必须完整保留
- 任何批量处理（脚本、正则、全文替换）都需要通过校验脚本验证

---

## 工作流程

### 阶段1：提取
```bash
python3 skills/docx-essay-writer/scripts/extract_docx.py <template.docx> <content.md>
```

### 阶段2：编辑

#### 2A. 参考文献（如需要）
```bash
python3 skills/docx-essay-writer/scripts/search_references.py \
  --query "关键词" --min-year 2015 --min-results 5 -o refs.txt
```

#### 2B. Markdown重写
**说明**：在这个阶段，你要根据用户的期望内容去重写整个markdown文件，生成新的内容，使得生成的markdown文件在去掉标记和域代码的情况下，有文本可读性，同时要注意新的markdown不能含有旧markdown的专业术语之类的。新的markdown文件可以命名为 `content_rewritten.md`（或其他你喜欢的名称）。

**注意事项**：
- 重写时可以自由修改文本内容
- 但要完全保留格式标记（`<!--...-->`）和域代码（`[FIELD:...]`）


#### 2C. 格式标记校验与修复
**说明**：运行校验脚本，对比原始markdown和重写后的markdown，确保所有格式标记和域代码都完整保留。

```bash
python3 skills/docx-essay-writer/scripts/validate_markdown.py \
  content.md content_rewritten.md content_final.md
```

**脚本功能**：
- 提取两个文件中的所有格式标记和域代码
- 对比差异，找出缺失或被修改的标记
- 如果有问题，自动生成修复后的文件 `content_final.md`
- 修复策略：使用原始文件的标记，替换为重写文件的文本内容

**校验结果**：
- ✅ 通过：直接进入阶段3
- ❌ 失败：查看错误信息，使用生成的 `content_final.md` 进入阶段3

**注意**：无论校验是否通过，阶段1生成的原始markdown（`content.md`）都不会被修改。

#### 2D. 文本校验
**说明

**标记示例**：

校验新生成的markdown文件是否还含有符合旧markdown主题内容，但不符合新markdown主题内容的文本内容，如果有就返回2B. Markdown重写

### 阶段3：还原
**说明**：使用校验通过后的markdown文件进行还原。

```bash
python3 skills/docx-essay-writer/scripts/restore_docx.py <content_final.md> <template.docx> <output.docx>
```

**文件选择优先级**：
1. 如果阶段2C校验通过并生成了 `content_final.md`，使用该文件
2. 如果阶段2C校验通过但未生成修复文件，使用 `content_rewritten.md`
3. 如果未使用阶段2B重写，直接使用 `content.md`

---

## 规则

| 规则 | 内容 |
|-----|------|
| 参考文献 | 仅使用OpenAlex真实数据，禁止编造 |
| 标记保护 | 绝不修改`<!--...-->`和`[FIELD:...]`（除非使用校验脚本修复） |
| 工具锁定 | 阶段2D仅使用replace_in_file |
| 执行粒度 | 阶段2D：1标记=1次调用，禁止批量；阶段2B：可整体重写 |
| 校验要求 | 阶段2B重写后必须执行阶段2C校验 |

---

## 常见错误模式（在此阶段会触发失败）

**错误A：效率焦虑（阶段2D）**
- 症状："我应该写个脚本批量处理"
- 后果：格式标记丢失，DOCX还原失败
- 纠正：信任标记系统，逐行是唯一正确路径（仅在阶段2D时）

**错误B：跳过校验**
- 症状：阶段2B重写后直接进入阶段3，跳过阶段2C校验
- 后果：格式标记或域代码可能丢失，导致DOCX还原失败
- 纠正：必须执行阶段2C校验，确保所有标记完整保留

**错误C：错误使用阶段**
- 症状：在应该使用阶段2D逐行替换时使用了阶段2B整体重写
- 后果：需要额外的校验和修复步骤
- 纠正：根据需求选择合适的编辑方式

---

## 依赖
```bash
pip install python-docx pyalex
```
