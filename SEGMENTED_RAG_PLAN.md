# 分段检索式改写流程计划

## 目标

将现有流程从“按字数切 chunk 后直接逐块改写”，升级为“先理解模板结构，再按语义切块，并结合用户素材检索后逐块改写”。

## 新流程

1. `docx -> full_paragraphs`
   Python 从模板 `docx` 提取整篇正文，生成一个按段落展开的全文文件。
   每行保留一个 `[PARA_ID]`，作为后续分段、改写、还原的唯一锚点。

2. `LLM 建立 segment -> para_ids`
   将 `full_paragraphs` 分批送给 LLM，识别每一段内容大致在讲什么。
   输出结构化结果：每个 `segment` 的摘要、用途、覆盖的 `para_ids`。
   示例 JSON：
   ```json
   [
     {
       "segment_id": "seg_1",
       "summary": "介绍永磁同步电机控制的技术背景与研究意义",
       "purpose": "background",
       "para_ids": ["A1B2C3D4", "A1B2C3D5", "A1B2C3D6"]
     },
     {
       "segment_id": "seg_2",
       "summary": "说明 FOC 的坐标变换与 dq 轴解耦控制原理",
       "purpose": "principle",
       "para_ids": ["A1B2C3D7", "A1B2C3D8"]
     }
   ]
   ```
   Python 负责校验：不能丢 `para_id`、不能重复、顺序不能乱、必须覆盖全文。

3. `segment -> chunks`
   Python 根据 `segment -> para_ids` 生成新的改写单元。
   切块原则改为“语义优先，字数兜底”：
   先按 `segment` 切；若单个 `segment` 过大，再在该段内部二次拆分。

4. `素材 -> 语料库`
   用户上传的 `txt/docx/md` 先提取文字，再切成素材片段。
   对素材片段建立索引，而不是对整篇文件建索引。
   每个素材片段保留来源信息，如 `source_id`、文件名、位置、原文。

5. `每个 chunk 检索语料`
   每个改写 `chunk` 用以下信息联合检索：
   用户写作任务 + 当前 chunk 原文 + 当前 segment 摘要。
   从语料库中找到最相关的素材片段，作为该 chunk 的专属写作参考。

6. `组装逐块写作任务并调用 LLM`
   对每个 `chunk`，Python 组装一个专属任务包，内容包括：
   全局写作需求、当前 chunk 原文、segment 摘要、检索到的素材片段、必须保留的信息。
   然后将该任务包发给 LLM 做逐块改写。

7. `Python 校验并还原`
   Python 校验改写结果是否保持：
   行数不变、`[PARA_ID]` 不变、格式标签不变、公式结构合法。
   校验通过后，再按现有流程还原为新的 `docx`。

## 这版方案的核心变化

- 改写前先做“模板结构理解”，而不是直接按字数切块。
- 改写单元从“纯字数 chunk”改为“语义优先 chunk”。
- 用户素材不再全文塞进 prompt，而是先建语料库，再按 chunk 定向检索。
- LLM 负责语义分析和写作，Python 负责切分、校验、索引和还原。
