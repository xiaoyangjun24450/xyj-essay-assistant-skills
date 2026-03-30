# 阶段一设计说明

## 1. 阶段一的目标

阶段一不直接生成最终正文，也不负责把内容写回 `docx`。

它的职责只有一件事：

**把多个碎片化素材整理成一份可供后续改写流程使用的“写作资料包”。**

更直白地说，阶段一做的是：

1. 读取多个素材文件。
2. 提取文本内容。
3. 压缩可以压缩的信息。
4. 保留不能丢失的硬信息。
5. 给每个 `chunk` 决定“应该写什么”和“可以参考哪些素材”。

## 2. 阶段一的输入

阶段一的输入包含两部分。

### 2.1 素材输入

用户上传的多个碎片化文件，计划支持：

- `pdf`
- `md`
- `txt`
- `docx`

这些素材可能是：

- 论文全文
- 参考资料
- 课程要求
- 已有笔记
- 项目说明

### 2.2 模板分析结果

阶段一还需要知道当前模板 `docx` 被拆成了哪些 `chunk`，因为它最终要给每个 `chunk` 分配写作资料。

这里默认复用当前系统已有流程：

1. 模板 `docx` 先经过预处理。
2. 生成 `chunks/`
3. 阶段一读取这些 `chunk` 的文本内容，判断每个 `chunk` 大概属于什么角色。


## 2. 阶段一的输出

阶段一建议输出两类结果。

### 5.1 全局资料

用于所有 `chunk` 共享的短信息。

例如：

- 文章主题
- 写作类型
- 全局核心观点
- 必须覆盖的点
- 禁止编造的点

建议文件：

- `global_requirement.txt`

### 5.2 每个 chunk 的局部资料

用于告诉后续改写器：

- 这个 `chunk` 应该写什么
- 这个 `chunk` 可以参考哪些素材
- 这个 `chunk` 是否需要带公式
- 这个 `chunk` 是否要强调背景、方法、结论或展望

建议文件：

- `chunk_contexts.json`


## 6. 阶段一的核心工作流

阶段一建议拆成 6 步。

### 第 1 步：提取所有素材文本

目标：

- 把不同格式的输入统一成纯文本或结构化文本。

建议产物：

- `sources/normalized/*.md`
- `sources_manifest.json`

每条素材至少记录：

- `source_id`
- 文件名
- 文件类型
- 来源路径
- 提取后的文本
- 如果有的话，页码、标题、章节名

### 第 2 步：切分素材

目标：

- 把长文档切成适合后续分析的小块，避免一次性处理整篇论文。

切分原则：

- 按章节切优先于按字数硬切。
- 控制每块长度，便于做摘要和检索。
- 每块保留来源信息。

建议产物：

- `source_chunks.jsonl`

每个块至少包含：

- `chunk_id`
- `source_id`
- 所属章节
- 页码范围
- 原始文本

### 第 3 步：抽取结构化信息

目标：

- 不直接生成长摘要，而是先抽取“后续写作真正需要的信息”。

建议抽取的信息类型：

- 主题
- 核心观点
- 关键事实
- 关键数字
- 核心公式
- 结论
- 局限性
- 适合引用到哪些章节

这里最重要的是区分两类信息。

#### 可压缩信息

这些可以被归纳成短摘要：

- 背景介绍
- 研究意义
- 一般性结论
- 方法概述
- 应用价值

#### 不可丢信息

这些不能只写成概括句，必须单独保留：

- 核心公式
- 关键数字
- 参数条件
- 明确的实验结论边界
- 定义和符号说明

建议产物：

- `source_facts.jsonl`

每条记录可包含：

```json
{
  "fact_id": "paper1_sec2_01",
  "source_id": "paper1",
  "section": "2.2 Park变换",
  "type": "formula",
  "summary": "说明dq轴电流分解关系",
  "content": "$i_d = i_\\alpha \\cos\\theta + i_\\beta \\sin\\theta$",
  "must_keep": true,
  "usable_for": ["principle", "method"],
  "notes": "公式符号不要改写"
}
```

### 第 4 步：生成全局资料包

目标：

- 将多个素材中的信息整合成一份短而稳的全局写作说明。

这里不是把所有内容全塞进去，而是只保留所有 `chunk` 都可能需要的共享信息。

应包含：

- 主题
- 文体
- 核心观点
- 必写点
- 禁止编造点
- 全局写作约束

建议产物：

- `global_requirement.txt`

示意：

```text
主题：基于 ESP32 的 FOC 控制器设计

必须覆盖：
- FOC 控制原理
- ESP32 的主控角色
- 硬件与软件设计流程
- 系统可行性与工程意义

禁止编造：
- 未提供的实验数据
- 未提供的芯片参数
- 未提供的参考文献条目
```

### 第 5 步：识别模板 chunks 的角色

目标：

- 判断每个模板 `chunk` 更像标题、摘要、引言、正文、结论还是参考文献区。

可结合的信息：

- `chunk` 原始文本中的关键词
- 当前模板中的标题格式和段落分布
- 章节标题文本

常见角色：

- `title`
- `abstract`
- `keywords`
- `introduction`
- `body_principle`
- `body_design`
- `body_experiment`
- `conclusion`
- `references`

建议产物：

- `chunk_roles.json`

### 第 6 步：为每个 chunk 分配资料

目标：

- 不把所有素材都发给每个 `chunk`，而是定向分配最相关的信息。

这是阶段一最关键的一步。

每个 `chunk` 需要拿到：

- 它应该写什么
- 它可以用哪些事实
- 它需要保留哪些公式
- 它不能编造什么

建议产物：

- `chunk_contexts.json`

示例：

```json
{
  "chunk_0.md": {
    "role": "title",
    "should_write": "论文标题，突出ESP32与FOC控制器设计",
    "use_facts": ["paper2_title", "paper2_abs_01"],
    "must_keep_formulas": []
  },
  "chunk_3.md": {
    "role": "body_principle",
    "should_write": "解释FOC原理、dq轴分解和控制思路",
    "use_facts": ["paper1_sec2_01", "paper1_sec2_02", "paper1_sec3_01"],
    "must_keep_formulas": [
      "$i_d = i_\\alpha \\cos\\theta + i_\\beta \\sin\\theta$",
      "$i_q = -i_\\alpha \\sin\\theta + i_\\beta \\cos\\theta$"
    ]
  }
}
```


## 7. 核心公式如何处理

核心公式不能像普通文字那样只做摘要。

阶段一里，公式应遵循下面的原则：

1. 公式单独提取。
2. 尽量保存为 LaTeX。
3. 公式原文和公式说明分开存。
4. 只把相关公式分配给相关 `chunk`。

例如：

- 摘要段通常不需要塞详细公式。
- 原理分析段应带核心公式。
- 结论段一般不需要重复公式。

如果素材来自 `docx/md/txt`，公式提取相对容易。

如果素材来自 `pdf`，第一版可以采用保守策略：

- 普通文字先提取。
- 核心公式尽量提取成文本。
- 对识别不稳定的公式，允许进入“人工确认”或后续增强流程。


## 8. 阶段一和阶段二的关系

阶段一负责“决定每个 chunk 写什么”。

阶段二负责“按照阶段一的资料去真正改写每个 chunk”。

所以阶段二的输入，不应该只有：

- 原始 `chunk`
- 用户写的简短 requirement

而应该变成：

- 原始 `chunk`
- 全局 `global_requirement`
- 当前 `chunk` 的 `role`
- 当前 `chunk` 的 `should_write`
- 当前 `chunk` 对应的素材片段
- 当前 `chunk` 必须保留的公式/数字/限制条件


## 9. MVP 范围建议

为了先做出可运行的第一版，建议阶段一先做最小范围。

### 第一版建议支持

- 素材格式：`txt`、`md`、`docx`
- 输出：`global_requirement.txt` + `chunk_contexts.json`
- 信息类型：主题、观点、事实、公式、禁止编造项

### 第一版暂时不强求

- 高质量 `pdf` 公式识别
- 图片、表格抽取
- 参考文献自动生成
- 多轮检索增强


## 10. 阶段一最终产物清单

建议阶段一最终至少生成以下文件：

```text
stage1/
├── sources_manifest.json
├── source_chunks.jsonl
├── source_facts.jsonl
├── chunk_roles.json
├── chunk_contexts.json
└── global_requirement.txt
```

它们分别负责：

- `sources_manifest.json`：记录上传了哪些素材
- `source_chunks.jsonl`：记录素材被切成了哪些块
- `source_facts.jsonl`：记录抽取出的结构化事实、公式、数字和限制
- `chunk_roles.json`：记录模板里的每个 `chunk` 是什么角色
- `chunk_contexts.json`：记录每个 `chunk` 该写什么、用什么材料
- `global_requirement.txt`：记录所有 `chunk` 共用的短写作要求


## 11. 一句话总结

阶段一不是“把所有素材压成一段摘要”。

阶段一真正要做的是：

**把多源素材整理成一份既能压缩 token、又尽量不丢关键信息、还能按 chunk 定向分发的写作资料包。**
