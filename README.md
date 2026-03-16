<div align="center">

# X-EAS

### Academic Essay Format Cloner for .docx

**提供一份 Word 模板，自动克隆其格式，生成内容全新的文档**

[![Python](https://img.shields.io/badge/Python-3.7+-blue?logo=python&logoColor=white)](https://www.python.org/)
[![Zero Dependencies](https://img.shields.io/badge/Dependencies-Zero-brightgreen)](./requirements.txt)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](./LICENSE)
[![GitHub Stars](https://img.shields.io/github/stars/xiaoyangjun24450/xyj-essay-assistant-skills?style=social)](https://github.com/xiaoyangjun24450/xyj-essay-assistant-skills)

[快速开始](#快速开始) · [工作原理](#工作原理) · [使用示例](#使用示例) · [项目结构](#项目结构)

</div>

---

## 为什么需要这个工具？

写论文最痛苦的是什么？**不是内容，是格式。**

页眉页脚怎么调都不对，行距字号改了又改，参考文献格式更是噩梦——更可怕的是，每个老师的要求都不一样，每次都要从头再来。

**X-EAS** 终结这种低效重复劳动：只需提供一份标准模板，工具即可自动解析其中的**全部格式规则**（字体、字号、加粗、页边距、页眉页脚……），并将这些样式精准应用到全新内容上。

> **一句话总结：内容换新，格式不动，一键生成。**

## 特性

- **格式零损失** — 深入 XML 层面解析 .docx，逐段逐 run 还原字体、字号、加粗、斜体、颜色等所有样式
- **公式支持** — 自动将 OMML 公式转为 LaTeX 标记，重写后再还原为 OMML
- **结构校验** — 内置校验脚本，确保重写后的段落结构、格式标记完整无误
- **零第三方依赖** — 纯 Python 标准库实现（`zipfile` + `xml.etree`），无需 `pip install`
- **AI Agent 友好** — 设计为 AI 编程助手的 Skill，可被 Cursor / Windsurf / Cline 等工具直接调用

## 适用场景

| 场景 | 说明 |
|------|------|
| 课程论文 | 快速匹配老师给的模板 |
| 毕业设计 | 符合学校官方格式要求 |
| 实验报告 | 统一实验室文档规范 |
| 竞赛文档 | 满足提交格式标准 |

## 快速开始

### 环境要求

- Python 3.7+（仅需标准库，无需安装任何依赖）

### 安装

```bash
git clone https://github.com/xiaoyangjun24450/xyj-essay-assistant-skills.git
cd xyj-essay-assistant-skills
```

无需执行 `pip install`，克隆即可使用。

### 作为 AI Skill 使用（推荐）

将此仓库配置为 AI 编程助手的 Skill 后，只需一句话：

```
参考 test/test2.docx 的 docx 模板，生成格式完全相同，
但主题内容为「基于 ESP32 的 FOC 控制器设计」的文档
```

AI 会自动完成格式提取 → 内容重写 → 校验 → 文档生成的全过程。

### 三步运行

```bash
# 1. 提取模板格式
python3 skills/docx-essay-writer/scripts/docx_preprocessor.py template.docx output/

# 2. 修改 output/chunks/ 中的文本内容（手动或 AI 辅助）

# 3. 还原为新文档
python3 skills/docx-essay-writer/scripts/docx_chunks_restorer.py \
    output/unzipped output/chunks_new output/new_document.docx
```


## 工作原理

```
┌─────────────┐     ┌──────────────┐     ┌──────────────┐     ┌─────────────┐
│  阶段 1      │     │  阶段 2       │     │  阶段 3       │     │  阶段 4      │
│  格式提取    │────▶│  内容重写     │────▶│  结构校验     │────▶│  文档还原    │
│              │     │              │     │              │     │              │
│ template.docx│     │ AI / 手动改写 │     │ 段落数·标记·  │     │ 合并格式与   │
│  ──▶ chunks  │     │ chunks 内容   │     │ 格式标记校验  │     │ 新内容为docx │
└─────────────┘     └──────────────┘     └──────┬───────┘     └─────────────┘
                                                │  ✗ 失败
                                                └──▶ 回退到阶段 2 修复（最多 3 次）
```

| 阶段 | 脚本 | 功能 |
|------|------|------|
| 1. 格式提取 | `docx_preprocessor.py` | 解压 docx，解析 XML，提取每个段落的格式指纹，输出带格式标记的 chunks |
| 2. 内容重写 | （AI / 手动） | 在保持格式标记和段落结构不变的前提下，替换文本内容 |
| 3. 结构校验 | `verify_chunks.py` | 对比原始与重写后的 chunks，检查行数、段落 ID、格式标记是否一致 |
| 4. 文档还原 | `docx_chunks_restorer.py` | 将重写内容写回 XML，处理 LaTeX→OMML 转换，重新打包为 .docx |

## 使用示例

### 命令行方式

**提取格式：**

```bash
python3 skills/docx-essay-writer/scripts/docx_preprocessor.py \
    test/test2.docx \
    output/test2_work
```

输出结构：

```
output/test2_work/
├── unzipped/           # docx 解压后的原始文件（还原用）
├── chunks/             # 带格式标记的段落文本
│   ├── chunk_0.md
│   ├── chunk_1.md
│   └── ...
├── origin_chunks/      # chunks 的原始备份
└── format.xml          # 格式化的 document.xml（调试用）
```

**校验重写内容：**

```bash
python3 skills/docx-essay-writer/scripts/verify_chunks.py \
    output/test2_work/chunks \
    output/test2_work/chunks_new
```

**还原为文档：**

```bash
python3 skills/docx-essay-writer/scripts/docx_chunks_restorer.py \
    output/test2_work/unzipped \
    output/test2_work/chunks_new \
    output/new_essay.docx
```

### AI Skill 方式

**语法格式：**

```
参考 [模板文件路径] 的 docx 模板，编写相同格式但内容不同的，标题为：[新文档标题]
```

**示例：**

```
参考模板：'test/test2.docx' 的格式，生成一篇格式完全相同，
但主题内容为'基于 ESP32 的 FOC 控制器设计'的文档
```

## 项目结构

```
xyj-essay-assistant-skills/
├── skills/
│   └── docx-essay-writer/
│       ├── SKILL.md                    # AI Skill 定义（工作流程描述）
│       └── scripts/
│           ├── docx_preprocessor.py    # 阶段 1：格式提取与 chunk 生成
│           ├── verify_chunks.py        # 阶段 3：结构校验
│           └── docx_chunks_restorer.py # 阶段 4：文档还原
├── test/                               # 示例模板文件
│   ├── test1.docx
│   ├── test2.docx
│   └── test3.docx
├── README.md
└── requirements.txt
```

## 注意事项

- 仅支持 `.docx` 格式（不支持 `.doc`）
- 模板文件路径必须正确且文件存在
- 重写时必须严格保持段落数量和格式标记结构不变
- 包含 LaTeX 公式的行，替换后仍需保持合法的 LaTeX 语法

## 常见问题

<details>
<summary><b>Q：支持 .doc 格式吗？</b></summary>

不支持。`.doc` 是二进制格式，本工具基于 XML 解析，仅支持 `.docx`。可以先用 Word 或 LibreOffice 将 `.doc` 转为 `.docx`。
</details>

<details>
<summary><b>Q：页眉页脚会被保留吗？</b></summary>

会。工具在 XML 层面操作，仅修改 `document.xml` 中的正文段落，页眉（`header*.xml`）、页脚（`footer*.xml`）、页面设置等均原样保留。
</details>

<details>
<summary><b>Q：图片和表格怎么处理？</b></summary>

当前版本聚焦于文本段落和公式的格式克隆。图片和表格在模板中的位置结构会被保留，但内容替换需人工处理。
</details>

<details>
<summary><b>Q：可以不用 AI，纯手动操作吗？</b></summary>

可以。阶段 2（内容重写）可以手动编辑 chunk 文件，只需遵守格式约束即可。
</details>

## 参与贡献

欢迎提交 Issue 和 Pull Request！

<!-- TODO: 添加 CONTRIBUTING.md 后更新此链接 -->

## 许可证

本项目基于 [MIT License](./LICENSE) 开源，你可以自由使用、修改和分发。

---

<div align="center">

**如果这个项目对你有帮助，请点个 Star ⭐**

</div>
