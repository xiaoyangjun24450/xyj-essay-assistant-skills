---
name: pyalex-researcher
description: |
  This skill should be used when the user needs to search for academic references and citations using OpenAlex database.
  Use this skill when:
  - User asks for academic references for a research topic
  - User needs citations for a thesis, proposal, or paper
  - User wants to find papers on a specific topic with filtering (year, language, etc.)
  - User needs to generate a bibliography/reference list
  - User wants to find recent papers (e.g., last 10 years) on a topic
---

# PyAlex Academic Researcher

Search for academic literature using OpenAlex database and generate formatted references for research papers, proposals, and theses.

## Overview

This skill provides a workflow and script for:
1. Searching academic papers from OpenAlex database using PyAlex
2. Filtering results by publication year, language, citation count, etc.
3. Generating formatted references in GB/T 7714-2015 (Chinese academic standard)

## Features

- **Smart Search**: Uses OpenAlex's powerful search to find relevant papers
- **Flexible Filtering**: Filter by year range, language, citation count, open access
- **Formatted Output**: Generate references in standard academic format
- **Batch Operations**: Search multiple topics and merge results

## Prerequisites

```bash
pip install pyalex
```

### API Key Configuration (Required)

Starting February 2026, OpenAlex requires an API key for most requests.

1. Get a free API key from [openalex.org/settings/api](https://openalex.org/settings/api)
2. Configure the API key in your code:

```python
from skills.pyalex-researcher.scripts.search_references import ReferenceSearcher

searcher = ReferenceSearcher(api_key="YOUR_API_KEY")
```

Or set it globally:

```python
import pyalex
pyalex.config.api_key = "YOUR_API_KEY"
```

## Quick Start

### Command Line Usage

```bash
python skills/pyalex-researcher/scripts/search_references.py \
  --query "ESP32 FOC motor control" \
  --min-year 2015 \
  --max-year 2025 \
  --min-results 5 \
  --output refs.txt
```

### Python API Usage

```python
from skills.pyalex-researcher.scripts.search_references import ReferenceSearcher

searcher = ReferenceSearcher()
references = searcher.search(
    query="ESP32 FOC motor control",
    min_year=2015,
    max_year=2025,
    min_results=5
)

for ref in references:
    print(ref)
```

## Parameters

| Parameter | Type | Default | Description |
|-----------|------|---------|-------------|
| `query` | str | required | Search query for papers |
| `min_year` | int | None | Minimum publication year |
| `max_year` | int | None | Maximum publication year |
| `min_results` | int | 5 | Minimum number of results to return |
| `language` | str | None | Filter by language (e.g., "en" for English) |
| `min_citations` | int | None | Minimum citation count |
| `open_access` | bool | None | Filter for open access papers only |
| `output` | str | None | Output file path (optional) |

## Reference Format

Generated references follow GB/T 7714-2015 format:

```
[1] Author A, Author B. Title of the paper[J]. Journal Name, Year, Vol(Issue): Pages.
[2] Author C. Book Title[M]. City: Publisher, Year.
```

## Example Use Cases

### Example 1: Opening Report References (开题报告文献搜索)

User request: "帮我找5篇10年内关于ESP32无感FOC控制的英文参考文献"

```python
import sys
sys.path.insert(0, 'skills/pyalex-researcher/scripts')

from search_references import ReferenceSearcher

# 配置 API Key (从 https://openalex.org/settings/api 获取)
API_KEY = "your_api_key_here"

# 创建搜索器
searcher = ReferenceSearcher(api_key=API_KEY)

# 使用多个关键词组合搜索
topics = [
    "ESP32 motor control",
    "sensorless FOC field oriented control", 
    "BLDC motor driver controller"
]

all_refs = []
for topic in topics:
    refs = searcher.search(
        query=topic,
        min_year=2015,
        max_year=2025,
        min_results=2,
        language="en"
    )
    all_refs.extend(refs)

# 输出前5篇不重复的文献
for ref in all_refs[:5]:
    print(ref)
```

**运行演示：**
```bash
python demo_reference_search.py
```

Output:
```
[1] Zhang S, Li M. Sensorless FOC Control of PMSM Based on ESP32[J]. IEEE Transactions on Industrial Electronics, 2023, 70(5): 4567-4576.
[2] Wang H, et al. Design of FOC Controller for BLDC Motor Using ESP32[J]. IEEE Access, 2022, 10: 12345-12356.
...
```

### Example 2: Multiple Topic Search

```python
# Search multiple related topics and merge results
topics = [
    "ESP32 motor control",
    "FOC field oriented control",
    "sensorless motor drive"
]

all_refs = []
for topic in topics:
    refs = searcher.search(topic, min_year=2018, min_results=2)
    all_refs.extend(refs)
```

### Example 3: High-Impact Papers Only

```python
# Search for highly cited papers
references = searcher.search(
    query="field oriented control",
    min_year=2010,
    min_citations=50,
    min_results=5
)
```

## Script Reference

### `scripts/search_references.py`

Main search script with classes:

- `ReferenceSearcher` - Main search class
  - `search(query, **filters)` - Search and return formatted references
  - `search_multi(topics, **filters)` - Search multiple topics
  - `format_reference(work)` - Format a single work to citation string

## Advanced Usage

### Custom Reference Format

```python
def custom_format(work):
    authors = work.get('authorships', [])
    author_names = [a['author']['display_name'] for a in authors[:3]]
    if len(authors) > 3:
        author_names.append("et al.")
    return f"{', '.join(author_names)}. {work['title']}. {work['publication_year']}"

searcher = ReferenceSearcher(formatter=custom_format)
```

### Using PyAlex Directly

For advanced queries, use PyAlex directly:

```python
from pyalex import Works

# Complex filter query
works = Works() \
    .search("ESP32 motor control") \
    .filter(publication_year=">2015", language="en") \
    .sort(cited_by_count="desc") \
    .get()
```

## Troubleshooting

### No Results Found
- Try broader search terms
- Remove some filters (year, language)
- Check if topic exists in OpenAlex database

### API Rate Limits
- OpenAlex requires API key for high-volume usage
- Free tier: 100,000 credits/day with API key
- Set API key: `pyalex.config.api_key = "YOUR_KEY"`

## References

- [PyAlex Documentation](https://github.com/J535D165/pyalex)
- [OpenAlex API Docs](https://docs.openalex.org/)
- GB/T 7714-2015: 信息与文献 参考文献著录规则
