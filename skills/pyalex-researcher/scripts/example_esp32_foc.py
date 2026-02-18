#!/usr/bin/env python3
"""
示例：为 ESP32 无感 FOC 控制器设计开题报告搜索参考文献
"""

from search_references import ReferenceSearcher


def main():
    # 创建搜索器
    searcher = ReferenceSearcher()
    
    print("=" * 60)
    print("为 ESP32 无感 FOC 控制器设计搜索参考文献")
    print("=" * 60)
    
    # 搜索策略：使用多个相关关键词进行搜索
    search_topics = [
        "ESP32 motor control FOC",
        "sensorless FOC motor drive",
        "field oriented control BLDC",
        "ESP32 BLDC controller"
    ]
    
    all_references = []
    
    # 搜索每个主题，获取最近10年的英文文献
    for topic in search_topics:
        print(f"\n搜索: {topic}")
        refs = searcher.search(
            query=topic,
            min_year=2015,      # 10年内
            max_year=2025,
            min_results=2,      # 每个主题至少2篇
            language="en"       # 英文文献
        )
        all_references.extend(refs)
    
    # 去重并限制总数
    seen_titles = set()
    unique_refs = []
    for ref in all_references:
        # 提取标题进行去重
        title = ref.split('.')[1].strip() if '.' in ref else ref
        if title not in seen_titles:
            seen_titles.add(title)
            unique_refs.append(ref)
    
    # 重新编号
    final_refs = []
    for i, ref in enumerate(unique_refs[:5], 1):
        # 替换编号
        new_ref = f"[{i}]" + ref[ref.find(']')+1:]
        final_refs.append(new_ref)
    
    print("\n" + "=" * 60)
    print("参考文献列表 (GB/T 7714-2015 格式)")
    print("=" * 60)
    print("\n".join(final_refs))
    
    return final_refs


if __name__ == "__main__":
    main()
