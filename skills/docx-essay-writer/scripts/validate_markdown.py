#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
校验原始markdown和重写后的markdown，确保所有格式标记和域代码都完整保留

用法:
    python validate_markdown.py <原始markdown路径> <重写后的markdown路径> [输出路径]

示例:
    python validate_markdown.py original.md rewritten.md
    python validate_markdown.py original.md rewritten.md fixed.md

功能:
    1. 提取原始markdown中的所有格式标记和域代码
    2. 提取重写后markdown中的所有格式标记和域代码
    3. 对比差异，找出缺失或被修改的标记
    4. 如果有缺失，自动修复重写后的markdown（保留阶段1的标记，只替换文本内容）
"""

import sys
import os
import re
from typing import Set, Dict, List, Tuple


def extract_markers_and_fields(markdown_path: str) -> Tuple[Set[str], Set[str], List[Dict]]:
    """
    提取markdown中的所有格式标记和域代码

    返回:
        markers: 所有格式标记的集合
        fields: 所有域代码的集合
        lines_info: 每一行的详细信息（标记、内容）
    """
    markers = set()
    fields = set()
    lines_info = []

    with open(markdown_path, 'r', encoding='utf-8') as f:
        for line_num, line in enumerate(f, 1):
            line = line.rstrip('\n')

            # 提取格式标记 <!-- PARA:... --> 或 <!-- TABLE:... -->
            marker_match = re.search(r'<!-- (PARA:\d+|TABLE:\d+|TABLE_END|TABLE|PAGE_BREAK|PAGE_BREAK_BEFORE)[^>]*-->', line)
            if marker_match:
                marker = marker_match.group(0)
                markers.add(marker)
                lines_info.append({
                    'line_num': line_num,
                    'line': line,
                    'marker': marker,
                    'content': line.split('-->')[-1] if '-->' in line else ''
                })

            # 提取域代码 [FIELD:...]
            field_matches = re.findall(r'\[FIELD:[^\]]+\]', line)
            for field in field_matches:
                fields.add(field)

    return markers, fields, lines_info


def extract_marker_dict(markers: Set[str]) -> Dict[str, str]:
    """
    将标记集合转换为字典，key为标记前缀（不含文本），value为完整标记
    用于快速查找某个标记是否存在
    """
    marker_dict = {}
    for marker in markers:
        # 提取标记部分（不包含后面的文本内容）
        marker_part = marker.split('-->')[0] + '-->'
        marker_dict[marker_part] = marker
    return marker_dict


def validate_and_fix(original_path: str, rewritten_path: str, output_path: str = None) -> bool:
    """
    校验并修复markdown文件

    返回:
        True: 文件已修复或无问题
        False: 需要修复但未生成输出文件
    """
    # 提取原始文件信息
    orig_markers, orig_fields, orig_lines = extract_markers_and_fields(original_path)
    orig_marker_dict = extract_marker_dict(orig_markers)

    # 提取重写文件信息
    rewrite_markers, rewrite_fields, rewrite_lines = extract_markers_and_fields(rewritten_path)
    rewrite_marker_dict = extract_marker_dict(rewrite_markers)

    # 查找缺失的标记
    missing_markers = orig_marker_dict.keys() - rewrite_marker_dict.keys()
    missing_fields = orig_fields - rewrite_fields

    # 查找被修改的标记（标记前缀相同但标记部分不同）
    modified_markers = []
    for marker_key in orig_marker_dict.keys():
        if marker_key in rewrite_marker_dict:
            if orig_marker_dict[marker_key] != rewrite_marker_dict[marker_key]:
                modified_markers.append({
                    'key': marker_key,
                    'original': orig_marker_dict[marker_key],
                    'rewritten': rewrite_marker_dict[marker_key]
                })

    # 输出校验结果
    print(f"=== Markdown格式标记校验结果 ===")
    print(f"原始文件标记数: {len(orig_markers)}")
    print(f"重写文件标记数: {len(rewrite_markers)}")
    print(f"原始文件域代码数: {len(orig_fields)}")
    print(f"重写文件域代码数: {len(rewrite_fields)}")
    print()

    if missing_markers or missing_fields or modified_markers:
        print("❌ 校验失败！发现以下问题：")
        print()

        if missing_markers:
            print(f"【缺失的格式标记】({len(missing_markers)}个):")
            for marker in sorted(missing_markers)[:10]:  # 只显示前10个
                print(f"  - {marker}")
            if len(missing_markers) > 10:
                print(f"  ... 还有 {len(missing_markers) - 10} 个")
            print()

        if missing_fields:
            print(f"【缺失的域代码】({len(missing_fields)}个):")
            for field in sorted(missing_fields)[:10]:
                print(f"  - {field}")
            if len(missing_fields) > 10:
                print(f"  ... 还有 {len(missing_fields) - 10} 个")
            print()

        if modified_markers:
            print(f"【被修改的格式标记】({len(modified_markers)}个):")
            for mod in modified_markers[:10]:
                print(f"  - 原始: {mod['original']}")
                print(f"    重写: {mod['rewritten']}")
            if len(modified_markers) > 10:
                print(f"  ... 还有 {len(modified_markers) - 10} 个")
            print()

        # 如果未指定输出路径，只返回False不修复
        if not output_path:
            print("未指定输出文件路径，不执行自动修复")
            return False

        # 执行修复：使用原始文件的标记，替换为重写文件的内容
        print("🔧 开始自动修复...")
        print()

        # 读取重写文件的所有行
        with open(rewritten_path, 'r', encoding='utf-8') as f:
            rewrite_lines_dict = {}
            for line in rewrite_lines:
                marker_key = line['marker'].split('-->')[0] + '-->'
                rewrite_lines_dict[marker_key] = line['content']

        # 生成修复后的内容
        fixed_lines = []
        for orig_line_info in orig_lines:
            marker_key = orig_line_info['marker'].split('-->')[0] + '-->'

            if marker_key in rewrite_lines_dict:
                # 使用重写后的内容
                new_content = rewrite_lines_dict[marker_key]
                fixed_lines.append(orig_line_info['line'].split('-->')[0] + '-->' + new_content)
            else:
                # 如果重写文件中没有对应标记，保留原始内容
                fixed_lines.append(orig_line_info['line'])

        # 读取原始文件的非标记行
        with open(original_path, 'r', encoding='utf-8') as f:
            original_all_lines = [line.rstrip('\n') for line in f]

        # 识别哪些行是空行或非内容行（标记之外的行）
        original_markers_set = set()
        for line_info in orig_lines:
            original_markers_set.add(line_info['line_num'])

        # 合并：保留原始文件的非内容行，替换有标记的行
        final_lines = []
        fixed_line_idx = 0

        for line_num, line in enumerate(original_all_lines, 1):
            if line_num in original_markers_set and fixed_line_idx < len(fixed_lines):
                # 使用修复后的行
                final_lines.append(fixed_lines[fixed_line_idx])
                fixed_line_idx += 1
            else:
                # 保留原始文件的其他行（空行、表格标记等）
                final_lines.append(line)

        # 写入输出文件
        with open(output_path, 'w', encoding='utf-8') as f:
            for line in final_lines:
                f.write(line + '\n')

        print(f"✅ 修复完成！已生成文件: {output_path}")
        print(f"   - 保留了原始文件的 {len(orig_markers)} 个格式标记")
        print(f"   - 使用了重写文件的文本内容")

        return True
    else:
        print("✅ 校验通过！所有格式标记和域代码都完整保留")
        return True


if __name__ == '__main__':
    if len(sys.argv) < 3:
        print("用法: python validate_markdown.py <原始markdown路径> <重写后的markdown路径> [输出路径]")
        print()
        print("示例:")
        print("  python validate_markdown.py original.md rewritten.md")
        print("  python validate_markdown.py original.md rewritten.md fixed.md")
        print()
        print("说明:")
        print("  - 如果只提供两个参数，仅执行校验，不生成修复文件")
        print("  - 如果提供三个参数，校验失败时会自动生成修复后的文件")
        sys.exit(1)

    original_path = sys.argv[1]
    rewritten_path = sys.argv[2]

    if len(sys.argv) >= 4:
        output_path = sys.argv[3]
    else:
        output_path = None

    # 检查文件是否存在
    if not os.path.exists(original_path):
        print(f"错误: 文件不存在: {original_path}")
        sys.exit(1)

    if not os.path.exists(rewritten_path):
        print(f"错误: 文件不存在: {rewritten_path}")
        sys.exit(1)

    # 执行校验和修复
    result = validate_and_fix(original_path, rewritten_path, output_path)

    # 如果有错误且未修复，返回错误码
    if not result:
        sys.exit(1)
