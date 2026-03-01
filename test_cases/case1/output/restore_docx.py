#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
从markdown文件还原到docx模板，保留所有格式和图片

用法:
    python restore_docx.py <markdown文件路径> <docx模板路径> <输出docx路径>

示例:
    python restore_docx.py content.md template.docx output.docx

原理:
    1. 读取原文档作为模板，保留所有格式、图片、表格样式
    2. 根据markdown中的位置标记，替换对应run的文本内容
    3. 保留其他所有内容不变
"""
import sys
import os
import re
import shutil
from docx import Document
from docx.shared import Pt, RGBColor
from docx.oxml.ns import qn


def string_to_format(format_string):
    """将字符串转换为格式信息"""
    if not format_string:
        return {}
    format_info = {}
    for part in format_string.split('|'):
        if '=' in part:
            key, value = part.split('=', 1)
            # 转换布尔值
            if value == 'True':
                value = True
            elif value == 'False':
                value = False
            # 转换数字
            elif value.replace('.', '').replace('-', '').isdigit():
                try:
                    value = float(value) if '.' in value else int(value)
                except:
                    pass
            format_info[key] = value
    return format_info


def parse_markdown_line(line):
    """解析markdown行，提取位置标记、格式和内容"""
    # 匹配段落标记
    para_match = re.match(r'<!-- PARA:(\d+),RUN:(\d+),FORMAT:(.*?) -->(.*)', line)
    if para_match:
        return {
            'type': 'paragraph',
            'para_idx': int(para_match.group(1)),
            'run_idx': int(para_match.group(2)),
            'format': string_to_format(para_match.group(3)),
            'content': para_match.group(4)
        }

    # 匹配表格标记
    table_match = re.match(r'<!-- TABLE:(\d+),ROW:(\d+),CELL:(\d+),PARA:(\d+),RUN:(\d+),FORMAT:(.*?) -->(.*)', line)
    if table_match:
        return {
            'type': 'table_cell',
            'table_idx': int(table_match.group(1)),
            'row_idx': int(table_match.group(2)),
            'cell_idx': int(table_match.group(3)),
            'para_idx': int(table_match.group(4)),
            'run_idx': int(table_match.group(5)),
            'format': string_to_format(table_match.group(6)),
            'content': table_match.group(7)
        }

    # 匹配表格标记
    if line.strip() == '<!-- TABLE -->':
        return {'type': 'table_start'}

    if line.strip() == '<!-- TABLE_END -->':
        return {'type': 'table_end'}

    # 普通空行
    if line.strip() == '':
        return {'type': 'empty'}

    return None


def restore_docx_from_markdown(markdown_path, template_path, output_path):
    """从markdown文件还原到docx模板"""
    # 先复制模板文件
    shutil.copy2(template_path, output_path)
    doc = Document(output_path)

    # 读取markdown文件
    with open(markdown_path, 'r', encoding='utf-8') as f:
        markdown_lines = [line.rstrip('\n') for line in f]

    # 解析markdown
    parsed_items = []
    for line in markdown_lines:
        item = parse_markdown_line(line)
        if item:
            parsed_items.append(item)

    # 统计
    paragraph_changes = 0
    table_changes = 0

    # 首先清除所有段落和表格的内容（只保留格式和图片）
    def has_non_text_elements(run):
        """检查run是否包含非文本元素（如图片）"""
        for child in run._element:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag
            if tag not in ['t', 'rPr', 'tab', 'br']:  # t是文本元素，rPr是格式，tab是制表符，br是换行
                return True
        return False

    for para in doc.paragraphs:
        for run in para.runs:
            # 只清空有文本内容的run，保留包含图片等元素的run
            if not has_non_text_elements(run):
                run.text = ""

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for run in para.runs:
                        # 只清空有文本内容的run，保留包含图片等元素的run
                        if not has_non_text_elements(run):
                            run.text = ""

    # 还原段落内容
    for item in parsed_items:
        if item['type'] == 'paragraph':
            para_idx = item['para_idx']
            run_idx = item['run_idx']
            content = item['content']

            if para_idx < len(doc.paragraphs):
                para = doc.paragraphs[para_idx]
                if run_idx < len(para.runs):
                    # 保留原始run，只修改文本
                    para.runs[run_idx].text = content
                    paragraph_changes += 1

    # 还原表格内容
    for item in parsed_items:
        if item['type'] == 'table_cell':
            table_idx = item['table_idx']
            row_idx = item['row_idx']
            cell_idx = item['cell_idx']
            para_idx = item['para_idx']
            run_idx = item['run_idx']
            content = item['content']

            if table_idx < len(doc.tables):
                table = doc.tables[table_idx]
                if row_idx < len(table.rows):
                    row = table.rows[row_idx]
                    if cell_idx < len(row.cells):
                        cell = row.cells[cell_idx]
                        if para_idx < len(cell.paragraphs):
                            para = cell.paragraphs[para_idx]
                            if run_idx < len(para.runs):
                                # 保留原始run，只修改文本
                                para.runs[run_idx].text = content
                                table_changes += 1

    # 保存文档
    doc.save(output_path)
    print(f"文档已还原到: {output_path}")
    print(f"共修改了 {paragraph_changes} 个段落中的run")
    print(f"共修改了 {table_changes} 个表格单元格中的run")
    print("所有格式、图片和样式都已保留")


if __name__ == '__main__':
    # 解析命令行参数
    if len(sys.argv) < 3:
        print("用法: python restore_docx.py <markdown文件路径> <docx模板路径> [输出docx路径]")
        print("示例:")
        print("  python restore_docx.py content.md template.docx output.docx")
        print("  python restore_docx.py content.md template.docx")
        sys.exit(1)

    markdown_path = sys.argv[1]
    template_path = sys.argv[2]

    if len(sys.argv) >= 4:
        output_path = sys.argv[3]
    else:
        # 如果未指定输出路径，使用与markdown相同的文件名，但扩展名改为.docx
        base_name = os.path.splitext(markdown_path)[0]
        output_path = base_name + '_restored.docx'

    # 检查文件是否存在
    if not os.path.exists(markdown_path):
        print(f"错误: 文件不存在: {markdown_path}")
        sys.exit(1)

    if not os.path.exists(template_path):
        print(f"错误: 文件不存在: {template_path}")
        sys.exit(1)

    # 还原文档
    restore_docx_from_markdown(markdown_path, template_path, output_path)
