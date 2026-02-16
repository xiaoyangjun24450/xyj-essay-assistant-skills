#!/usr/bin/env python3
"""
Verify that output DOCX preserves all formatting from template.
Checks page margins, size, headers, footers, spacing, etc.
"""

import re
import sys
import zipfile
from typing import Dict, List, Tuple


def extract_formatting(doc_xml: str) -> Dict:
    """Extract formatting elements from document.xml."""
    return {
        'pgMar': re.search(r'<w:pgMar[^>]*>', doc_xml),
        'pgSz': re.search(r'<w:pgSz[^>]*>', doc_xml),
        'header': re.search(r'<w:headerReference[^>]*>', doc_xml),
        'footer': re.search(r'<w:footerReference[^>]*>', doc_xml),
        'spacing_520': len(re.findall(r'w:line="520"', doc_xml)),
        'spacing_500': len(re.findall(r'w:line="500"', doc_xml)),
        'spacing_400': len(re.findall(r'w:line="400"', doc_xml)),
    }


def format_value(v):
    """Format value for display."""
    if v is None:
        return "None"
    if hasattr(v, 'group'):
        return v.group(0)
    return str(v)


def verify_docx(template_path: str, output_path: str) -> Tuple[bool, List[str]]:
    """
    Verify output DOCX preserves template formatting.
    
    Returns:
        (success, messages)
    """
    messages = []
    passed = True
    
    try:
        with zipfile.ZipFile(template_path, 'r') as z_tmpl, \
             zipfile.ZipFile(output_path, 'r') as z_out:
            
            # Check critical files exist
            critical_files = [
                'word/document.xml',
                'word/header1.xml',
                'word/footer1.xml',
                'word/styles.xml',
                'word/_rels/document.xml.rels',
            ]
            
            messages.append("=== 文件完整性检查 ===")
            for cf in critical_files:
                if cf not in z_out.namelist():
                    messages.append(f"✗ 缺少文件: {cf}")
                    passed = False
                else:
                    messages.append(f"✓ {cf}")
            
            # Check document formatting
            messages.append("\n=== 格式元素检查 ===")
            doc_tmpl = z_tmpl.read('word/document.xml').decode('utf-8')
            doc_out = z_out.read('word/document.xml').decode('utf-8')
            
            f1 = extract_formatting(doc_tmpl)
            f2 = extract_formatting(doc_out)
            
            checks = [
                ('页边距 (w:pgMar)', 'pgMar'),
                ('页面大小 (w:pgSz)', 'pgSz'),
                ('页眉引用', 'header'),
                ('页脚引用', 'footer'),
                ('行距520段落数', 'spacing_520'),
                ('行距500段落数', 'spacing_500'),
                ('行距400段落数', 'spacing_400'),
            ]
            
            for name, key in checks:
                v1 = f1[key]
                v2 = f2[key]
                
                # Handle regex match objects
                if hasattr(v1, 'group'):
                    v1 = v1.group(0)
                if hasattr(v2, 'group'):
                    v2 = v2.group(0)
                
                match = v1 == v2
                status = '✓' if match else '✗'
                messages.append(f"{status} {name}")
                messages.append(f"   模板: {format_value(v1)}")
                messages.append(f"   输出: {format_value(v2)}")
                
                if not match:
                    passed = False
            
            # Check header/footer content
            messages.append("\n=== 页眉页脚内容检查 ===")
            try:
                h_tmpl = z_tmpl.read('word/header1.xml').decode('utf-8')
                h_out = z_out.read('word/header1.xml').decode('utf-8')
                if h_tmpl == h_out:
                    messages.append("✓ 页眉内容相同")
                else:
                    messages.append("✗ 页眉内容不同")
                    passed = False
            except KeyError:
                messages.append("⚠ 无页眉文件")
            
            try:
                f_tmpl = z_tmpl.read('word/footer1.xml').decode('utf-8')
                f_out = z_out.read('word/footer1.xml').decode('utf-8')
                if f_tmpl == f_out:
                    messages.append("✓ 页脚内容相同")
                else:
                    messages.append("✗ 页脚内容不同")
                    passed = False
            except KeyError:
                messages.append("⚠ 无页脚文件")
            
    except Exception as e:
        messages.append(f"✗ 验证出错: {e}")
        passed = False
    
    return passed, messages


def main():
    if len(sys.argv) < 3:
        print("Usage: python verify_format.py <template.docx> <output.docx>")
        sys.exit(1)
    
    template = sys.argv[1]
    output = sys.argv[2]
    
    print(f"验证中...")
    print(f"模板: {template}")
    print(f"输出: {output}")
    print()
    
    passed, messages = verify_docx(template, output)
    
    for msg in messages:
        print(msg)
    
    print(f"\n{'='*40}")
    if passed:
        print("✓ 验证通过 - 格式完全匹配！")
    else:
        print("✗ 验证失败 - 存在格式差异")
    
    sys.exit(0 if passed else 1)


if __name__ == "__main__":
    main()
