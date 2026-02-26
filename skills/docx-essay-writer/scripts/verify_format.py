#!/usr/bin/env python3
"""
DOCX Format Verification for docx-essay-writer skill.

Compares an output DOCX against its template to ensure formatting fidelity:
page margins, page size, headers/footers, styles, spacing, etc.

Self-contained — does not import any other skill modules.
"""

import re
import sys
import zipfile
from typing import Any, Dict, List, Tuple


def _read_zip_text(z: zipfile.ZipFile, path: str) -> str:
    try:
        return z.read(path).decode('utf-8')
    except KeyError:
        return ''


def _extract_formatting(doc_xml: str) -> Dict[str, Any]:
    return {
        'pgMar': _first_match(r'<w:pgMar[^>]*>', doc_xml),
        'pgSz': _first_match(r'<w:pgSz[^>]*>', doc_xml),
        'headers': [_normalise_xml_ws(h) for h in re.findall(r'<w:headerReference[^>]*/>', doc_xml)],
        'footers': [_normalise_xml_ws(f) for f in re.findall(r'<w:footerReference[^>]*/>', doc_xml)],
        'spacing_520': len(re.findall(r'w:line="520"', doc_xml)),
        'spacing_500': len(re.findall(r'w:line="500"', doc_xml)),
        'spacing_400': len(re.findall(r'w:line="400"', doc_xml)),
    }


def _first_match(pattern: str, text: str) -> str:
    m = re.search(pattern, text)
    return _normalise_xml_ws(m.group(0)) if m else ''


def _normalise_xml_ws(tag: str) -> str:
    """Normalise whitespace in self-closing XML tags for reliable comparison."""
    return re.sub(r'\s+/>', '/>', tag)


def verify_docx(template_path: str, output_path: str) -> Tuple[bool, List[str]]:
    """
    Compare output DOCX formatting against template.

    Returns (passed, messages).
    """
    msgs: List[str] = []
    passed = True

    try:
        z_tmpl = zipfile.ZipFile(template_path, 'r')
        z_out = zipfile.ZipFile(output_path, 'r')
    except Exception as e:
        return False, [f"✗ Cannot open files: {e}"]

    try:
        # 1. File integrity
        msgs.append("=== 文件完整性检查 ===")
        critical = [
            'word/document.xml',
            'word/styles.xml',
            'word/_rels/document.xml.rels',
        ]
        for header_file in ('word/header1.xml', 'word/header2.xml'):
            if header_file in z_tmpl.namelist():
                critical.append(header_file)
        for footer_file in ('word/footer1.xml', 'word/footer2.xml'):
            if footer_file in z_tmpl.namelist():
                critical.append(footer_file)

        for cf in critical:
            if cf in z_tmpl.namelist() and cf not in z_out.namelist():
                msgs.append(f"✗ 缺少文件: {cf}")
                passed = False
            elif cf in z_out.namelist():
                msgs.append(f"✓ {cf}")

        # 2. Page layout
        msgs.append("\n=== 页面布局检查 ===")
        doc_tmpl = _read_zip_text(z_tmpl, 'word/document.xml')
        doc_out = _read_zip_text(z_out, 'word/document.xml')

        f1 = _extract_formatting(doc_tmpl)
        f2 = _extract_formatting(doc_out)

        layout_checks = [
            ('页边距 (w:pgMar)', 'pgMar'),
            ('页面大小 (w:pgSz)', 'pgSz'),
        ]
        for label, key in layout_checks:
            v1, v2 = f1[key], f2[key]
            ok = v1 == v2
            msgs.append(f"{'✓' if ok else '✗'} {label}")
            msgs.append(f"   模板: {v1 or 'None'}")
            msgs.append(f"   输出: {v2 or 'None'}")
            if not ok:
                passed = False

        # 3. Header / Footer references
        msgs.append("\n=== 页眉页脚引用检查 ===")
        for label, key in [('页眉引用', 'headers'), ('页脚引用', 'footers')]:
            v1, v2 = sorted(f1[key]), sorted(f2[key])
            ok = v1 == v2
            msgs.append(f"{'✓' if ok else '✗'} {label} (数量: 模板={len(v1)}, 输出={len(v2)})")
            if not ok:
                passed = False

        # 4. Header / Footer content
        msgs.append("\n=== 页眉页脚内容检查 ===")
        for pattern, label in [('word/header*.xml', '页眉'), ('word/footer*.xml', '页脚')]:
            for name in z_tmpl.namelist():
                if re.match(r'word/(header|footer)\d+\.xml', name):
                    tmpl_content = _read_zip_text(z_tmpl, name)
                    out_content = _read_zip_text(z_out, name)
                    if not out_content:
                        msgs.append(f"✗ {label} {name} 缺失")
                        passed = False
                    elif tmpl_content == out_content:
                        msgs.append(f"✓ {label} {name} 内容一致")
                    else:
                        msgs.append(f"✗ {label} {name} 内容不同")
                        passed = False

        # 5. Styles.xml preserved
        msgs.append("\n=== 样式文件检查 ===")
        styles_tmpl = _read_zip_text(z_tmpl, 'word/styles.xml')
        styles_out = _read_zip_text(z_out, 'word/styles.xml')
        if styles_tmpl and styles_out:
            if styles_tmpl == styles_out:
                msgs.append("✓ styles.xml 完全一致")
            else:
                msgs.append("⚠ styles.xml 有差异（可能因转换器修改）")
        elif not styles_out:
            msgs.append("✗ styles.xml 缺失")
            passed = False

        # 6. Spacing summary
        msgs.append("\n=== 行距统计 ===")
        for sp_key in ('spacing_520', 'spacing_500', 'spacing_400'):
            label = sp_key.replace('spacing_', '行距')
            v1, v2 = f1[sp_key], f2[sp_key]
            msgs.append(f"{'✓' if v1 == v2 else '⚠'} {label}: 模板={v1}, 输出={v2}")

    finally:
        z_tmpl.close()
        z_out.close()

    return passed, msgs


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

    print(f"\n{'=' * 40}")
    if passed:
        print("✓ 验证通过 - 格式完全匹配！")
    else:
        print("✗ 验证失败 - 存在格式差异")

    sys.exit(0 if passed else 1)


if __name__ == '__main__':
    main()
