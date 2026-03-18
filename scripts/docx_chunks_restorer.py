#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx_chunks_restorer.py - 从 chunks 和 format_registry.json 还原 DOCX

根据 chunks/ 目录中的文本内容修改 unzipped/ 中的 document.xml，
将修改后的内容重新打包成 docx 文件。

还原规则：
1. 公式（$...$ 或 $$...$$）使用 LatexToOmmlConverter 转换为 OMML
2. 文字格式通过查找 format_registry.json 中的类别 ID 获取对应的 rPr XML

用法:
    python docx_chunks_restorer.py <unzipped_dir> <chunks_dir> <output_docx>

示例:
    python docx_chunks_restorer.py output/unzipped output/chunks restored.docx
"""

import copy
import json
import os
import re
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import Dict, List, Optional, Tuple, Any


# ====================================================================
# Namespace constants
# ====================================================================

NS_W   = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
NS_W14 = 'http://schemas.microsoft.com/office/word/2010/wordml'
NS_M   = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
NS_R   = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'

_ALL_NS = {
    'w':           NS_W,
    'w14':         NS_W14,
    'm':           NS_M,
    'r':           NS_R,
    'wp':          'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a':           'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic':         'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'mc':          'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'wpc':         'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
    'o':           'urn:schemas-microsoft-com:office:office',
    'v':           'urn:schemas-microsoft-com:vml',
    'w15':         'http://schemas.microsoft.com/office/word/2012/wordml',
    'wne':         'http://schemas.microsoft.com/office/word/2006/wordml',
    'w10':         'urn:schemas-microsoft-com:office:word',
    'wps':         'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
    'wpg':         'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
    'wpi':         'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
    'wp14':        'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
    'wpsCustomData': 'http://www.wps.cn/officeDocument/2013/wpsCustomData',
}

# Clark-notation shortcuts
_W   = f'{{{NS_W}}}'
_W14 = f'{{{NS_W14}}}'
_M   = f'{{{NS_M}}}'


def _register_namespaces():
    for prefix, uri in _ALL_NS.items():
        ET.register_namespace(prefix, uri)


# ====================================================================
# LaTeX → OMML Converter
# ====================================================================

class LatexToOmmlConverter:
    """将 LaTeX 数学表达式转换为 OMML XML 元素。"""

    def __init__(self):
        self.greek_map = {
            'alpha': 'α', 'beta': 'β', 'gamma': 'γ', 'delta': 'δ',
            'epsilon': 'ε', 'zeta': 'ζ', 'eta': 'η', 'theta': 'θ',
            'iota': 'ι', 'kappa': 'κ', 'lambda': 'λ', 'mu': 'μ',
            'nu': 'ν', 'xi': 'ξ', 'pi': 'π', 'rho': 'ρ',
            'sigma': 'σ', 'tau': 'τ', 'upsilon': 'υ', 'phi': 'φ',
            'chi': 'χ', 'psi': 'ψ', 'omega': 'ω',
            'Gamma': 'Γ', 'Delta': 'Δ', 'Theta': 'Θ', 'Lambda': 'Λ',
            'Xi': 'Ξ', 'Pi': 'Π', 'Sigma': 'Σ', 'Phi': 'Φ',
            'Psi': 'Ψ', 'Omega': 'Ω',
        }

    @staticmethod
    def _extract_brace(s: str, pos: int):
        """从 pos（应指向 '{'）起提取平衡花括号内容。"""
        if pos >= len(s) or s[pos] != '{':
            return '', pos
        depth, i = 0, pos
        while i < len(s):
            if s[i] == '{':
                depth += 1
            elif s[i] == '}':
                depth -= 1
                if depth == 0:
                    return s[pos + 1:i], i + 1
            i += 1
        return s[pos + 1:], len(s)

    def _ctrl_pr(self):
        ctrlPr = ET.Element(f'{_M}ctrlPr')
        rPr = ET.SubElement(ctrlPr, f'{_W}rPr')
        rFonts = ET.SubElement(rPr, f'{_W}rFonts')
        rFonts.set(f'{_W}hint', 'eastAsia')
        rFonts.set(f'{_W}ascii', 'Cambria Math')
        rFonts.set(f'{_W}hAnsi', 'Cambria Math')
        return ctrlPr

    def _math_run(self, text, hint='default'):
        r = ET.Element(f'{_M}r')
        rPr = ET.SubElement(r, f'{_M}rPr')
        sty = ET.SubElement(rPr, f'{_M}sty')
        sty.set(f'{_M}val', 'p')
        wrPr = ET.SubElement(r, f'{_W}rPr')
        rFonts = ET.SubElement(wrPr, f'{_W}rFonts')
        rFonts.set(f'{_W}hint', hint)
        rFonts.set(f'{_W}ascii', 'Cambria Math')
        rFonts.set(f'{_W}hAnsi', 'Cambria Math')
        t = ET.SubElement(r, f'{_M}t')
        t.text = text
        return r

    def _math_run_ea(self, text):
        return self._math_run(text, hint='eastAsia')

    _XML_ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')

    def convert(self, latex: str) -> ET.Element:
        latex = self._XML_ILLEGAL.sub('', latex.strip())
        if not latex:
            omath = ET.Element(f'{_M}oMath')
            return omath
        if latex.startswith('$$') and latex.endswith('$$'):
            latex = latex[2:-2].strip()
        elif latex.startswith('$') and latex.endswith('$'):
            latex = latex[1:-1].strip()
        omath = ET.Element(f'{_M}oMath')
        self._parse_expr(omath, latex)
        return omath

    def _parse_expr(self, parent, expr):
        expr = expr.strip()
        if not expr:
            return

        _UC = r'[a-zA-Z0-9\u0080-\uffff]'

        # 1. \func(...)
        m = re.match(r'\\(arctan|arcsin|arccos|cos|sin|tan|cot|log|ln|exp|lim|max|min|sup|inf)\s*\(([^)]+)\)', expr)
        if m:
            parent.append(self._math_run(m.group(1)))
            d = ET.SubElement(parent, f'{_M}d')
            dPr = ET.SubElement(d, f'{_M}dPr')
            ET.SubElement(dPr, f'{_M}begChr').set(f'{_M}val', '(')
            ET.SubElement(dPr, f'{_M}endChr').set(f'{_M}val', ')')
            dPr.append(self._ctrl_pr())
            e = ET.SubElement(d, f'{_M}e')
            self._parse_expr(e, m.group(2))
            e.append(self._ctrl_pr())
            d.append(self._ctrl_pr())
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 2. Parentheses with complex inner content, optionally followed by ^{...}
        m = re.match(r'\(([^)]+)\)(\^\{)', expr)
        if m:
            inner = m.group(1)
            pos = m.end() - 1
            sup_content, end_pos = self._extract_brace(expr, pos)
            sup_el = ET.SubElement(parent, f'{_M}sSup')
            sup_pr = ET.SubElement(sup_el, f'{_M}sSupPr')
            sup_pr.append(self._ctrl_pr())
            sup_e = ET.SubElement(sup_el, f'{_M}e')
            if '\\' in inner:
                d = ET.SubElement(sup_e, f'{_M}d')
                dPr = ET.SubElement(d, f'{_M}dPr')
                ET.SubElement(dPr, f'{_M}begChr').set(f'{_M}val', '(')
                ET.SubElement(dPr, f'{_M}endChr').set(f'{_M}val', ')')
                dPr.append(self._ctrl_pr())
                e = ET.SubElement(d, f'{_M}e')
                self._parse_expr(e, inner)
                e.append(self._ctrl_pr())
                d.append(self._ctrl_pr())
            else:
                self._parse_expr(sup_e, '(' + inner + ')')
            sup_e.append(self._ctrl_pr())
            sup = ET.SubElement(sup_el, f'{_M}sup')
            self._parse_expr(sup, sup_content)
            sup.append(self._ctrl_pr())
            sup_el.append(self._ctrl_pr())
            rest = expr[end_pos:]
            if rest:
                self._parse_expr(parent, rest)
            return
        
        # 2b. Parentheses without superscript
        m = re.match(r'\(([^)]+)\)', expr)
        if m:
            inner = m.group(1)
            if '\\' in inner:
                d = ET.SubElement(parent, f'{_M}d')
                dPr = ET.SubElement(d, f'{_M}dPr')
                ET.SubElement(dPr, f'{_M}begChr').set(f'{_M}val', '(')
                ET.SubElement(dPr, f'{_M}endChr').set(f'{_M}val', ')')
                dPr.append(self._ctrl_pr())
                e = ET.SubElement(d, f'{_M}e')
                self._parse_expr(e, inner)
                e.append(self._ctrl_pr())
                d.append(self._ctrl_pr())
            else:
                parent.append(self._math_run('('))
                self._parse_expr(parent, inner)
                parent.append(self._math_run(')'))
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 3. Matrix \begin{bmatrix}...\end{bmatrix}
        m = re.match(r'\\begin\{bmatrix\}(.*?)\\end\{bmatrix\}', expr, re.DOTALL)
        if m:
            self._build_matrix(parent, m.group(1), '[', ']')
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 4. Cases \begin{cases}...\end{cases}
        m = re.match(r'\\begin\{cases\}(.*?)\\end\{cases\}', expr, re.DOTALL)
        if m:
            content = m.group(1).strip()
            rows = [r.strip() for r in content.split('\\\\') if r.strip()]
            d = ET.SubElement(parent, f'{_M}d')
            dPr = ET.SubElement(d, f'{_M}dPr')
            ET.SubElement(dPr, f'{_M}begChr').set(f'{_M}val', '{')
            ET.SubElement(dPr, f'{_M}endChr').set(f'{_M}val', '')
            dPr.append(self._ctrl_pr())
            e = ET.SubElement(d, f'{_M}e')
            for i, row in enumerate(rows):
                if i > 0:
                    e.append(self._math_run(' '))
                part = row.split('&')[0].strip() if '&' in row else row
                self._parse_expr(e, part)
            e.append(self._ctrl_pr())
            d.append(self._ctrl_pr())
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 4.5 \sqrt{...}
        if expr.startswith('\\sqrt{'):
            content, pos = self._extract_brace(expr, 5)
            rad = ET.SubElement(parent, f'{_M}rad')
            radPr = ET.SubElement(rad, f'{_M}radPr')
            ET.SubElement(radPr, f'{_M}degHide').set(f'{_M}val', '1')
            radPr.append(self._ctrl_pr())
            deg = ET.SubElement(rad, f'{_M}deg')
            deg.append(self._ctrl_pr())
            e = ET.SubElement(rad, f'{_M}e')
            self._parse_expr(e, content)
            e.append(self._ctrl_pr())
            rest = expr[pos:]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 4.6 \vec{...}
        if expr.startswith('\\vec{'):
            content, pos = self._extract_brace(expr, 4)
            acc = ET.SubElement(parent, f'{_M}acc')
            accPr = ET.SubElement(acc, f'{_M}accPr')
            ET.SubElement(accPr, f'{_M}chr').set(f'{_M}val', '\u20d7')
            accPr.append(self._ctrl_pr())
            e = ET.SubElement(acc, f'{_M}e')
            self._parse_expr(e, content)
            e.append(self._ctrl_pr())
            rest = expr[pos:]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 4.7 \text{...}
        if expr.startswith('\\text{'):
            content, pos = self._extract_brace(expr, 5)
            parent.append(self._math_run(content))
            rest = expr[pos:]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 4.8 \tag{...}
        if expr.startswith('\\tag{'):
            content, pos = self._extract_brace(expr, 4)
            parent.append(self._math_run(f'  ({content})'))
            rest = expr[pos:]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 5. Fractions \frac{a}{b}
        if expr.startswith('\\frac{'):
            num_content, pos = self._extract_brace(expr, 5)
            if pos < len(expr) and expr[pos] == '{':
                den_content, pos = self._extract_brace(expr, pos)
                f_el = ET.SubElement(parent, f'{_M}f')
                fPr = ET.SubElement(f_el, f'{_M}fPr')
                fPr.append(self._ctrl_pr())
                num_el = ET.SubElement(f_el, f'{_M}num')
                self._parse_expr(num_el, num_content)
                num_el.append(self._ctrl_pr())
                den_el = ET.SubElement(f_el, f'{_M}den')
                self._parse_expr(den_el, den_content)
                den_el.append(self._ctrl_pr())
                parent.append(self._ctrl_pr())
                rest = expr[pos:]
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 6. Combined sub+sup
        combined_pats = [
            r'\\([a-zA-Z]+)_\{([^}]+)\}\^\{([^}]+)\}',
            fr'\\([a-zA-Z]+)_({_UC})\^({_UC})',
            fr'({_UC})_\\([a-zA-Z]+)\^\{{([^}}]+)\}}',
            fr'({_UC})_\\([a-zA-Z]+)\^({_UC})',
            fr'({_UC})_\{{([^}}]+)\}}\^\{{([^}}]+)\}}',
            fr'({_UC})_({_UC})\^({_UC})',
        ]
        for pat in combined_pats:
            m = re.match(pat, expr)
            if m:
                self._build_subsup(parent, m.group(1), m.group(2), m.group(3))
                rest = expr[m.end():]
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 7. Subscript
        sub_pats_simple = [
            r'\\([a-zA-Z]+)_\{([^}]+)\}',
            fr'\\([a-zA-Z]+)_({_UC})',
            fr'({_UC})_\\([a-zA-Z]+)',
            f'({_UC})_\\{{([^}}]+)\\}}',
            fr'({_UC})_({_UC})',
        ]
        for pat in sub_pats_simple:
            m = re.match(pat, expr)
            if m:
                self._build_sub(parent, m.group(1), m.group(2))
                rest = expr[m.end():]
                if rest:
                    self._parse_expr(parent, rest)
                return
        
        m = re.match(r'([∥‖\(\)\[\]\{\}])_\{', expr)
        if m:
            base = m.group(1)
            pos = m.end() - 1
            sub_content, end_pos = self._extract_brace(expr, pos)
            self._build_sub(parent, base, sub_content)
            rest = expr[end_pos:]
            if rest:
                self._parse_expr(parent, rest)
            return
        
        if expr.startswith('_{'):
            pos = 1
            sub_content, end_pos = self._extract_brace(expr, pos)
            last_child = None
            if len(parent) > 0:
                last_child = parent[-1]
                parent.remove(last_child)
            sub_el = ET.SubElement(parent, f'{_M}sSub')
            sub_pr = ET.SubElement(sub_el, f'{_M}sSubPr')
            sub_pr.append(self._ctrl_pr())
            sub_e = ET.SubElement(sub_el, f'{_M}e')
            if last_child is not None:
                sub_e.append(last_child)
            sub_e.append(self._ctrl_pr())
            sub = ET.SubElement(sub_el, f'{_M}sub')
            self._parse_expr(sub, sub_content)
            sub.append(self._ctrl_pr())
            sub_el.append(self._ctrl_pr())
            rest = expr[end_pos:]
            if rest:
                self._parse_expr(parent, rest)
            return
        
        m = re.match(r'([a-zA-Z]+)_', expr)
        if m:
            base = m.group(1)
            pos = m.end()
            if pos < len(expr) and expr[pos] == '{':
                sub_content, end_pos = self._extract_brace(expr, pos)
                self._build_sub(parent, base, sub_content)
                rest = expr[end_pos:]
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 8. Superscript
        sup_pats_simple = [
            r'\\([a-zA-Z]+)\^\{([^}]+)\}',
            fr'\\([a-zA-Z]+)\^({_UC})',
            fr'({_UC})^\{{([^}}]+)\}}',
            fr'({_UC})\^({_UC})',
        ]
        for pat in sup_pats_simple:
            m = re.match(pat, expr)
            if m:
                self._build_sup(parent, m.group(1), m.group(2))
                rest = expr[m.end():]
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 9. Greek letters + other known backslash commands
        m = re.match(r'\\([a-zA-Z]+)', expr)
        if m:
            name = m.group(1)
            rest = expr[m.end():]
            if name in self.greek_map:
                parent.append(self._math_run_ea(self.greek_map[name]))
            elif name == 'quad':
                parent.append(self._math_run('\u2003'))
            elif name in ('cdot', 'times'):
                parent.append(self._math_run('⋅'))
            elif name in ('arctan', 'arcsin', 'arccos', 'cot', 'sec', 'csc',
                          'sin', 'cos', 'tan', 'log', 'ln', 'exp',
                          'lim', 'max', 'min', 'sup', 'inf'):
                parent.append(self._math_run(name))
            elif name == 'sqrt':
                if rest:
                    rad = ET.SubElement(parent, f'{_M}rad')
                    radPr = ET.SubElement(rad, f'{_M}radPr')
                    ET.SubElement(radPr, f'{_M}degHide').set(f'{_M}val', '1')
                    radPr.append(self._ctrl_pr())
                    deg = ET.SubElement(rad, f'{_M}deg')
                    deg.append(self._ctrl_pr())
                    e = ET.SubElement(rad, f'{_M}e')
                    self._parse_expr(e, rest[0])
                    e.append(self._ctrl_pr())
                    rest = rest[1:]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 10. Operators & symbols
        if expr[0] in '=+-*/()[]{},;: |\'':
            parent.append(self._math_run(expr[0]))
            if len(expr) > 1:
                self._parse_expr(parent, expr[1:])
            return

        # 11. Regular text/numbers
        m = re.match(fr'({_UC}+)', expr)
        if m:
            token = m.group(1)
            if len(token) == 1 and token in self.greek_map.values():
                parent.append(self._math_run_ea(token))
            else:
                parent.append(self._math_run(token))
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        if len(expr) > 1:
            self._parse_expr(parent, expr[1:])

    def _resolve(self, token):
        if token.startswith('\\') and token[1:] in self.greek_map:
            return self._math_run_ea(self.greek_map[token[1:]])
        if token in self.greek_map:
            return self._math_run_ea(self.greek_map[token])
        return self._math_run(token)

    def _build_matrix(self, parent, content, beg, end):
        rows = [r.strip() for r in content.strip().split('\\\\') if r.strip()]
        d = ET.SubElement(parent, f'{_M}d')
        dPr = ET.SubElement(d, f'{_M}dPr')
        ET.SubElement(dPr, f'{_M}begChr').set(f'{_M}val', beg)
        ET.SubElement(dPr, f'{_M}endChr').set(f'{_M}val', end)
        dPr.append(self._ctrl_pr())
        e = ET.SubElement(d, f'{_M}e')
        mat = ET.SubElement(e, f'{_M}m')
        mPr = ET.SubElement(mat, f'{_M}mPr')
        mcs = ET.SubElement(mPr, f'{_M}mcs')
        mc = ET.SubElement(mcs, f'{_M}mc')
        mcPr = ET.SubElement(mc, f'{_M}mcPr')
        ET.SubElement(mcPr, f'{_M}count').set(f'{_M}val', '1')
        ET.SubElement(mcPr, f'{_M}mcJc').set(f'{_M}val', 'center')
        ET.SubElement(mPr, f'{_M}plcHide').set(f'{_M}val', '1')
        mPr.append(self._ctrl_pr())
        for row in rows:
            mr = ET.SubElement(mat, f'{_M}mr')
            cells = [c.strip() for c in row.split('&')]
            for cell in cells:
                ec = ET.SubElement(mr, f'{_M}e')
                self._parse_expr(ec, cell)
                ec.append(self._ctrl_pr())
        e.append(self._ctrl_pr())
        d.append(self._ctrl_pr())

    def _build_subsup(self, parent, base, sub_val, sup_val):
        el = ET.SubElement(parent, f'{_M}sSubSup')
        pr = ET.SubElement(el, f'{_M}sSubSupPr')
        pr.append(self._ctrl_pr())
        e = ET.SubElement(el, f'{_M}e')
        e.append(self._resolve(base))
        e.append(self._ctrl_pr())
        sub = ET.SubElement(el, f'{_M}sub')
        self._parse_expr(sub, sub_val)
        sub.append(self._ctrl_pr())
        sup = ET.SubElement(el, f'{_M}sup')
        self._parse_expr(sup, sup_val)
        sup.append(self._ctrl_pr())
        el.append(self._ctrl_pr())

    def _build_sub(self, parent, base, sub_val):
        el = ET.SubElement(parent, f'{_M}sSub')
        pr = ET.SubElement(el, f'{_M}sSubPr')
        pr.append(self._ctrl_pr())
        e = ET.SubElement(el, f'{_M}e')
        e.append(self._resolve(base))
        e.append(self._ctrl_pr())
        sub = ET.SubElement(el, f'{_M}sub')
        self._parse_expr(sub, sub_val)
        sub.append(self._ctrl_pr())
        el.append(self._ctrl_pr())

    def _build_sup(self, parent, base, sup_val):
        el = ET.SubElement(parent, f'{_M}sSup')
        pr = ET.SubElement(el, f'{_M}sSupPr')
        pr.append(self._ctrl_pr())
        e = ET.SubElement(el, f'{_M}e')
        e.append(self._resolve(base))
        e.append(self._ctrl_pr())
        sup = ET.SubElement(el, f'{_M}sup')
        self._parse_expr(sup, sup_val)
        sup.append(self._ctrl_pr())
        el.append(self._ctrl_pr())


# ====================================================================
# Chunks Parser
# ====================================================================

def parse_chunks(chunks_dir: Path) -> Dict[str, str]:
    """
    解析 chunks 目录中的所有文件，返回 para_id -> text 的映射
    """
    para_map = {}
    chunk_files = sorted(chunks_dir.glob('chunk_*.md'))
    
    for chunk_file in chunk_files:
        with open(chunk_file, 'r', encoding='utf-8') as f:
            for line in f:
                # 只移除行首空格，保留行尾空格（用于保留段落末尾的空格格式）
                line = line.lstrip()
                if not line:
                    continue
                m = re.match(r'\[([A-Fa-f0-9]+)\]\s+(.*)', line, re.DOTALL)
                if m:
                    para_id = m.group(1).upper()
                    text = m.group(2).rstrip('\n')
                    para_map[para_id] = text
    
    return para_map


# ====================================================================
# Format Registry Loader
# ====================================================================

class FormatRegistry:
    """格式注册表，管理类别ID到rPr XML的映射（支持每段落独立格式）"""
    
    def __init__(self, registry_path: Path):
        self.baseline_id: str = 'F0'
        self.categories: Dict[str, Dict[str, Any]] = {}
        self._rpr_cache: Dict[str, Optional[ET.Element]] = {}
        self.para_registries: Dict[str, Dict[str, Any]] = {}  # para_id -> {baseline, categories}
        self.default_categories: Dict[str, Dict[str, Any]] = {}  # 全局默认类别
        
        if registry_path.exists():
            with open(registry_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
                # 新格式：每个段落有自己的注册表
                if 'paragraphs' in data:
                    self.para_registries = data['paragraphs']
                    print(f"已加载段落格式注册表: {len(self.para_registries)} 个段落")
                # 兼容旧格式
                elif 'categories' in data:
                    self.default_categories = data.get('categories', {})
                    self.baseline_id = data.get('baseline', 'F0')
                    print(f"已加载格式注册表: {len(self.default_categories)} 个类别")
                    print(f"基准类别: {self.baseline_id}")
        else:
            print(f"警告: 未找到格式注册表 {registry_path}")
    
    def get_para_registry(self, para_id: str) -> Optional[Dict[str, Any]]:
        """获取指定段落的格式注册表"""
        return self.para_registries.get(para_id)
    
    def get_rpr(self, cat_id: str, para_id: str = None) -> Optional[ET.Element]:
        """获取类别ID对应的rPr元素（优先使用段落注册表）"""
        # 先尝试从段落注册表获取
        if para_id and para_id in self.para_registries:
            para_reg = self.para_registries[para_id]
            categories = para_reg.get('categories', {})
            cache_key = f"{para_id}:{cat_id}"
            
            if cache_key in self._rpr_cache:
                return self._rpr_cache[cache_key]
            
            if cat_id not in categories:
                return None
            
            rpr_xml = categories[cat_id].get('rpr_xml')
            if not rpr_xml:
                return None
            
            try:
                rpr_elem = ET.fromstring(rpr_xml)
                self._rpr_cache[cache_key] = rpr_elem
                return rpr_elem
            except ET.ParseError as e:
                print(f"警告: 解析rPr XML失败 ({cat_id}): {e}")
                return None
        
        # 回退到全局注册表
        cache_key = f"global:{cat_id}"
        if cache_key in self._rpr_cache:
            return self._rpr_cache[cache_key]
        
        if cat_id not in self.default_categories:
            return None
        
        rpr_xml = self.default_categories[cat_id].get('rpr_xml')
        if not rpr_xml:
            return None
        
        try:
            rpr_elem = ET.fromstring(rpr_xml)
            self._rpr_cache[cache_key] = rpr_elem
            return rpr_elem
        except ET.ParseError as e:
            print(f"警告: 解析rPr XML失败 ({cat_id}): {e}")
            return None
    
    def get_baseline_rpr(self, para_id: str = None) -> Optional[ET.Element]:
        """获取基准类别的rPr"""
        # 优先使用段落基准
        if para_id and para_id in self.para_registries:
            baseline_id = self.para_registries[para_id].get('baseline', 'F0')
            return self.get_rpr(baseline_id, para_id)
        
        # 回退到全局基准
        return self.get_rpr(self.baseline_id)


# ====================================================================
# Main Restorer
# ====================================================================

class DocxChunksRestorer:
    """
    根据 chunks 和 format_registry.json 还原 word/document.xml 并重新打包成 docx。
    """

    def __init__(self, unzipped_dir: str, chunks_dir: str):
        self.unzipped_dir = Path(unzipped_dir)
        self.chunks_dir = Path(chunks_dir)
        self._converter = LatexToOmmlConverter()
        
        # 加载格式注册表
        registry_path = self.chunks_dir.parent / 'format_registry.json'
        self._registry = FormatRegistry(registry_path)

    def restore(self, output_docx_path: str):
        """执行还原并输出 docx。"""
        para_map = parse_chunks(self.chunks_dir)
        print(f"从 chunks 解析到 {len(para_map)} 个段落")

        _register_namespaces()

        doc_path = self.unzipped_dir / 'word' / 'document.xml'
        tree = ET.parse(str(doc_path))
        root = tree.getroot()

        # 处理所有段落
        for para in root.iter(f'{_W}p'):
            para_id = para.get(f'{_W14}paraId')
            if para_id:
                para_id_upper = para_id.upper()
                if para_id_upper in para_map:
                    self._restore_paragraph(para, para_map[para_id_upper], para_id_upper)

        xml_str = (
            '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n'
            + ET.tostring(root, encoding='unicode')
        )

        self._pack_docx(output_docx_path, override_files={'word/document.xml': xml_str.encode('utf-8')})
        print(f"还原完成: {output_docx_path}")

    def _is_formula_only(self, text: str) -> Tuple[bool, str, Optional[str]]:
        """判断文本是否为单个 $...$ 或 $$...$$ 公式"""
        s = text.strip()
        m = re.match(r'(\$\$[^$].*?\$\$|\$[^$]+\$)(\s*\([^)]*\))?$', s)
        if not m:
            return False, text, None
        
        formula = m.group(1)
        tag = m.group(2).strip() if m.group(2) else None
        
        inner = formula[1:-1] if formula.startswith('$') and not formula.startswith('$$') else formula[2:-2]
        if inner.count('$') > 0:
            return False, text, None
            
        return True, formula, tag

    def _restore_paragraph(self, para: ET.Element, new_text: str, para_id: str = None):
        """根据新文本重建段落。"""
        # 收集原始 runs
        original_runs: List[ET.Element] = [
            c for c in para if c.tag == f'{_W}r'
        ]

        # 结构性 tab runs
        tab_runs: List[ET.Element] = [
            r for r in original_runs
            if r.find(f'{_W}tab') is not None and r.find(f'{_W}t') is None
        ]

        # 保留非文本元素
        _IMG_TAGS = {f'{_W}pict', f'{_W}drawing', f'{_W}object'}
        _MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
        keepers = []
        
        for child in para:
            if child.tag == f'{_W}pPr':
                keepers.append(('ppr', child))
                continue
            if child.tag == f'{_W}bookmarkStart' or child.tag == f'{_W}bookmarkEnd':
                keepers.append(('bookmark', child))
                continue
            if child.tag == f'{_W}r':
                if child.find(f'{_W}t') is None:
                    if any(c.tag in _IMG_TAGS for c in child):
                        keepers.append(('image', child))
                        continue
            if child.tag == f'{{{_MC_NS}}}AlternateContent':
                keepers.append(('alt', child))
                continue

        # 清空段落内容
        for child in list(para):
            para.remove(child)

        # 重建段落：先保留 pPr
        for kind, elem in keepers:
            if kind == 'ppr':
                para.append(elem)
                break

        # 获取基准rPr（使用段落注册表）
        baseline_rpr = self._registry.get_baseline_rpr(para_id)

        # 是否公式段落
        is_formula, formula_part, tag_text = self._is_formula_only(new_text)
        is_formula_para = bool(tab_runs) and is_formula

        # 重建内容
        if is_formula_para:
            # 公式段落
            if tab_runs:
                para.append(copy.deepcopy(tab_runs[0]))

            inner = formula_part.strip()
            if inner.startswith('$$'):
                inner = inner[2:-2].strip()
            else:
                inner = inner[1:-1].strip()

            tag_m = re.search(r'\\tag\{([^}]+)\}', inner)
            if tag_m:
                tag_text = f'({tag_m.group(1)})'
                inner = inner[:tag_m.start()].rstrip()

            para.append(self._converter.convert(inner))

            if len(tab_runs) > 1 and tag_text:
                para.append(copy.deepcopy(tab_runs[1]))

            if tag_text:
                para.append(self._make_text_run(tag_text, baseline_rpr))
        else:
            # 普通段落（含内联公式和格式类别标签）
            parts = re.split(r'(\$\$[^$].*?\$\$|\$[^$]+\$)', new_text, flags=re.DOTALL)
            
            for part in parts:
                if not part:
                    continue
                
                # 处理 $$...$$ 格式
                if part.startswith('$$') and part.endswith('$$') and len(part) > 4:
                    formula = part[2:-2].strip()
                    para.append(self._converter.convert(formula))
                # 处理 $...$ 格式
                elif part.startswith('$') and part.endswith('$') and len(part) > 2:
                    formula = part[1:-1].strip()
                    para.append(self._converter.convert(formula))
                else:
                    # 普通文本 - 解析格式类别标签
                    if part:
                        segments = self._parse_category_tags(part)
                        for text_segment, cat_id in segments:
                            if text_segment:
                                rpr = self._registry.get_rpr(cat_id, para_id) if cat_id else baseline_rpr
                                para.append(self._make_text_run(text_segment, rpr))

        # 在末尾添加保留的书签和图片
        for kind, elem in keepers:
            if kind != 'ppr':
                para.append(elem)

    _XML_ILLEGAL = re.compile(r'[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]')

    def _parse_category_tags(self, text: str) -> List[Tuple[str, Optional[str]]]:
        """
        解析格式类别标签，返回 [(文本片段, 类别ID), ...]
        格式标签示例: ‹F1:u›文本内容‹/›
        基准文本无标签
        """
        result = []
        pos = 0
        
        # 匹配格式类别标签: ‹F数字:提示›内容‹/›
        # 使用非贪婪匹配，支持任意提示内容
        tag_pattern = r'‹(F\d+):[^›]*›([^‹]*)‹/›'
        
        while pos < len(text):
            match = re.search(tag_pattern, text[pos:])
            if not match:
                # 没有更多标记，添加剩余文本（使用基准格式）
                remaining = text[pos:]
                if remaining:
                    result.append((remaining, None))
                break
            
            match_start = pos + match.start()
            match_end = pos + match.end()
            
            # 添加标记前的普通文本（使用基准格式）
            if match_start > pos:
                result.append((text[pos:match_start], None))
            
            # 提取类别ID和文本内容
            cat_id = match.group(1)
            content = match.group(2)
            
            if content:
                result.append((content, cat_id))
            
            pos = match_end
        
        return result

    def _make_text_run(self, text: str, rpr: Optional[ET.Element]) -> ET.Element:
        """构建 <w:r> 元素，可选带 <w:rPr>。"""
        text = self._XML_ILLEGAL.sub('', text)
        r = ET.Element(f'{_W}r')
        if rpr is not None:
            r.append(copy.deepcopy(rpr))
        t = ET.SubElement(r, f'{_W}t')
        # 保留首尾空格
        if text.startswith(' ') or text.endswith(' '):
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        return r

    def _pack_docx(self, output_path: str, override_files: Dict[str, bytes] = None):
        """
        将 unzipped_dir 重新打包成 .docx（ZIP_DEFLATED）。
        override_files: {arcname: bytes} 用于替换指定文件内容，不修改源文件。
        """
        override_files = override_files or {}
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_d, _dirs, files in os.walk(str(self.unzipped_dir)):
                for fn in files:
                    fp = os.path.join(root_d, fn)
                    arcname = os.path.relpath(fp, str(self.unzipped_dir))
                    # 在 Windows 上，将反斜杠转换为正斜杠（ZIP 文件使用正斜杠）
                    arcname = arcname.replace('\\', '/')
                    if arcname in override_files:
                        zf.writestr(arcname, override_files[arcname])
                    else:
                        zf.write(fp, arcname)


# ====================================================================
# CLI
# ====================================================================

def main():
    import sys

    if len(sys.argv) < 4:
        print("用法: python docx_chunks_restorer.py <unzipped_dir> <chunks_dir> <output_docx>")
        print("示例: python docx_chunks_restorer.py output/unzipped output/chunks restored.docx")
        sys.exit(1)

    unzipped_dir = sys.argv[1]
    chunks_dir = sys.argv[2]
    output_docx = sys.argv[3]

    restorer = DocxChunksRestorer(unzipped_dir, chunks_dir)
    restorer.restore(output_docx)


if __name__ == '__main__':
    main()
