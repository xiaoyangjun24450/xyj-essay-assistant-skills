#!/usr/bin/env python3
"""
Markdown → DOCX Converter for docx-essay-writer skill.

Converts Markdown to a Word document that exactly matches a given template's
formatting — styles, page layout, headers/footers, fonts, etc.

All style IDs and page layout are dynamically extracted from the template.
Supports: headings, body text, bold/italic, inline & display math (LaTeX→OMML),
tables, images, ordered/unordered lists, and reference paragraphs.

Self-contained — does not import any other skill modules.
"""

import argparse
import hashlib
import os
import re
import shutil
import struct
import sys
import zipfile
import xml.etree.ElementTree as ET
from collections import Counter
from typing import Any, Dict, List, Optional, Tuple

# ---------------------------------------------------------------------------
# XML namespace constants
# ---------------------------------------------------------------------------
NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
NS_W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'
NS_WP = '{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}'
NS_A = '{http://schemas.openxmlformats.org/drawingml/2006/main}'
NS_PIC = '{http://schemas.openxmlformats.org/drawingml/2006/picture}'
NS_R_DOC = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'

NS_MAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
}

# EMU per inch / cm helpers
EMU_PER_INCH = 914400
EMU_PER_CM = 360000


# ===================================================================
# LaTeX → OMML Converter
# ===================================================================

class LatexToOmmlConverter:
    """Convert LaTeX math expressions to OMML XML elements."""

    def __init__(self, font_hint: str = 'eastAsia'):
        self._font_hint = font_hint
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
            'cdot': '⋅', 'times': '×', 'pm': '±', 'mp': '∓',
            'leq': '≤', 'geq': '≥', 'neq': '≠', 'approx': '≈',
            'infty': '∞', 'partial': '∂', 'nabla': '∇',
            'sum': '∑', 'prod': '∏', 'int': '∫',
            'rightarrow': '→', 'leftarrow': '←', 'Rightarrow': '⇒',
            'ldots': '…', 'dots': '…',
        }
        self._accent_map = {
            'hat': '\u0302', 'tilde': '\u0303', 'bar': '\u0304',
            'vec': '\u20d7', 'dot': '\u0307', 'ddot': '\u0308',
        }

    def _ctrl_pr(self):
        ctrlPr = ET.Element(f'{NS_M}ctrlPr')
        rPr = ET.SubElement(ctrlPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        if self._font_hint:
            rFonts.set(f'{NS_W}hint', self._font_hint)
        rFonts.set(f'{NS_W}ascii', 'Cambria Math')
        rFonts.set(f'{NS_W}hAnsi', 'Cambria Math')
        return ctrlPr

    def _math_run(self, text, hint=None):
        if hint is None:
            hint = 'default'
        r = ET.Element(f'{NS_M}r')
        rPr = ET.SubElement(r, f'{NS_M}rPr')
        sty = ET.SubElement(rPr, f'{NS_M}sty')
        sty.set(f'{NS_M}val', 'p')
        wrPr = ET.SubElement(r, f'{NS_W}rPr')
        rFonts = ET.SubElement(wrPr, f'{NS_W}rFonts')
        if hint:
            rFonts.set(f'{NS_W}hint', hint)
        rFonts.set(f'{NS_W}ascii', 'Cambria Math')
        rFonts.set(f'{NS_W}hAnsi', 'Cambria Math')
        t = ET.SubElement(r, f'{NS_M}t')
        t.text = text
        return r

    def _math_run_ea(self, text):
        return self._math_run(text, hint=self._font_hint or 'eastAsia')

    # ------------------------------------------------------------------

    def convert(self, latex: str) -> Optional[ET.Element]:
        latex = latex.strip()
        if not latex:
            return None
        if latex.startswith('$$') and latex.endswith('$$'):
            latex = latex[2:-2].strip()
        elif latex.startswith('$') and latex.endswith('$'):
            latex = latex[1:-1].strip()
        omath = ET.Element(f'{NS_M}oMath')
        self._parse_expr(omath, latex)
        return omath

    def _parse_expr(self, parent, expr):
        expr = expr.strip()
        if not expr:
            return

        # 1. \func(...)
        m = re.match(r'\\(cos|sin|tan|log|ln|exp|lim|max|min|sup|inf)\s*\(([^)]+)\)', expr)
        if m:
            parent.append(self._math_run(m.group(1)))
            d = ET.SubElement(parent, f'{NS_M}d')
            dPr = ET.SubElement(d, f'{NS_M}dPr')
            ET.SubElement(dPr, f'{NS_M}begChr').set(f'{NS_M}val', '(')
            ET.SubElement(dPr, f'{NS_M}endChr').set(f'{NS_M}val', ')')
            dPr.append(self._ctrl_pr())
            e = ET.SubElement(d, f'{NS_M}e')
            self._parse_expr(e, m.group(2))
            e.append(self._ctrl_pr())
            d.append(self._ctrl_pr())
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 2. Parentheses with complex inner content
        m = re.match(r'\(([^)]+)\)', expr)
        if m:
            inner = m.group(1)
            if '\\' in inner:
                d = ET.SubElement(parent, f'{NS_M}d')
                dPr = ET.SubElement(d, f'{NS_M}dPr')
                ET.SubElement(dPr, f'{NS_M}begChr').set(f'{NS_M}val', '(')
                ET.SubElement(dPr, f'{NS_M}endChr').set(f'{NS_M}val', ')')
                dPr.append(self._ctrl_pr())
                e = ET.SubElement(d, f'{NS_M}e')
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
            d = ET.SubElement(parent, f'{NS_M}d')
            dPr = ET.SubElement(d, f'{NS_M}dPr')
            ET.SubElement(dPr, f'{NS_M}begChr').set(f'{NS_M}val', '{')
            ET.SubElement(dPr, f'{NS_M}endChr').set(f'{NS_M}val', '')
            dPr.append(self._ctrl_pr())
            e = ET.SubElement(d, f'{NS_M}e')
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

        # 5a. Accents: \hat{x}, \tilde{x}, \bar{x}, etc.
        m = re.match(r'\\(hat|tilde|bar|vec|dot|ddot)\{([^}]+)\}', expr)
        if not m:
            m = re.match(r'\\(hat|tilde|bar|vec|dot|ddot)\s+([a-zA-Z])', expr)
        if m:
            rest = expr[m.end():]
            m_sub = re.match(r'_\{([^}]+)\}', rest) if rest else None
            if not m_sub and rest:
                m_sub = re.match(r'_([a-zA-Z0-9])', rest)
            m_sup = re.match(r'\^\{([^}]+)\}', rest) if rest else None
            if not m_sup and rest:
                m_sup = re.match(r'\^([a-zA-Z0-9])', rest)
            if m_sub:
                sub_el = ET.SubElement(parent, f'{NS_M}sSub')
                pr = ET.SubElement(sub_el, f'{NS_M}sSubPr')
                pr.append(self._ctrl_pr())
                base = ET.SubElement(sub_el, f'{NS_M}e')
                self._build_accent(base, m.group(2), self._accent_map[m.group(1)])
                base.append(self._ctrl_pr())
                sub = ET.SubElement(sub_el, f'{NS_M}sub')
                self._parse_expr(sub, m_sub.group(1))
                sub.append(self._ctrl_pr())
                rest = rest[m_sub.end():]
            elif m_sup:
                sup_el = ET.SubElement(parent, f'{NS_M}sSup')
                pr = ET.SubElement(sup_el, f'{NS_M}sSupPr')
                pr.append(self._ctrl_pr())
                base = ET.SubElement(sup_el, f'{NS_M}e')
                self._build_accent(base, m.group(2), self._accent_map[m.group(1)])
                base.append(self._ctrl_pr())
                sup = ET.SubElement(sup_el, f'{NS_M}sup')
                self._parse_expr(sup, m_sup.group(1))
                sup.append(self._ctrl_pr())
                rest = rest[m_sup.end():]
            else:
                self._build_accent(parent, m.group(2), self._accent_map[m.group(1)])
            if rest:
                self._parse_expr(parent, rest)
            return

        # 5b. \text{...} — regular text in math
        m = re.match(r'\\text\{([^}]+)\}', expr)
        if m:
            parent.append(self._math_run(m.group(1)))
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 5c. \sqrt{x} — square root
        m = re.match(r'\\sqrt\{([^}]+)\}', expr)
        if m:
            self._build_radical(parent, m.group(1))
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 6. Fractions \frac{a}{b}
        m = re.match(r'\\frac\{([^}]+)\}\{([^}]+)\}', expr)
        if m:
            f = ET.SubElement(parent, f'{NS_M}f')
            fPr = ET.SubElement(f, f'{NS_M}fPr')
            fPr.append(self._ctrl_pr())
            num = ET.SubElement(f, f'{NS_M}num')
            self._parse_expr(num, m.group(1))
            num.append(self._ctrl_pr())
            den = ET.SubElement(f, f'{NS_M}den')
            self._parse_expr(den, m.group(2))
            den.append(self._ctrl_pr())
            parent.append(self._ctrl_pr())
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 6. Combined sub+sup: x_d^*, \omega_e^{*}
        combined_pats = [
            r'\\([a-zA-Z]+)_\{([^}]+)\}\^\{([^}]+)\}',
            r'\\([a-zA-Z]+)_([a-zA-Z0-9])\^([a-zA-Z0-9*])',
            r'([a-zA-Z0-9])_\\([a-zA-Z]+)\^\{([^}]+)\}',
            r'([a-zA-Z0-9])_\\([a-zA-Z]+)\^([a-zA-Z0-9*])',
            r'([a-zA-Z0-9])_\{([^}]+)\}\^\{([^}]+)\}',
            r'([a-zA-Z0-9])_([a-zA-Z0-9])\^([a-zA-Z0-9*])',
        ]
        for pat in combined_pats:
            m = re.match(pat, expr)
            if m:
                self._build_subsup(parent, m.group(1), m.group(2), m.group(3))
                rest = expr[m.end():]
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 7. Subscript (with optional trailing superscript → sSubSup)
        sub_pats = [
            r'\\([a-zA-Z]+)_\{([^}]+)\}',
            r'\\([a-zA-Z]+)_([a-zA-Z0-9])',
            r'([a-zA-Z0-9])_\\([a-zA-Z]+)',
            r'([a-zA-Z0-9])_\{([^}]+)\}',
            r'([a-zA-Z0-9])_([a-zA-Z0-9])',
        ]
        for pat in sub_pats:
            m = re.match(pat, expr)
            if m:
                rest = expr[m.end():]
                m2 = re.match(r'\^\{([^}]+)\}', rest) if rest else None
                if not m2 and rest:
                    m2 = re.match(r'\^([a-zA-Z0-9*])', rest)
                if m2:
                    self._build_subsup(parent, m.group(1), m.group(2), m2.group(1))
                    rest = rest[m2.end():]
                else:
                    self._build_sub(parent, m.group(1), m.group(2))
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 8. Superscript (with optional trailing subscript → sSubSup)
        sup_pats = [
            r'\\([a-zA-Z]+)\^\{([^}]+)\}',
            r'\\([a-zA-Z]+)\^([a-zA-Z0-9*])',
            r'([a-zA-Z0-9])\^\{([^}]+)\}',
            r'([a-zA-Z0-9])\^([a-zA-Z0-9*])',
        ]
        for pat in sup_pats:
            m = re.match(pat, expr)
            if m:
                rest = expr[m.end():]
                m2 = re.match(r'_\{([^}]+)\}', rest) if rest else None
                if not m2 and rest:
                    m2 = re.match(r'_([a-zA-Z0-9])', rest)
                if m2:
                    self._build_subsup(parent, m.group(1), m2.group(1), m.group(2))
                    rest = rest[m2.end():]
                else:
                    self._build_sup(parent, m.group(1), m.group(2))
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 9. Greek letters
        m = re.match(r'\\([a-zA-Z]+)', expr)
        if m and m.group(1) in self.greek_map:
            parent.append(self._math_run_ea(self.greek_map[m.group(1)]))
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        # 10. Operators & symbols
        if expr[0] in '=+-*/()[]{},;: |!<>~':
            parent.append(self._math_run(expr[0]))
            if len(expr) > 1:
                self._parse_expr(parent, expr[1:])
            return

        # 11. Regular text/numbers
        m = re.match(r'([a-zA-Z0-9.]+)', expr)
        if m:
            parent.append(self._math_run(m.group(1)))
            rest = expr[m.end():]
            if rest:
                self._parse_expr(parent, rest)
            return

        if len(expr) > 1:
            self._parse_expr(parent, expr[1:])

    # ------------------------------------------------------------------
    # Builder helpers
    # ------------------------------------------------------------------

    def _resolve(self, token):
        if token in self.greek_map:
            return self._math_run_ea(self.greek_map[token])
        return self._math_run(token)

    def _build_matrix(self, parent, content, beg, end):
        rows = [r.strip() for r in content.strip().split('\\\\') if r.strip()]
        d = ET.SubElement(parent, f'{NS_M}d')
        dPr = ET.SubElement(d, f'{NS_M}dPr')
        ET.SubElement(dPr, f'{NS_M}begChr').set(f'{NS_M}val', beg)
        ET.SubElement(dPr, f'{NS_M}endChr').set(f'{NS_M}val', end)
        dPr.append(self._ctrl_pr())
        e = ET.SubElement(d, f'{NS_M}e')
        mat = ET.SubElement(e, f'{NS_M}m')
        mPr = ET.SubElement(mat, f'{NS_M}mPr')
        mcs = ET.SubElement(mPr, f'{NS_M}mcs')
        mc = ET.SubElement(mcs, f'{NS_M}mc')
        mcPr = ET.SubElement(mc, f'{NS_M}mcPr')
        ET.SubElement(mcPr, f'{NS_M}count').set(f'{NS_M}val', '1')
        ET.SubElement(mcPr, f'{NS_M}mcJc').set(f'{NS_M}val', 'center')
        ET.SubElement(mPr, f'{NS_M}plcHide').set(f'{NS_M}val', '1')
        mPr.append(self._ctrl_pr())
        for row in rows:
            mr = ET.SubElement(mat, f'{NS_M}mr')
            cells = [c.strip() for c in row.split('&')]
            for cell in cells:
                ec = ET.SubElement(mr, f'{NS_M}e')
                self._parse_expr(ec, cell)
                ec.append(self._ctrl_pr())
        e.append(self._ctrl_pr())
        d.append(self._ctrl_pr())

    def _build_accent(self, parent, content, accent_char):
        acc = ET.SubElement(parent, f'{NS_M}acc')
        accPr = ET.SubElement(acc, f'{NS_M}accPr')
        chrEl = ET.SubElement(accPr, f'{NS_M}chr')
        chrEl.set(f'{NS_M}val', accent_char)
        accPr.append(self._ctrl_pr())
        e = ET.SubElement(acc, f'{NS_M}e')
        self._parse_expr(e, content)
        e.append(self._ctrl_pr())

    def _build_radical(self, parent, content):
        rad = ET.SubElement(parent, f'{NS_M}rad')
        radPr = ET.SubElement(rad, f'{NS_M}radPr')
        ET.SubElement(radPr, f'{NS_M}degHide').set(f'{NS_M}val', '1')
        radPr.append(self._ctrl_pr())
        deg = ET.SubElement(rad, f'{NS_M}deg')
        deg.append(self._ctrl_pr())
        e = ET.SubElement(rad, f'{NS_M}e')
        self._parse_expr(e, content)
        e.append(self._ctrl_pr())

    def _build_subsup(self, parent, base, sub_val, sup_val):
        el = ET.SubElement(parent, f'{NS_M}sSubSup')
        pr = ET.SubElement(el, f'{NS_M}sSubSupPr')
        pr.append(self._ctrl_pr())
        e = ET.SubElement(el, f'{NS_M}e')
        e.append(self._resolve(base))
        e.append(self._ctrl_pr())
        sub = ET.SubElement(el, f'{NS_M}sub')
        sub.append(self._resolve(sub_val))
        sub.append(self._ctrl_pr())
        sup = ET.SubElement(el, f'{NS_M}sup')
        sup.append(self._resolve(sup_val))
        sup.append(self._ctrl_pr())
        el.append(self._ctrl_pr())

    def _build_sub(self, parent, base, sub_val):
        el = ET.SubElement(parent, f'{NS_M}sSub')
        pr = ET.SubElement(el, f'{NS_M}sSubPr')
        pr.append(self._ctrl_pr())
        e = ET.SubElement(el, f'{NS_M}e')
        e.append(self._resolve(base))
        e.append(self._ctrl_pr())
        sub = ET.SubElement(el, f'{NS_M}sub')
        sub.append(self._resolve(sub_val))
        sub.append(self._ctrl_pr())
        el.append(self._ctrl_pr())

    def _build_sup(self, parent, base, sup_val):
        el = ET.SubElement(parent, f'{NS_M}sSup')
        pr = ET.SubElement(el, f'{NS_M}sSupPr')
        pr.append(self._ctrl_pr())
        e = ET.SubElement(el, f'{NS_M}e')
        e.append(self._resolve(base))
        e.append(self._ctrl_pr())
        sup = ET.SubElement(el, f'{NS_M}sup')
        sup.append(self._resolve(sup_val))
        sup.append(self._ctrl_pr())
        el.append(self._ctrl_pr())


# ===================================================================
# Style Extractor
# ===================================================================

class StyleExtractor:
    """Extract style IDs and sectPr from a DOCX template."""

    def __init__(self, temp_dir: str):
        self.temp_dir = temp_dir
        self._styles_root: Optional[ET.Element] = None
        self._doc_root: Optional[ET.Element] = None

    def _read(self, rel_path: str) -> Optional[ET.Element]:
        p = os.path.join(self.temp_dir, rel_path)
        if os.path.isfile(p):
            return ET.parse(p).getroot()
        return None

    @property
    def styles_root(self):
        if self._styles_root is None:
            self._styles_root = self._read('word/styles.xml')
        return self._styles_root

    @property
    def doc_root(self):
        if self._doc_root is None:
            self._doc_root = self._read('word/document.xml')
        return self._doc_root

    def heading_ids(self) -> Dict[int, str]:
        """Map heading level → styleId from styles.xml."""
        result: Dict[int, str] = {}
        root = self.styles_root
        if root is None:
            return result
        for style in root.findall(f'.//{NS_W}style'):
            sid = style.get(f'{NS_W}styleId', '')
            name_el = style.find(f'{NS_W}name')
            if name_el is None:
                continue
            name = name_el.get(f'{NS_W}val', '').lower()
            for lvl in range(1, 10):
                if name in (f'heading {lvl}', f'heading{lvl}',
                            f'标题 {lvl}', f'标题{lvl}'):
                    result[lvl] = sid
        return result

    def body_style_id(self) -> str:
        root = self.styles_root
        if root is None:
            return ''
        for style in root.findall(f'.//{NS_W}style'):
            name_el = style.find(f'{NS_W}name')
            if name_el is None:
                continue
            name = name_el.get(f'{NS_W}val', '').lower()
            if name in ('body text', 'body', '正文', '正文 首行缩进 2'):
                return style.get(f'{NS_W}styleId', '')
        return ''

    def find_style_id(self, *candidates: str) -> str:
        """Find first matching style by name (case-insensitive)."""
        root = self.styles_root
        if root is None:
            return ''
        lower_cands = [c.lower() for c in candidates]
        for style in root.findall(f'.//{NS_W}style'):
            name_el = style.find(f'{NS_W}name')
            if name_el is None:
                continue
            name = name_el.get(f'{NS_W}val', '').lower()
            if name in lower_cands:
                return style.get(f'{NS_W}styleId', '')
        return ''

    def extract_sect_pr_xml(self) -> Optional[str]:
        """Return raw XML string of the template's last w:sectPr."""
        doc_path = os.path.join(self.temp_dir, 'word', 'document.xml')
        if not os.path.isfile(doc_path):
            return None
        with open(doc_path, 'r', encoding='utf-8') as f:
            content = f.read()
        m = re.search(r'(<w:sectPr[\s\S]*?</w:sectPr>)', content)
        if m:
            return m.group(1)
        m = re.search(r'(<w:sectPr[^/]*/\s*>)', content)
        if m:
            return m.group(1)
        return None


# ===================================================================
# Template Format Extractor
# ===================================================================

def extract_template_formats(temp_dir: str) -> Dict[str, Any]:
    """Dynamically extract ALL formatting info from the template.

    Returns a dict with keys like font_hint, body_style_id, body_lang,
    heading_numPr, formula_rfonts, reference_style_id, image_style_id, etc.
    This eliminates the need for any hardcoded format values.
    """
    styles_path = os.path.join(temp_dir, 'word', 'styles.xml')
    doc_path = os.path.join(temp_dir, 'word', 'document.xml')

    result: Dict[str, Any] = {}

    styles_root = None
    doc_root = None
    if os.path.isfile(styles_path):
        styles_root = ET.parse(styles_path).getroot()
    if os.path.isfile(doc_path):
        doc_root = ET.parse(doc_path).getroot()

    if styles_root is None:
        return result

    # Build lookups: styleId -> element, lower(name) -> element
    style_by_id: Dict[str, ET.Element] = {}
    style_by_name: Dict[str, ET.Element] = {}
    for s in styles_root.findall(f'.//{NS_W}style'):
        sid = s.get(f'{NS_W}styleId', '')
        if sid:
            style_by_id[sid] = s
        name_el = s.find(f'{NS_W}name')
        if name_el is not None:
            nm = name_el.get(f'{NS_W}val', '')
            if nm:
                style_by_name[nm.lower()] = s

    # ---- font_hint from docDefaults ----
    has_east_asian = False
    defaults = styles_root.find(f'{NS_W}docDefaults')
    if defaults is not None:
        rPrDef = defaults.find(f'.//{NS_W}rPrDefault')
        if rPrDef is not None:
            rPr = rPrDef.find(f'{NS_W}rPr')
            if rPr is not None:
                rFonts_el = rPr.find(f'{NS_W}rFonts')
                if rFonts_el is not None:
                    hint_val = rFonts_el.get(f'{NS_W}hint')
                    if hint_val:
                        result['font_hint'] = hint_val
                    if rFonts_el.get(f'{NS_W}eastAsia') or rFonts_el.get(f'{NS_W}eastAsiaTheme'):
                        has_east_asian = True
                lang_el = rPr.find(f'{NS_W}lang')
                if lang_el is not None:
                    if lang_el.get(f'{NS_W}eastAsia'):
                        has_east_asian = True
                    result['default_lang'] = {
                        k.split('}')[-1]: v for k, v in lang_el.attrib.items()
                    }

    if 'font_hint' not in result and has_east_asian:
        result['font_hint'] = 'eastAsia'

    # ---- body_style_id: most-frequent paragraph style in document ----
    heading_sids: set = set()
    for s in styles_root.findall(f'.//{NS_W}style'):
        name_el = s.find(f'{NS_W}name')
        if name_el is not None:
            nm = name_el.get(f'{NS_W}val', '').lower()
            if 'heading' in nm or '标题' in nm:
                heading_sids.add(s.get(f'{NS_W}styleId', ''))

    if doc_root is not None:
        body_el = doc_root.find(f'{NS_W}body')
        if body_el is not None:
            style_counter: Counter = Counter()
            for p in body_el.findall(f'{NS_W}p'):
                pPr = p.find(f'{NS_W}pPr')
                if pPr is not None:
                    ps = pPr.find(f'{NS_W}pStyle')
                    if ps is not None:
                        sid = ps.get(f'{NS_W}val', '')
                        if sid and sid not in heading_sids:
                            style_counter[sid] += 1
            if style_counter:
                result['body_style_id'] = style_counter.most_common(1)[0][0]

    # ---- body_rfonts, body_lang from body style ----
    body_sid = result.get('body_style_id', '')
    if body_sid and body_sid in style_by_id:
        body_s = style_by_id[body_sid]
        rPr = body_s.find(f'{NS_W}rPr')
        if rPr is not None:
            rFonts_el = rPr.find(f'{NS_W}rFonts')
            if rFonts_el is not None:
                result['body_rfonts'] = {
                    k.split('}')[-1]: v for k, v in rFonts_el.attrib.items()
                    if 'Theme' not in k
                }
            lang_el = rPr.find(f'{NS_W}lang')
            if lang_el is not None:
                result['body_lang'] = {
                    k.split('}')[-1]: v for k, v in lang_el.attrib.items()
                }

    # ---- heading_numPr: per heading level from styles.xml ----
    heading_numpr: Dict[str, Optional[Dict[str, str]]] = {}
    for lvl in range(1, 10):
        candidates = [f'heading {lvl}', f'heading{lvl}', f'标题 {lvl}', f'标题{lvl}']
        for cand in candidates:
            if cand in style_by_name:
                s = style_by_name[cand]
                pPr = s.find(f'{NS_W}pPr')
                key = f'heading{lvl}'
                if pPr is not None:
                    numPr = pPr.find(f'{NS_W}numPr')
                    if numPr is not None:
                        ilvl_el = numPr.find(f'{NS_W}ilvl')
                        numId_el = numPr.find(f'{NS_W}numId')
                        heading_numpr[key] = {
                            'ilvl': ilvl_el.get(f'{NS_W}val', str(lvl - 1)) if ilvl_el is not None else str(lvl - 1),
                            'numId': numId_el.get(f'{NS_W}val', '0') if numId_el is not None else '0',
                        }
                    else:
                        heading_numpr[key] = None
                else:
                    heading_numpr[key] = None
                break
    result['heading_numPr'] = heading_numpr

    # ---- formula format: follow link → basedOn chain ----
    def _follow_chain_rpr(start_sid: str, attr_name: str, visited: Optional[set] = None) -> Dict[str, str]:
        if visited is None:
            visited = set()
        if start_sid in visited or start_sid not in style_by_id:
            return {}
        visited.add(start_sid)
        s = style_by_id[start_sid]
        rPr = s.find(f'{NS_W}rPr')
        if rPr is not None:
            el = rPr.find(f'{NS_W}{attr_name}')
            if el is not None and len(el.attrib) > 0:
                return {k.split('}')[-1]: v for k, v in el.attrib.items()}
        link = s.find(f'{NS_W}link')
        if link is not None:
            attrs = _follow_chain_rpr(link.get(f'{NS_W}val', ''), attr_name, visited)
            if attrs:
                return attrs
        basedOn = s.find(f'{NS_W}basedOn')
        if basedOn is not None:
            return _follow_chain_rpr(basedOn.get(f'{NS_W}val', ''), attr_name, visited)
        return {}

    formula_sid = ''
    for name_key in ('公式', 'formula', 'equation', 'mtdisplayequation'):
        if name_key in style_by_name:
            formula_sid = style_by_name[name_key].get(f'{NS_W}styleId', '')
            break
    if formula_sid:
        f_rfonts = _follow_chain_rpr(formula_sid, 'rFonts')
        if f_rfonts:
            result['formula_rfonts'] = {k: v for k, v in f_rfonts.items() if 'Theme' not in k}
        f_lang = _follow_chain_rpr(formula_sid, 'lang')
        if f_lang:
            result['formula_lang'] = f_lang

    # ---- img_max_width_emu from sectPr ----
    if doc_root is not None:
        for sect_pr in doc_root.iter(f'{NS_W}sectPr'):
            pgSz = sect_pr.find(f'{NS_W}pgSz')
            pgMar = sect_pr.find(f'{NS_W}pgMar')
            if pgSz is not None and pgMar is not None:
                try:
                    page_w = int(pgSz.get(f'{NS_W}w', '11906'))
                    margin_l = int(pgMar.get(f'{NS_W}left', '1800'))
                    margin_r = int(pgMar.get(f'{NS_W}right', '1800'))
                    content_twips = page_w - margin_l - margin_r
                    result['img_max_width_emu'] = int(content_twips * EMU_PER_INCH / 1440)
                except ValueError:
                    pass

    # ---- special style IDs ----
    for name_key in ('参考文献', 'reference', 'bibliography'):
        if name_key in style_by_name:
            result['reference_style_id'] = style_by_name[name_key].get(f'{NS_W}styleId', '')
            break

    for name_key in ('图片', 'image', 'figure'):
        if name_key in style_by_name:
            s = style_by_name[name_key]
            if s.get(f'{NS_W}type') == 'paragraph':
                result['image_style_id'] = s.get(f'{NS_W}styleId', '')
                break

    for name_key in ('三线表', 'table grid', 'normal table'):
        if name_key in style_by_name:
            s = style_by_name[name_key]
            if s.get(f'{NS_W}type') == 'table':
                result['table_style_id'] = s.get(f'{NS_W}styleId', '')
                break

    # ---- table format: extract from ACTUAL tables in document ----
    def _extract_cell_borders(tc_el: ET.Element) -> Optional[Dict[str, Dict[str, str]]]:
        tcPr = tc_el.find(f'{NS_W}tcPr')
        if tcPr is None:
            return None
        borders = tcPr.find(f'{NS_W}tcBorders')
        if borders is None:
            return None
        border_dict: Dict[str, Dict[str, str]] = {}
        for b in borders:
            bname = b.tag.split('}')[-1]
            border_dict[bname] = {k.split('}')[-1]: v for k, v in b.attrib.items()}
        return border_dict if border_dict else None

    if doc_root is not None:
        body_el = doc_root.find(f'{NS_W}body')
        if body_el is not None:
            for tbl in body_el.findall(f'{NS_W}tbl'):
                tbl_fmt: Dict[str, Any] = {}
                tblPr = tbl.find(f'{NS_W}tblPr')
                if tblPr is not None:
                    ts = tblPr.find(f'{NS_W}tblStyle')
                    if ts is not None:
                        tbl_fmt['style'] = ts.get(f'{NS_W}val', '')
                    tw = tblPr.find(f'{NS_W}tblW')
                    if tw is not None:
                        tbl_fmt['width_w'] = tw.get(f'{NS_W}w', '5000')
                        tbl_fmt['width_type'] = tw.get(f'{NS_W}type', 'pct')
                    jc_el = tblPr.find(f'{NS_W}jc')
                    if jc_el is not None:
                        tbl_fmt['jc'] = jc_el.get(f'{NS_W}val', 'center')
                    layout_el = tblPr.find(f'{NS_W}tblLayout')
                    if layout_el is not None:
                        tbl_fmt['layout'] = layout_el.get(f'{NS_W}type', 'autofit')
                    borders_el = tblPr.find(f'{NS_W}tblBorders')
                    if borders_el is not None:
                        tbl_fmt['borders'] = {}
                        for b in borders_el:
                            bname = b.tag.split('}')[-1]
                            tbl_fmt['borders'][bname] = {
                                k.split('}')[-1]: v for k, v in b.attrib.items()
                            }
                    look_el = tblPr.find(f'{NS_W}tblLook')
                    if look_el is not None:
                        tbl_fmt['look'] = {
                            k.split('}')[-1]: v for k, v in look_el.attrib.items()
                        }
                rows = tbl.findall(f'{NS_W}tr')
                if rows:
                    trPr = rows[0].find(f'{NS_W}trPr')
                    if trPr is not None:
                        tr_fmt: Dict[str, Dict] = {}
                        for child in trPr:
                            cname = child.tag.split('}')[-1]
                            tr_fmt[cname] = {
                                k.split('}')[-1]: v for k, v in child.attrib.items()
                            }
                        tbl_fmt['row_format'] = tr_fmt

                    # Cell-level border patterns by row position
                    cell_borders_map: Dict[str, Dict] = {}
                    first_tc = rows[0].find(f'{NS_W}tc')
                    if first_tc is not None:
                        hdr_borders = _extract_cell_borders(first_tc)
                        if hdr_borders:
                            cell_borders_map['header'] = hdr_borders
                    if len(rows) > 2:
                        mid_idx = len(rows) // 2
                        mid_tc = rows[mid_idx].find(f'{NS_W}tc')
                        if mid_tc is not None:
                            body_borders = _extract_cell_borders(mid_tc)
                            if body_borders:
                                cell_borders_map['body'] = body_borders
                    if len(rows) > 1:
                        last_tc = rows[-1].find(f'{NS_W}tc')
                        if last_tc is not None:
                            last_borders = _extract_cell_borders(last_tc)
                            if last_borders:
                                cell_borders_map['last'] = last_borders
                    if cell_borders_map:
                        tbl_fmt['cell_borders'] = cell_borders_map

                    # Cell formatting from first header cell
                    tcs = rows[0].findall(f'{NS_W}tc')
                    if tcs:
                        tc0 = tcs[0]
                        tcPr0 = tc0.find(f'{NS_W}tcPr')
                        if tcPr0 is not None:
                            shd = tcPr0.find(f'{NS_W}shd')
                            if shd is not None:
                                tbl_fmt['cell_shd'] = {
                                    k.split('}')[-1]: v for k, v in shd.attrib.items()
                                    if 'Theme' not in k
                                }
                        p = tc0.find(f'{NS_W}p')
                        if p is not None:
                            pPr = p.find(f'{NS_W}pPr')
                            if pPr is not None:
                                ps = pPr.find(f'{NS_W}pStyle')
                                if ps is not None:
                                    tbl_fmt['cell_style'] = ps.get(f'{NS_W}val', '')
                                ind_el = pPr.find(f'{NS_W}ind')
                                if ind_el is not None:
                                    tbl_fmt['cell_ind'] = {
                                        k.split('}')[-1]: v for k, v in ind_el.attrib.items()
                                    }
                                rPr = pPr.find(f'{NS_W}rPr')
                                if rPr is not None:
                                    cell_rpr: Dict[str, Any] = {}
                                    for child in rPr:
                                        cname = child.tag.split('}')[-1]
                                        if cname == 'rFonts':
                                            continue
                                        if child.attrib:
                                            cell_rpr[cname] = {
                                                k.split('}')[-1]: v for k, v in child.attrib.items()
                                            }
                                        else:
                                            cell_rpr[cname] = True
                                    if cell_rpr:
                                        tbl_fmt['cell_rPr'] = cell_rpr
                result['table_format'] = tbl_fmt
                if 'style' in tbl_fmt:
                    result['table_style_id'] = tbl_fmt['style']
                break

            # ---- table caption format: centered paragraph before first table ----
            prev_p = None
            for child in body_el:
                if child.tag == f'{NS_W}tbl' and prev_p is not None:
                    pPr = prev_p.find(f'{NS_W}pPr')
                    if pPr is not None:
                        jc_el = pPr.find(f'{NS_W}jc')
                        if jc_el is not None and jc_el.get(f'{NS_W}val', '') == 'center':
                            cap_fmt: Dict[str, Any] = {}
                            ps = pPr.find(f'{NS_W}pStyle')
                            if ps is not None:
                                cap_fmt['style'] = ps.get(f'{NS_W}val', '')
                            ind_el = pPr.find(f'{NS_W}ind')
                            if ind_el is not None:
                                cap_fmt['ind'] = {
                                    k.split('}')[-1]: v for k, v in ind_el.attrib.items()
                                }
                            cap_fmt['jc'] = 'center'
                            rPr = pPr.find(f'{NS_W}rPr')
                            if rPr is not None:
                                if rPr.find(f'{NS_W}b') is not None:
                                    cap_fmt['bold'] = True
                                sz_el = rPr.find(f'{NS_W}sz')
                                if sz_el is not None:
                                    cap_fmt['sz'] = sz_el.get(f'{NS_W}val', '')
                                szCs_el = rPr.find(f'{NS_W}szCs')
                                if szCs_el is not None:
                                    cap_fmt['szCs'] = szCs_el.get(f'{NS_W}val', '')
                            result['table_caption_format'] = cap_fmt
                            break
                prev_p = child if child.tag == f'{NS_W}p' else None

    return result


# ===================================================================
# Image helper
# ===================================================================

def _image_size_emu(img_path: str, max_width_emu: int = 5400000) -> Tuple[int, int]:
    """Get image dimensions in EMU, fitting within max_width_emu."""
    try:
        from PIL import Image
        with Image.open(img_path) as im:
            w_px, h_px = im.size
            dpi = im.info.get('dpi', (96, 96))
            dpi_x = dpi[0] if dpi[0] else 96
            dpi_y = dpi[1] if dpi[1] else 96
    except Exception:
        w_px, h_px = 400, 300
        dpi_x = dpi_y = 96

    w_emu = int(w_px / dpi_x * EMU_PER_INCH)
    h_emu = int(h_px / dpi_y * EMU_PER_INCH)

    if w_emu > max_width_emu:
        ratio = max_width_emu / w_emu
        w_emu = max_width_emu
        h_emu = int(h_emu * ratio)
    return w_emu, h_emu


# ===================================================================
# Markdown → DOCX Converter
# ===================================================================

class MarkdownToDocxConverter:
    """Convert Markdown text to DOCX using a template for formatting."""

    def __init__(self, template_path: str):
        self.template_path = template_path
        self.latex_converter = LatexToOmmlConverter()
        self._para_id = 0x10000000
        self._rel_id_counter = 100
        self._image_rels: List[Dict] = []

    def _next_para_id(self) -> str:
        pid = hex(self._para_id)[2:].upper().zfill(8)
        self._para_id += 1
        return pid

    def _next_rel_id(self) -> str:
        self._rel_id_counter += 1
        return f'rId{self._rel_id_counter}'

    # ------------------------------------------------------------------
    # Rels update for images
    # ------------------------------------------------------------------

    def _update_rels(self, temp_dir: str):
        rels_path = os.path.join(temp_dir, 'word', '_rels', 'document.xml.rels')
        if not os.path.isfile(rels_path):
            return
        tree = ET.parse(rels_path)
        root = tree.getroot()
        ns = 'http://schemas.openxmlformats.org/package/2006/relationships'
        for img_rel in self._image_rels:
            rel = ET.SubElement(root, f'{{{ns}}}Relationship')
            rel.set('Id', img_rel['rId'])
            rel.set('Type',
                     'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image')
            rel.set('Target', img_rel['target'])
        tree.write(rels_path, xml_declaration=True, encoding='UTF-8')

    # ------------------------------------------------------------------
    # Document XML generation
    # ------------------------------------------------------------------

    def _generate_document_xml(self, markdown: str) -> str:
        lines = markdown.split('\n')
        root = ET.Element(f'{NS_W}document')
        root.set('xmlns:wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas')
        root.set('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        root.set('xmlns:o', 'urn:schemas-microsoft-com:office:office')
        root.set('xmlns:v', 'urn:schemas-microsoft-com:vml')
        root.set('xmlns:wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')
        root.set('xmlns:wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        root.set('xmlns:w10', 'urn:schemas-microsoft-com:office:word')
        root.set('xmlns:w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
        root.set('xmlns:wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')
        root.set('xmlns:wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk')
        root.set('xmlns:wne', 'http://schemas.microsoft.com/office/word/2006/wordml')
        root.set('xmlns:wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
        root.set('xmlns:wpsCustomData', 'http://www.wps.cn/officeDocument/2013/wpsCustomData')
        root.set('xmlns:r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        root.set('xmlns:a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
        root.set('xmlns:pic', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
        root.set('mc:Ignorable', 'w14 w15 wp14')

        body = ET.SubElement(root, f'{NS_W}body')
        body.append(self._empty_para())

        i = 0
        while i < len(lines):
            line = lines[i].rstrip()
            stripped = line.strip()

            if not stripped:
                i += 1
                continue

            # Headings
            if stripped.startswith('### '):
                body.append(self._heading_para(3, stripped[4:].strip()))
                i += 1
            elif stripped.startswith('## '):
                body.append(self._heading_para(2, stripped[3:].strip()))
                i += 1
            elif stripped.startswith('# '):
                body.append(self._heading_para(1, stripped[2:].strip()))
                i += 1

            # Table
            elif stripped.startswith('|'):
                tbl_lines: List[str] = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    tbl_lines.append(lines[i])
                    i += 1
                tbl = self._create_table(tbl_lines)
                if tbl is not None:
                    body.append(tbl)

            # Code block
            elif stripped.startswith('```'):
                code_lines: List[str] = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                if i < len(lines):
                    i += 1
                body.append(self._code_para('\n'.join(code_lines)))

            # Display math $$...$$
            elif stripped.startswith('$$'):
                formula_parts = [stripped]
                if not stripped.endswith('$$') or stripped == '$$':
                    i += 1
                    while i < len(lines) and not lines[i].strip().endswith('$$'):
                        formula_parts.append(lines[i])
                        i += 1
                    if i < len(lines):
                        formula_parts.append(lines[i].strip())
                        i += 1
                else:
                    i += 1
                body.append(self._formula_para('\n'.join(formula_parts)))

            # Image ![caption](path)
            elif re.match(r'^!\[.*?\]\(.*?\)\s*$', stripped):
                body.append(self._image_para(stripped))
                i += 1

            # Unordered list - item
            elif stripped.startswith('- '):
                body.append(self._list_para(stripped[2:], ordered=False))
                i += 1

            # Ordered list 1. item
            elif re.match(r'^\d+\.\s', stripped):
                text = re.sub(r'^\d+\.\s+', '', stripped)
                body.append(self._list_para(text, ordered=True))
                i += 1

            # Reference paragraph [N] Author...
            elif re.match(r'^\[\d+\]\s', stripped):
                body.append(self._reference_para(stripped))
                i += 1

            # Table / figure caption: 表X-Y ... or 图X-Y ...
            elif re.match(r'^(\*\*)?[表图]\d+-\d+', stripped):
                cap_text = stripped.strip('*').strip()
                body.append(self._table_caption_para(cap_text))
                i += 1

            # Normal body paragraph (may span multiple lines)
            else:
                para_parts = [line]
                i += 1
                while i < len(lines):
                    nxt = lines[i].rstrip()
                    if not nxt.strip():
                        break
                    if nxt.strip().startswith(('#', '|', '$$', '```', '![', '- ')):
                        break
                    if re.match(r'^\d+\.\s', nxt.strip()):
                        break
                    if re.match(r'^\[\d+\]\s', nxt.strip()):
                        break
                    if re.match(r'^(\*\*)?[表图]\d+-\d+', nxt.strip()):
                        break
                    para_parts.append(nxt)
                    i += 1
                body.append(self._body_para('\n'.join(para_parts)))

        # Append sectPr
        if self._sect_pr_xml:
            body_xml_placeholder = '__SECT_PR_PLACEHOLDER__'
            placeholder_el = ET.SubElement(body, f'{NS_W}p')
            placeholder_el.text = body_xml_placeholder
        else:
            body.append(self._empty_para())

        xml_str = ET.tostring(root, encoding='unicode')

        # Fix namespace prefixes
        replacements = [
            ('xmlns:ns0=', 'xmlns:w='), ('<ns0:', '<w:'), ('</ns0:', '</w:'), (' ns0:', ' w:'),
            ('xmlns:ns1=', 'xmlns:w14='), ('<ns1:', '<w14:'), ('</ns1:', '</w14:'), (' ns1:', ' w14:'),
            ('xmlns:ns2=', 'xmlns:m='), ('<ns2:', '<m:'), ('</ns2:', '</m:'), (' ns2:', ' m:'),
            ('xmlns:ns3=', 'xmlns:r='), ('<ns3:', '<r:'), ('</ns3:', '</r:'), (' ns3:', ' r:'),
            ('xmlns:ns4=', 'xmlns:wp='), ('<ns4:', '<wp:'), ('</ns4:', '</wp:'), (' ns4:', ' wp:'),
            ('xmlns:ns5=', 'xmlns:a='), ('<ns5:', '<a:'), ('</ns5:', '</a:'), (' ns5:', ' a:'),
            ('xmlns:ns6=', 'xmlns:pic='), ('<ns6:', '<pic:'), ('</ns6:', '</pic:'), (' ns6:', ' pic:'),
            (' />', '/>'),
        ]
        for old, new in replacements:
            xml_str = xml_str.replace(old, new)

        if self._sect_pr_xml:
            sect_tag = f'<w:p>{body_xml_placeholder}</w:p>'
            xml_str = xml_str.replace(sect_tag, self._sect_pr_xml)

        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str

    # ------------------------------------------------------------------
    # Paragraph builders
    # ------------------------------------------------------------------

    def _empty_para(self) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        hint = self.tmpl_fmt.get('font_hint', '')
        if hint:
            rPr = ET.SubElement(pPr, f'{NS_W}rPr')
            rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
            rFonts.set(f'{NS_W}hint', hint)
        return p

    def _heading_para(self, level: int, text: str) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        sid = self._heading_ids.get(level, '')
        if sid:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', sid)
        numpr_info = self.tmpl_fmt.get('heading_numPr', {}).get(f'heading{level}')
        if numpr_info is not None:
            numPr = ET.SubElement(pPr, f'{NS_W}numPr')
            ilvl_el = ET.SubElement(numPr, f'{NS_W}ilvl')
            ilvl_el.set(f'{NS_W}val', numpr_info['ilvl'])
            numId_el = ET.SubElement(numPr, f'{NS_W}numId')
            numId_el.set(f'{NS_W}val', '0')
        hint = self.tmpl_fmt.get('font_hint', '')
        if hint:
            rPr = ET.SubElement(pPr, f'{NS_W}rPr')
            rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
            rFonts.set(f'{NS_W}hint', hint)
        r = ET.SubElement(p, f'{NS_W}r')
        if hint:
            rPr2 = ET.SubElement(r, f'{NS_W}rPr')
            rFonts2 = ET.SubElement(rPr2, f'{NS_W}rFonts')
            rFonts2.set(f'{NS_W}hint', hint)
        t = ET.SubElement(r, f'{NS_W}t')
        t.text = text
        return p

    def _body_para(self, text: str) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        if self._body_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', self._body_id)
        hint = self.tmpl_fmt.get('font_hint', '')
        body_lang = self.tmpl_fmt.get('body_lang') or self.tmpl_fmt.get('default_lang', {})
        if hint or body_lang:
            rPr = ET.SubElement(pPr, f'{NS_W}rPr')
            if hint:
                rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
                rFonts.set(f'{NS_W}hint', hint)
            if body_lang:
                lang_el = ET.SubElement(rPr, f'{NS_W}lang')
                for attr, val in body_lang.items():
                    lang_el.set(f'{NS_W}{attr}', val)

        parts = re.split(r'(\$[^$]+\$)', text)
        for part in parts:
            if part.startswith('$') and part.endswith('$'):
                omml = self.latex_converter.convert(part[1:-1])
                if omml is not None:
                    p.append(omml)
            else:
                self._add_rich_runs(p, part)
        return p

    def _add_rich_runs(self, p: ET.Element, text: str):
        """Handle **bold** and *italic* in text, add w:r elements to p."""
        hint = self.tmpl_fmt.get('font_hint', '')
        segments = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
        for seg in segments:
            if not seg or not seg.strip():
                if seg and (' ' in seg or '\t' in seg):
                    r = ET.SubElement(p, f'{NS_W}r')
                    if hint:
                        rPr = ET.SubElement(r, f'{NS_W}rPr')
                        rF = ET.SubElement(rPr, f'{NS_W}rFonts')
                        rF.set(f'{NS_W}hint', hint)
                    t = ET.SubElement(r, f'{NS_W}t')
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    t.text = seg
                continue
            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            if hint:
                rF = ET.SubElement(rPr, f'{NS_W}rFonts')
                rF.set(f'{NS_W}hint', hint)
            if seg.startswith('**') and seg.endswith('**'):
                ET.SubElement(rPr, f'{NS_W}b')
                seg = seg[2:-2]
            elif seg.startswith('*') and seg.endswith('*'):
                ET.SubElement(rPr, f'{NS_W}i')
                seg = seg[1:-1]
            t = ET.SubElement(r, f'{NS_W}t')
            if seg.startswith(' ') or seg.endswith(' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = seg

    def _formula_para(self, formula: str) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        if self._formula_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', self._formula_id)
        bidi = ET.SubElement(pPr, f'{NS_W}bidi')
        bidi.set(f'{NS_W}val', '0')

        f_rfonts = self.tmpl_fmt.get('formula_rfonts', {})
        f_lang = self.tmpl_fmt.get('formula_lang') or self.tmpl_fmt.get('default_lang', {})
        hint = self.tmpl_fmt.get('font_hint', '')

        if f_rfonts or f_lang:
            rPr = ET.SubElement(pPr, f'{NS_W}rPr')
            if f_rfonts:
                rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
                rFonts.set(f'{NS_W}hint', f_rfonts.get('hint', 'default'))
                for attr in ('eastAsia', 'ascii', 'hAnsi'):
                    if attr in f_rfonts:
                        rFonts.set(f'{NS_W}{attr}', f_rfonts[attr])
            if f_lang:
                lang_el = ET.SubElement(rPr, f'{NS_W}lang')
                for attr, val in f_lang.items():
                    lang_el.set(f'{NS_W}{attr}', val)

        # Tab before formula
        r_tab = ET.SubElement(p, f'{NS_W}r')
        rPr_tab = ET.SubElement(r_tab, f'{NS_W}rPr')
        rF_tab = ET.SubElement(rPr_tab, f'{NS_W}rFonts')
        if hint:
            rF_tab.set(f'{NS_W}hint', hint)
        rF_tab.set(f'{NS_W}hAnsi', 'Cambria Math')
        b_el = ET.SubElement(rPr_tab, f'{NS_W}b')
        b_el.set(f'{NS_W}val', '0')
        i_el = ET.SubElement(rPr_tab, f'{NS_W}i')
        i_el.set(f'{NS_W}val', '0')
        if f_lang:
            lang_t = ET.SubElement(rPr_tab, f'{NS_W}lang')
            for attr, val in f_lang.items():
                lang_t.set(f'{NS_W}{attr}', val)
        ET.SubElement(r_tab, f'{NS_W}tab')

        omml = self.latex_converter.convert(formula)
        if omml is not None:
            p.append(omml)

        # Tab after formula
        r_tab2 = ET.SubElement(p, f'{NS_W}r')
        rPr_tab2 = ET.SubElement(r_tab2, f'{NS_W}rPr')
        rF_tab2 = ET.SubElement(rPr_tab2, f'{NS_W}rFonts')
        if hint:
            rF_tab2.set(f'{NS_W}hint', hint)
        rF_tab2.set(f'{NS_W}hAnsi', 'Cambria Math')
        i_el2 = ET.SubElement(rPr_tab2, f'{NS_W}i')
        i_el2.set(f'{NS_W}val', '0')
        if f_lang:
            lang_t2 = ET.SubElement(rPr_tab2, f'{NS_W}lang')
            for attr, val in f_lang.items():
                lang_t2.set(f'{NS_W}{attr}', val)
        ET.SubElement(r_tab2, f'{NS_W}tab')
        return p

    def _list_para(self, text: str, ordered: bool = False) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        if self._body_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', self._body_id)
        hint = self.tmpl_fmt.get('font_hint', '')
        if hint:
            rPr = ET.SubElement(pPr, f'{NS_W}rPr')
            rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
            rFonts.set(f'{NS_W}hint', hint)

        prefix = '• ' if not ordered else ''
        parts = re.split(r'(\$[^$]+\$)', text)
        first = True
        for part in parts:
            if part.startswith('$') and part.endswith('$'):
                omml = self.latex_converter.convert(part[1:-1])
                if omml is not None:
                    p.append(omml)
            else:
                if first and prefix:
                    part = prefix + part
                    first = False
                self._add_rich_runs(p, part)
        return p

    def _code_para(self, code_text: str) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        if self._body_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', self._body_id)
        for j, ln in enumerate(code_text.split('\n')):
            if j > 0:
                br_r = ET.SubElement(p, f'{NS_W}r')
                ET.SubElement(br_r, f'{NS_W}br')
            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            rF = ET.SubElement(rPr, f'{NS_W}rFonts')
            rF.set(f'{NS_W}ascii', 'Courier New')
            rF.set(f'{NS_W}hAnsi', 'Courier New')
            t = ET.SubElement(r, f'{NS_W}t')
            if ln.startswith(' ') or ln.endswith(' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = ln if ln else ''
        return p

    # ------------------------------------------------------------------
    # Table / figure caption paragraph
    # ------------------------------------------------------------------

    def _table_caption_para(self, text: str) -> ET.Element:
        """Create a centered caption paragraph using the template's caption format."""
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')

        cap_fmt = self.tmpl_fmt.get('table_caption_format', {})
        style = cap_fmt.get('style', self._body_id)
        if style:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', style)

        ind = cap_fmt.get('ind')
        if ind:
            ind_el = ET.SubElement(pPr, f'{NS_W}ind')
            for attr, val in ind.items():
                ind_el.set(f'{NS_W}{attr}', val)

        jc = ET.SubElement(pPr, f'{NS_W}jc')
        jc.set(f'{NS_W}val', cap_fmt.get('jc', 'center'))

        hint = self.tmpl_fmt.get('font_hint', '')
        rPr_pPr = ET.SubElement(pPr, f'{NS_W}rPr')
        if hint:
            rF = ET.SubElement(rPr_pPr, f'{NS_W}rFonts')
            rF.set(f'{NS_W}hint', hint)
        if cap_fmt.get('bold'):
            ET.SubElement(rPr_pPr, f'{NS_W}b')
            ET.SubElement(rPr_pPr, f'{NS_W}bCs')
        sz = cap_fmt.get('sz')
        if sz:
            ET.SubElement(rPr_pPr, f'{NS_W}sz').set(f'{NS_W}val', sz)
        szCs = cap_fmt.get('szCs')
        if szCs:
            ET.SubElement(rPr_pPr, f'{NS_W}szCs').set(f'{NS_W}val', szCs)

        r = ET.SubElement(p, f'{NS_W}r')
        rPr = ET.SubElement(r, f'{NS_W}rPr')
        if hint:
            rF2 = ET.SubElement(rPr, f'{NS_W}rFonts')
            rF2.set(f'{NS_W}hint', hint)
        if cap_fmt.get('bold'):
            ET.SubElement(rPr, f'{NS_W}b')
            ET.SubElement(rPr, f'{NS_W}bCs')
        if sz:
            ET.SubElement(rPr, f'{NS_W}sz').set(f'{NS_W}val', sz)
        if szCs:
            ET.SubElement(rPr, f'{NS_W}szCs').set(f'{NS_W}val', szCs)

        t = ET.SubElement(r, f'{NS_W}t')
        t.text = text
        return p

    # ------------------------------------------------------------------
    # Reference paragraph
    # ------------------------------------------------------------------

    def _reference_para(self, text: str) -> ET.Element:
        """Create a reference paragraph using template's reference style if available."""
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        style_id = self._ref_style_id or self._body_id
        if style_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', style_id)
        hint = self.tmpl_fmt.get('font_hint', '')
        if hint:
            rPr = ET.SubElement(pPr, f'{NS_W}rPr')
            rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
            rFonts.set(f'{NS_W}hint', hint)
        self._add_rich_runs(p, text)
        return p

    # ------------------------------------------------------------------
    # Table
    # ------------------------------------------------------------------

    def _create_table(self, table_lines: List[str]) -> Optional[ET.Element]:
        if len(table_lines) < 2:
            return None
        headers = [h.strip() for h in table_lines[0].split('|')[1:-1]]
        if not headers:
            return None
        data_lines = table_lines[2:]
        rows: List[List[str]] = []
        for ln in data_lines:
            if ln.strip():
                cells = [c.strip() for c in ln.split('|')[1:-1]]
                if cells:
                    rows.append(cells)

        tbl_fmt = self.tmpl_fmt.get('table_format', {})

        tbl = ET.Element(f'{NS_W}tbl')
        tblPr = ET.SubElement(tbl, f'{NS_W}tblPr')

        style = tbl_fmt.get('style', self._table_style)
        if style:
            ts = ET.SubElement(tblPr, f'{NS_W}tblStyle')
            ts.set(f'{NS_W}val', style)

        tW = ET.SubElement(tblPr, f'{NS_W}tblW')
        tW.set(f'{NS_W}w', tbl_fmt.get('width_w', '5000'))
        tW.set(f'{NS_W}type', tbl_fmt.get('width_type', 'pct'))

        jc = ET.SubElement(tblPr, f'{NS_W}jc')
        jc.set(f'{NS_W}val', tbl_fmt.get('jc', 'center'))

        layout = ET.SubElement(tblPr, f'{NS_W}tblLayout')
        layout.set(f'{NS_W}type', tbl_fmt.get('layout', 'autofit'))

        borders = tbl_fmt.get('borders')
        if borders:
            tblBorders = ET.SubElement(tblPr, f'{NS_W}tblBorders')
            for pos, attrs in borders.items():
                b = ET.SubElement(tblBorders, f'{NS_W}{pos}')
                for attr, val in attrs.items():
                    b.set(f'{NS_W}{attr}', val)
        else:
            tblBorders = ET.SubElement(tblPr, f'{NS_W}tblBorders')
            for pos in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
                b = ET.SubElement(tblBorders, f'{NS_W}{pos}')
                b.set(f'{NS_W}val', 'single')
                b.set(f'{NS_W}sz', '4')
                b.set(f'{NS_W}space', '0')
                b.set(f'{NS_W}color', 'auto')

        look = tbl_fmt.get('look')
        if look:
            tblLook = ET.SubElement(tblPr, f'{NS_W}tblLook')
            for attr, val in look.items():
                tblLook.set(f'{NS_W}{attr}', val)

        grid = ET.SubElement(tbl, f'{NS_W}tblGrid')
        for _ in headers:
            gc = ET.SubElement(grid, f'{NS_W}gridCol')
            gc.set(f'{NS_W}w', '0')

        cell_style = tbl_fmt.get('cell_style', self._body_id)
        row_fmt = tbl_fmt.get('row_format')
        self._add_table_row(tbl, headers, bold=True, cell_style=cell_style,
                            row_fmt=row_fmt, row_position='header')
        for ri, row in enumerate(rows):
            pos = 'last' if ri == len(rows) - 1 else 'body'
            self._add_table_row(tbl, row, bold=False, cell_style=cell_style,
                                row_fmt=row_fmt, row_position=pos)
        return tbl

    def _add_table_row(self, tbl, cells, bold=False, cell_style='',
                       row_fmt=None, row_position='body'):
        hint = self.tmpl_fmt.get('font_hint', '')
        tbl_fmt = self.tmpl_fmt.get('table_format', {})
        cell_borders_map = tbl_fmt.get('cell_borders', {})
        cell_borders = cell_borders_map.get(row_position, cell_borders_map.get('body'))
        cell_shd = tbl_fmt.get('cell_shd')
        cell_ind = tbl_fmt.get('cell_ind')
        cell_rPr_tmpl = tbl_fmt.get('cell_rPr', {})

        tr = ET.SubElement(tbl, f'{NS_W}tr')
        if row_fmt:
            trPr = ET.SubElement(tr, f'{NS_W}trPr')
            for rname, rattrs in row_fmt.items():
                el = ET.SubElement(trPr, f'{NS_W}{rname}')
                for attr, val in rattrs.items():
                    el.set(f'{NS_W}{attr}', val)
        for cell_text in cells:
            tc = ET.SubElement(tr, f'{NS_W}tc')
            tcPr = ET.SubElement(tc, f'{NS_W}tcPr')
            tcW = ET.SubElement(tcPr, f'{NS_W}tcW')
            tcW.set(f'{NS_W}w', '0')
            tcW.set(f'{NS_W}type', 'auto')

            if cell_borders:
                tcBorders = ET.SubElement(tcPr, f'{NS_W}tcBorders')
                for pos, attrs in cell_borders.items():
                    b = ET.SubElement(tcBorders, f'{NS_W}{pos}')
                    for attr, val in attrs.items():
                        b.set(f'{NS_W}{attr}', val)

            if cell_shd:
                shd = ET.SubElement(tcPr, f'{NS_W}shd')
                for attr, val in cell_shd.items():
                    shd.set(f'{NS_W}{attr}', val)

            p = ET.SubElement(tc, f'{NS_W}p')
            p.set(f'{NS_W14}paraId', self._next_para_id())
            pPr = ET.SubElement(p, f'{NS_W}pPr')
            if cell_style:
                pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
                pStyle.set(f'{NS_W}val', cell_style)
            if cell_ind:
                ind_el = ET.SubElement(pPr, f'{NS_W}ind')
                for attr, val in cell_ind.items():
                    ind_el.set(f'{NS_W}{attr}', val)
            jc_el = ET.SubElement(pPr, f'{NS_W}jc')
            jc_el.set(f'{NS_W}val', 'center')

            if cell_rPr_tmpl:
                pPr_rPr = ET.SubElement(pPr, f'{NS_W}rPr')
                if hint:
                    rF_p = ET.SubElement(pPr_rPr, f'{NS_W}rFonts')
                    rF_p.set(f'{NS_W}hint', hint)
                for rpr_name in ('b', 'color', 'sz', 'szCs', 'vertAlign'):
                    rpr_info = cell_rPr_tmpl.get(rpr_name)
                    if rpr_info is not None:
                        el = ET.SubElement(pPr_rPr, f'{NS_W}{rpr_name}')
                        if isinstance(rpr_info, dict):
                            for attr, val in rpr_info.items():
                                el.set(f'{NS_W}{attr}', val)

            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            if hint:
                rF = ET.SubElement(rPr, f'{NS_W}rFonts')
                rF.set(f'{NS_W}hint', hint)
            if cell_rPr_tmpl.get('b') is not None:
                b_info = cell_rPr_tmpl['b']
                b_el = ET.SubElement(rPr, f'{NS_W}b')
                if isinstance(b_info, dict) and 'val' in b_info:
                    b_el.set(f'{NS_W}val', b_info['val'])
            elif bold:
                ET.SubElement(rPr, f'{NS_W}b')
            for rpr_name in ('color', 'sz', 'szCs', 'vertAlign'):
                rpr_info = cell_rPr_tmpl.get(rpr_name)
                if rpr_info and isinstance(rpr_info, dict):
                    el = ET.SubElement(rPr, f'{NS_W}{rpr_name}')
                    for attr, val in rpr_info.items():
                        el.set(f'{NS_W}{attr}', val)

            t = ET.SubElement(r, f'{NS_W}t')
            t.text = cell_text

    # ------------------------------------------------------------------
    # Image
    # ------------------------------------------------------------------

    def _image_para(self, md_line: str) -> ET.Element:
        m = re.match(r'!\[([^\]]*)\]\(([^)]+)\)', md_line.strip())
        caption = m.group(1) if m else ''
        img_path = m.group(2) if m else ''

        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        if self._image_style_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', self._image_style_id)
        jc = ET.SubElement(pPr, f'{NS_W}jc')
        jc.set(f'{NS_W}val', 'center')

        if not os.path.isfile(img_path):
            r = ET.SubElement(p, f'{NS_W}r')
            t = ET.SubElement(r, f'{NS_W}t')
            t.text = f'[Image not found: {img_path}]'
            return p

        max_w = self.tmpl_fmt.get('img_max_width_emu', 5400000)
        w_emu, h_emu = _image_size_emu(img_path, max_width_emu=max_w)
        rid = self._next_rel_id()

        ext = os.path.splitext(img_path)[1].lower()
        media_name = f'image_{rid}{ext}'
        self._image_rels.append({
            'rId': rid,
            'target': f'media/{media_name}',
            'src_path': os.path.abspath(img_path),
            'media_name': media_name,
        })

        temp_dir_match = re.search(r'/tmp/docx_ew_[a-f0-9]+', str(self._image_rels))
        if self._image_rels:
            last = self._image_rels[-1]
            last['_copy_needed'] = True

        r = ET.SubElement(p, f'{NS_W}r')
        drawing = ET.SubElement(r, f'{NS_W}drawing')
        inline = ET.SubElement(drawing, f'{NS_WP}inline')
        inline.set('distT', '0')
        inline.set('distB', '0')
        inline.set('distL', '0')
        inline.set('distR', '0')
        extent = ET.SubElement(inline, f'{NS_WP}extent')
        extent.set('cx', str(w_emu))
        extent.set('cy', str(h_emu))
        docPr = ET.SubElement(inline, f'{NS_WP}docPr')
        docPr.set('id', str(self._rel_id_counter))
        docPr.set('name', caption or f'Image {self._rel_id_counter}')
        graphic = ET.SubElement(inline, f'{NS_A}graphic')
        graphicData = ET.SubElement(graphic, f'{NS_A}graphicData')
        graphicData.set('uri', 'http://schemas.openxmlformats.org/drawingml/2006/picture')
        pic_el = ET.SubElement(graphicData, f'{NS_PIC}pic')
        nvPicPr = ET.SubElement(pic_el, f'{NS_PIC}nvPicPr')
        cNvPr = ET.SubElement(nvPicPr, f'{NS_PIC}cNvPr')
        cNvPr.set('id', str(self._rel_id_counter))
        cNvPr.set('name', media_name)
        ET.SubElement(nvPicPr, f'{NS_PIC}cNvPicPr')
        blipFill = ET.SubElement(pic_el, f'{NS_PIC}blipFill')
        blip = ET.SubElement(blipFill, f'{NS_A}blip')
        blip.set(f'{NS_R_DOC}embed', rid)
        stretch = ET.SubElement(blipFill, f'{NS_A}stretch')
        ET.SubElement(stretch, f'{NS_A}fillRect')
        spPr = ET.SubElement(pic_el, f'{NS_PIC}spPr')
        xfrm = ET.SubElement(spPr, f'{NS_A}xfrm')
        off = ET.SubElement(xfrm, f'{NS_A}off')
        off.set('x', '0')
        off.set('y', '0')
        ext_el = ET.SubElement(xfrm, f'{NS_A}ext')
        ext_el.set('cx', str(w_emu))
        ext_el.set('cy', str(h_emu))
        prstGeom = ET.SubElement(spPr, f'{NS_A}prstGeom')
        prstGeom.set('prst', 'rect')

        return p

    # ------------------------------------------------------------------
    # Document type detection
    # ------------------------------------------------------------------

    def _is_form_template(self, temp_dir: str) -> bool:
        """Check if template is form-type (no heading-style paragraphs in body)."""
        sx = StyleExtractor(temp_dir)
        heading_sids = set(sx.heading_ids().values())
        if not heading_sids:
            return True
        doc_root = sx.doc_root
        if doc_root is None:
            return True
        body = doc_root.find(f'{NS_W}body')
        if body is None:
            return True
        for p in body.findall(f'{NS_W}p'):
            pPr = p.find(f'{NS_W}pPr')
            if pPr is not None:
                ps = pPr.find(f'{NS_W}pStyle')
                if ps is not None and ps.get(f'{NS_W}val', '') in heading_sids:
                    return False
        return True

    # ------------------------------------------------------------------
    # Form-mode: template fill
    # ------------------------------------------------------------------

    @staticmethod
    def _parse_md_sections(markdown: str) -> Dict[str, List[str]]:
        """Parse Markdown into sections keyed by their header label.

        Handles both ``## N、label`` and ``N、label`` as section delimiters.
        Returns {normalised_label: [content_lines]}.
        """
        section_re = re.compile(r'^#{1,3}\s+(.+)')
        numbered_re = re.compile(r'^(\d+)\s*[、.．]')

        sections: Dict[str, List[str]] = {}
        current_label: Optional[str] = None
        current_lines: List[str] = []

        for line in markdown.split('\n'):
            stripped = line.strip()

            # Check for Markdown heading that looks like a section label
            hm = section_re.match(stripped)
            if hm:
                inner = hm.group(1).strip()
                if numbered_re.match(inner):
                    if current_label is not None:
                        sections[current_label] = current_lines
                    current_label = inner
                    current_lines = []
                    continue

            # Check for bare numbered section header (no # prefix)
            if numbered_re.match(stripped) and len(stripped) > 5:
                if current_label is not None:
                    sections[current_label] = current_lines
                current_label = stripped
                current_lines = []
                continue

            if current_label is not None:
                current_lines.append(line.rstrip())

        if current_label is not None:
            sections[current_label] = current_lines

        return sections

    @staticmethod
    def _clean_form_content(lines: List[str]) -> List[str]:
        """Clean Markdown-specific syntax from content lines for form filling.

        Removes: horizontal rules, table separators/headers, bold markers, heading markers.
        Converts Markdown table data rows to plain text.
        """
        separator_re = re.compile(r'^-{3,}\s*$')
        table_sep_re = re.compile(r'^\|[\s\-:|]+\|$')
        table_row_re = re.compile(r'^\|(.+)\|$')
        footer_re = re.compile(
            r'^\*{0,2}(备\s*注|指导教师|教研室|年\s+月\s+日)\*{0,2}',
        )

        # Pre-scan: identify table header rows (row immediately before a separator)
        stripped_lines = [l.strip() for l in lines]
        table_header_indices: set = set()
        for i, sl in enumerate(stripped_lines):
            if table_sep_re.match(sl) and i > 0 and table_row_re.match(stripped_lines[i - 1]):
                table_header_indices.add(i - 1)

        cleaned: List[str] = []

        for li, line in enumerate(lines):
            stripped = stripped_lines[li]
            if not stripped:
                continue
            if separator_re.match(stripped):
                continue
            if table_sep_re.match(stripped):
                continue
            if li in table_header_indices:
                continue
            if footer_re.match(stripped.lstrip('*').strip()):
                break

            # Convert table data rows to joined plain text
            tm = table_row_re.match(stripped)
            if tm:
                cells = [c.strip() for c in tm.group(1).split('|') if c.strip()]
                cleaned.append('  '.join(cells))
                continue

            text = re.sub(r'\*\*([^*]+)\*\*', r'\1', stripped)
            text = re.sub(r'^#{1,3}\s+', '', text)
            cleaned.append(text)

        return cleaned

    @staticmethod
    def _normalise_label(text: str) -> str:
        """Normalise section label for fuzzy matching: strip spaces/punctuation."""
        return re.sub(r'\s+', '', text).rstrip('：:')

    def _match_section_label(self, md_label: str,
                             template_labels: List[str]) -> Optional[str]:
        """Find the best matching template label for a Markdown section label."""
        norm_md = self._normalise_label(md_label)
        # Extract just the numbered prefix + core title for matching
        m = re.match(r'(\d+[、.．])', md_label)
        prefix = m.group(1) if m else ''

        best_match = None
        best_score = 0
        for tl in template_labels:
            norm_tl = self._normalise_label(tl)
            if norm_md == norm_tl:
                return tl
            # Prefix match: same section number
            if prefix and tl.lstrip().startswith(prefix.rstrip()):
                score = len(set(norm_md) & set(norm_tl)) / max(len(norm_md), len(norm_tl), 1)
                if score > best_score:
                    best_score = score
                    best_match = tl
        return best_match if best_score > 0.3 else None

    @staticmethod
    def _para_text(p_el: ET.Element) -> str:
        """Get concatenated text from a paragraph."""
        parts = []
        for t in p_el.iter(f'{NS_W}t'):
            if t.text:
                parts.append(t.text)
        return ''.join(parts)

    @staticmethod
    def _clone_run_rpr(source_run: ET.Element) -> Optional[ET.Element]:
        """Deep-copy the w:rPr from a run element."""
        rPr = source_run.find(f'{NS_W}rPr')
        if rPr is not None:
            import copy
            return copy.deepcopy(rPr)
        return None

    @staticmethod
    def _find_content_rpr(para: ET.Element) -> Optional[ET.Element]:
        """Find the rPr most likely representing the paragraph's *content* formatting.

        Heuristic: the run carrying the most text characters is the actual content
        run — label prefixes and trailing metadata are typically shorter.
        Falls back to the last run if no run contains text.
        """
        import copy
        runs = para.findall(f'{NS_W}r')
        if not runs:
            return None

        best_run = None
        best_len = -1
        for r in runs:
            text_len = sum(len(t.text or '') for t in r.iter(f'{NS_W}t'))
            if text_len > best_len:
                best_len = text_len
                best_run = r

        if best_run is None:
            best_run = runs[-1]

        rPr = best_run.find(f'{NS_W}rPr')
        return copy.deepcopy(rPr) if rPr is not None else None

    def _replace_para_content_runs(self, para: ET.Element,
                                   new_text: str,
                                   keep_label_prefix: bool = False,
                                   label_text: str = '') -> None:
        """Replace a paragraph's content runs with new text, preserving formatting.

        If keep_label_prefix is True, keeps runs matching label_text and replaces
        only subsequent runs (for paragraphs where section header and content coexist).
        """
        runs = para.findall(f'{NS_W}r')
        if not runs:
            return

        # Pick the rPr representing the content area (longest-text heuristic)
        ref_rpr = self._find_content_rpr(para)

        if keep_label_prefix and label_text:
            # Find where the label ends in the runs
            accumulated = ''
            label_norm = re.sub(r'\s+', '', label_text)
            split_idx = len(runs)
            for ri, run in enumerate(runs):
                for t in run.iter(f'{NS_W}t'):
                    if t.text:
                        accumulated += t.text
                acc_norm = re.sub(r'\s+', '', accumulated)
                if label_norm and acc_norm.startswith(label_norm):
                    split_idx = ri + 1
                    break

            # Remove content runs (after label)
            for run in runs[split_idx:]:
                para.remove(run)

            # Add separator space + new content as a single run
            new_run = ET.SubElement(para, f'{NS_W}r')
            if ref_rpr is not None:
                new_run.append(ref_rpr)
            t = ET.SubElement(new_run, f'{NS_W}t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = '  ' + new_text
        else:
            # Replace ALL runs with new text
            for run in runs:
                para.remove(run)
            new_run = ET.SubElement(para, f'{NS_W}r')
            if ref_rpr is not None:
                new_run.append(ref_rpr)
            t = ET.SubElement(new_run, f'{NS_W}t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = new_text

    def _create_para_like(self, ref_para: ET.Element, text: str) -> ET.Element:
        """Create a new paragraph with same pPr/rPr as ref_para, but with new text."""
        import copy
        new_p = ET.Element(f'{NS_W}p')

        # Copy paragraph properties
        pPr = ref_para.find(f'{NS_W}pPr')
        if pPr is not None:
            new_p.append(copy.deepcopy(pPr))

        # Use the content-representative rPr from the reference paragraph
        ref_rpr = self._find_content_rpr(ref_para)

        new_run = ET.SubElement(new_p, f'{NS_W}r')
        if ref_rpr is not None:
            new_run.append(ref_rpr)
        t = ET.SubElement(new_run, f'{NS_W}t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        return new_p

    def _fill_template(self, markdown: str, temp_dir: str) -> None:
        """Fill a form-type template with content from Markdown.

        Instead of generating new document.xml, modifies the template's existing
        XML in-place — preserving all paragraph/run formatting from the template.
        """
        from analyze_template import TemplateAnalyzer

        # Analyse the template to get content structure
        analyzer = TemplateAnalyzer(self.template_path)
        cs = analyzer.extract_content_structure()
        analyzer.close()

        template_sections = cs.get('sections', [])
        if not template_sections:
            print("Warning: no numbered sections found in template, "
                  "falling back to essay mode.", file=sys.stderr)
            return None  # signal to caller to fall back

        # Parse Markdown into sections
        md_sections = self._parse_md_sections(markdown)

        # Build label → template section mapping
        template_labels = [s['label'] for s in template_sections]

        # Register namespace prefixes so ET serializes with correct prefixes
        ns_registrations = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
            'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
            'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
            'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
            'o': 'urn:schemas-microsoft-com:office:office',
            'v': 'urn:schemas-microsoft-com:vml',
            'wpc': 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas',
            'w10': 'urn:schemas-microsoft-com:office:word',
            'w15': 'http://schemas.microsoft.com/office/word/2012/wordml',
            'wpg': 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup',
            'wpi': 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk',
            'wne': 'http://schemas.microsoft.com/office/word/2006/wordml',
            'wps': 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape',
            'wp14': 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing',
        }
        for prefix, uri in ns_registrations.items():
            ET.register_namespace(prefix, uri)

        # Also register any custom namespaces from the template
        doc_path = os.path.join(temp_dir, 'word', 'document.xml')
        with open(doc_path, 'r', encoding='utf-8') as f:
            raw_xml = f.read()
        for m in re.finditer(r'xmlns:(\w+)="([^"]+)"', raw_xml):
            prefix, uri = m.group(1), m.group(2)
            if prefix not in ns_registrations:
                ET.register_namespace(prefix, uri)

        tree = ET.parse(doc_path)
        doc_root = tree.getroot()
        body = doc_root.find(f'{NS_W}body')
        if body is None:
            return None

        # Index body children for fast lookup
        body_children = list(body)
        idx_to_child: Dict[int, ET.Element] = {}
        for bi, child in enumerate(body_children):
            idx_to_child[bi] = child

        # Process each Markdown section
        for md_label, md_lines in md_sections.items():
            tmpl_label = self._match_section_label(md_label, template_labels)
            if tmpl_label is None:
                print(f"Warning: no template section matches '{md_label[:40]}', skipping.",
                      file=sys.stderr)
                continue

            # Find the template section info
            sec_info = next(s for s in template_sections if s['label'] == tmpl_label)
            content_indices = sec_info['content_para_indices']
            header_has_content = sec_info['header_has_content']
            header_idx = sec_info['header_para_index']

            # Clean and filter content lines
            new_paragraphs = self._clean_form_content(md_lines)
            if not new_paragraphs:
                continue

            if header_has_content:
                # Section header + content in same paragraph
                header_para = idx_to_child.get(header_idx)
                if header_para is None:
                    continue
                # Replace content runs in header para, keep label prefix
                self._replace_para_content_runs(
                    header_para, new_paragraphs[0],
                    keep_label_prefix=True,
                    label_text=tmpl_label)
                extra_paragraphs = new_paragraphs[1:]

                if extra_paragraphs:
                    # Insert additional paragraphs after header
                    insert_pos = list(body).index(header_para) + 1
                    for ep_text in reversed(extra_paragraphs):
                        new_p = self._create_para_like(header_para, ep_text)
                        body.insert(insert_pos, new_p)
            else:
                # Content in separate paragraphs
                if not content_indices:
                    continue

                # Get reference formatting from first content paragraph
                first_content_para = idx_to_child.get(content_indices[0])
                if first_content_para is None:
                    continue

                # Determine insert position (position of first content paragraph)
                insert_pos = list(body).index(first_content_para)

                # Remove old content paragraphs
                for ci in content_indices:
                    old_para = idx_to_child.get(ci)
                    if old_para is not None and old_para in list(body):
                        body.remove(old_para)

                # Insert new content paragraphs at the same position
                for pi, para_text in enumerate(new_paragraphs):
                    new_p = self._create_para_like(first_content_para, para_text)
                    body.insert(insert_pos + pi, new_p)

        # Write back modified document.xml with proper namespace prefixes
        xml_str = ET.tostring(doc_root, encoding='unicode', xml_declaration=False)

        # Fix any remaining nsN: prefixes that ET might generate
        for i in range(20):
            prefix = f'ns{i}'
            # Find what URI this nsN maps to
            m_uri = re.search(rf'xmlns:{prefix}="([^"]+)"', xml_str)
            if not m_uri:
                continue
            uri = m_uri.group(1)
            # Find the correct prefix for this URI
            correct = None
            for p, u in ns_registrations.items():
                if u == uri:
                    correct = p
                    break
            if correct and correct != prefix:
                xml_str = xml_str.replace(f'xmlns:{prefix}=', f'xmlns:{correct}=')
                xml_str = xml_str.replace(f'<{prefix}:', f'<{correct}:')
                xml_str = xml_str.replace(f'</{prefix}:', f'</{correct}:')
                xml_str = xml_str.replace(f' {prefix}:', f' {correct}:')

        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n')
            f.write(xml_str)

        return True  # success

    # ------------------------------------------------------------------
    # Main conversion entry point
    # ------------------------------------------------------------------

    def convert(self, markdown_content: str, output_path: str):
        temp_dir = '/tmp/docx_ew_' + hashlib.md5(output_path.encode()).hexdigest()[:8]
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

        with zipfile.ZipFile(self.template_path, 'r') as z:
            z.extractall(temp_dir)

        # Detect document type: form vs essay
        is_form = self._is_form_template(temp_dir)

        if is_form:
            result = self._fill_template(markdown_content, temp_dir)
            if result is None:
                # Fallback to essay mode if form fill fails
                is_form = False

        if not is_form:
            # Essay mode: generate new document.xml from Markdown
            self.tmpl_fmt = extract_template_formats(temp_dir)

            sx = StyleExtractor(temp_dir)
            self._heading_ids = sx.heading_ids()
            self._body_id = self.tmpl_fmt.get('body_style_id', '') or sx.body_style_id()
            self._formula_id = sx.find_style_id('formula', '公式', 'Equation')
            self._ref_style_id = self.tmpl_fmt.get('reference_style_id', '')
            self._image_style_id = self.tmpl_fmt.get('image_style_id', '')
            self._table_style = self.tmpl_fmt.get('table_style_id', '') or sx.find_style_id(
                'Normal Table', '普通表格', 'Table Grid', '网格型')
            self._sect_pr_xml = sx.extract_sect_pr_xml()
            self._image_rels = []

            font_hint = self.tmpl_fmt.get('font_hint', '')
            self.latex_converter = LatexToOmmlConverter(font_hint=font_hint)

            doc_xml = self._generate_document_xml(markdown_content)
            doc_path = os.path.join(temp_dir, 'word', 'document.xml')
            with open(doc_path, 'w', encoding='utf-8') as f:
                f.write(doc_xml)

            # Copy image files into word/media/
            media_dir = os.path.join(temp_dir, 'word', 'media')
            os.makedirs(media_dir, exist_ok=True)
            for img_rel in self._image_rels:
                src = img_rel.get('src_path', '')
                if src and os.path.isfile(src):
                    shutil.copy2(src, os.path.join(media_dir, img_rel['media_name']))

            if self._image_rels:
                self._update_rels(temp_dir)
                self._update_content_types(temp_dir)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root_d, _dirs, files in os.walk(temp_dir):
                for fn in files:
                    fp = os.path.join(root_d, fn)
                    arcname = os.path.relpath(fp, temp_dir)
                    zf.write(fp, arcname)

        shutil.rmtree(temp_dir)
        mode_str = "form-fill" if is_form else "essay"
        print(f"Converted ({mode_str} mode) to: {output_path}")

    def _update_content_types(self, temp_dir: str):
        ct_path = os.path.join(temp_dir, '[Content_Types].xml')
        if not os.path.isfile(ct_path):
            return
        tree = ET.parse(ct_path)
        root = tree.getroot()
        ns = 'http://schemas.openxmlformats.org/package/2006/content-types'
        existing_exts = set()
        for el in root.findall(f'{{{ns}}}Default'):
            existing_exts.add(el.get('Extension', '').lower())

        img_ext_map = {
            '.png': 'image/png',
            '.jpg': 'image/jpeg',
            '.jpeg': 'image/jpeg',
            '.gif': 'image/gif',
            '.bmp': 'image/bmp',
            '.tiff': 'image/tiff',
            '.tif': 'image/tiff',
        }
        for img_rel in self._image_rels:
            ext = os.path.splitext(img_rel['media_name'])[1].lower()
            if ext.lstrip('.') not in existing_exts and ext in img_ext_map:
                default = ET.SubElement(root, f'{{{ns}}}Default')
                default.set('Extension', ext.lstrip('.'))
                default.set('ContentType', img_ext_map[ext])
                existing_exts.add(ext.lstrip('.'))

        tree.write(ct_path, xml_declaration=True, encoding='UTF-8')


# ===================================================================
# CLI
# ===================================================================

def main():
    parser = argparse.ArgumentParser(
        description='Convert Markdown to DOCX using a template for formatting'
    )
    parser.add_argument('--template', '-t', required=True,
                        help='Template DOCX file')
    parser.add_argument('--markdown', '-m', required=True,
                        help='Markdown input file')
    parser.add_argument('--output', '-o', required=True,
                        help='Output DOCX file path')
    args = parser.parse_args()

    with open(args.markdown, 'r', encoding='utf-8') as f:
        md_content = f.read()

    converter = MarkdownToDocxConverter(args.template)
    converter.convert(md_content, args.output)


if __name__ == '__main__':
    main()
