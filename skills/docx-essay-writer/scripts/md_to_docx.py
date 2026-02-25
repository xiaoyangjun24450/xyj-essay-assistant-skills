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

    def _ctrl_pr(self):
        ctrlPr = ET.Element(f'{NS_M}ctrlPr')
        rPr = ET.SubElement(ctrlPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        rFonts.set(f'{NS_W}ascii', 'Cambria Math')
        rFonts.set(f'{NS_W}hAnsi', 'Cambria Math')
        return ctrlPr

    def _math_run(self, text, hint='default'):
        r = ET.Element(f'{NS_M}r')
        rPr = ET.SubElement(r, f'{NS_M}rPr')
        sty = ET.SubElement(rPr, f'{NS_M}sty')
        sty.set(f'{NS_M}val', 'p')
        wrPr = ET.SubElement(r, f'{NS_W}rPr')
        rFonts = ET.SubElement(wrPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', hint)
        rFonts.set(f'{NS_W}ascii', 'Cambria Math')
        rFonts.set(f'{NS_W}hAnsi', 'Cambria Math')
        t = ET.SubElement(r, f'{NS_M}t')
        t.text = text
        return r

    def _math_run_ea(self, text):
        return self._math_run(text, hint='eastAsia')

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

        # 5. Fractions \frac{a}{b}
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

        # 7. Subscript
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
                self._build_sub(parent, m.group(1), m.group(2))
                rest = expr[m.end():]
                if rest:
                    self._parse_expr(parent, rest)
                return

        # 8. Superscript
        sup_pats = [
            r'\\([a-zA-Z]+)\^\{([^}]+)\}',
            r'\\([a-zA-Z]+)\^([a-zA-Z0-9*])',
            r'([a-zA-Z0-9])\^\{([^}]+)\}',
            r'([a-zA-Z0-9])\^([a-zA-Z0-9*])',
        ]
        for pat in sup_pats:
            m = re.match(pat, expr)
            if m:
                self._build_sup(parent, m.group(1), m.group(2))
                rest = expr[m.end():]
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
        if expr[0] in '=+-*/()[]{},;: ':
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
    # Main entry point
    # ------------------------------------------------------------------

    def convert(self, markdown_content: str, output_path: str):
        temp_dir = '/tmp/docx_ew_' + hashlib.md5(output_path.encode()).hexdigest()[:8]
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

        with zipfile.ZipFile(self.template_path, 'r') as z:
            z.extractall(temp_dir)

        sx = StyleExtractor(temp_dir)
        self._heading_ids = sx.heading_ids()
        self._body_id = sx.body_style_id()
        self._formula_id = sx.find_style_id('formula', '公式', 'Equation')
        self._table_style = sx.find_style_id('Normal Table', '普通表格',
                                              'Table Grid', '网格型')
        self._sect_pr_xml = sx.extract_sect_pr_xml()

        doc_xml = self._generate_document_xml(markdown_content)
        doc_path = os.path.join(temp_dir, 'word', 'document.xml')
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(doc_xml)

        if self._image_rels:
            self._update_rels(temp_dir)

        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zf:
            for root, _dirs, files in os.walk(temp_dir):
                for fn in files:
                    fp = os.path.join(root, fn)
                    arcname = os.path.relpath(fp, temp_dir)
                    zf.write(fp, arcname)

        shutil.rmtree(temp_dir)
        print(f"Converted to: {output_path}")

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
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        return p

    def _heading_para(self, level: int, text: str) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        sid = self._heading_ids.get(level, '')
        if sid:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', sid)
        if level >= 2:
            numPr = ET.SubElement(pPr, f'{NS_W}numPr')
            ilvl = ET.SubElement(numPr, f'{NS_W}ilvl')
            ilvl.set(f'{NS_W}val', '1')
            numId = ET.SubElement(numPr, f'{NS_W}numId')
            numId.set(f'{NS_W}val', '0')
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        r = ET.SubElement(p, f'{NS_W}r')
        rPr2 = ET.SubElement(r, f'{NS_W}rPr')
        rFonts2 = ET.SubElement(rPr2, f'{NS_W}rFonts')
        rFonts2.set(f'{NS_W}hint', 'eastAsia')
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
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        lang = ET.SubElement(rPr, f'{NS_W}lang')
        lang.set(f'{NS_W}eastAsia', 'zh-CN')

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
        segments = re.split(r'(\*\*[^*]+\*\*|\*[^*]+\*)', text)
        for seg in segments:
            if not seg or not seg.strip():
                if seg and (' ' in seg or '\t' in seg):
                    r = ET.SubElement(p, f'{NS_W}r')
                    rPr = ET.SubElement(r, f'{NS_W}rPr')
                    rF = ET.SubElement(rPr, f'{NS_W}rFonts')
                    rF.set(f'{NS_W}hint', 'eastAsia')
                    t = ET.SubElement(r, f'{NS_W}t')
                    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    t.text = seg
                continue
            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            rF = ET.SubElement(rPr, f'{NS_W}rFonts')
            rF.set(f'{NS_W}hint', 'eastAsia')
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
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'default')
        lang = ET.SubElement(rPr, f'{NS_W}lang')
        lang.set(f'{NS_W}val', 'en-US')
        lang.set(f'{NS_W}eastAsia', 'zh-CN')

        # Tab before formula
        r_tab = ET.SubElement(p, f'{NS_W}r')
        rPr_tab = ET.SubElement(r_tab, f'{NS_W}rPr')
        rF_tab = ET.SubElement(rPr_tab, f'{NS_W}rFonts')
        rF_tab.set(f'{NS_W}hint', 'eastAsia')
        rF_tab.set(f'{NS_W}hAnsi', 'Cambria Math')
        b_el = ET.SubElement(rPr_tab, f'{NS_W}b')
        b_el.set(f'{NS_W}val', '0')
        i_el = ET.SubElement(rPr_tab, f'{NS_W}i')
        i_el.set(f'{NS_W}val', '0')
        lang_t = ET.SubElement(rPr_tab, f'{NS_W}lang')
        lang_t.set(f'{NS_W}val', 'en-US')
        lang_t.set(f'{NS_W}eastAsia', 'zh-CN')
        ET.SubElement(r_tab, f'{NS_W}tab')

        omml = self.latex_converter.convert(formula)
        if omml is not None:
            p.append(omml)

        # Tab after formula
        r_tab2 = ET.SubElement(p, f'{NS_W}r')
        rPr_tab2 = ET.SubElement(r_tab2, f'{NS_W}rPr')
        rF_tab2 = ET.SubElement(rPr_tab2, f'{NS_W}rFonts')
        rF_tab2.set(f'{NS_W}hint', 'eastAsia')
        rF_tab2.set(f'{NS_W}hAnsi', 'Cambria Math')
        i_el2 = ET.SubElement(rPr_tab2, f'{NS_W}i')
        i_el2.set(f'{NS_W}val', '0')
        lang_t2 = ET.SubElement(rPr_tab2, f'{NS_W}lang')
        lang_t2.set(f'{NS_W}val', 'en-US')
        lang_t2.set(f'{NS_W}eastAsia', 'zh-CN')
        ET.SubElement(r_tab2, f'{NS_W}tab')
        return p

    def _list_para(self, text: str, ordered: bool = False) -> ET.Element:
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._next_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        if self._body_id:
            pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
            pStyle.set(f'{NS_W}val', self._body_id)
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')

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
    # Table
    # ------------------------------------------------------------------

    def _create_table(self, table_lines: List[str]) -> Optional[ET.Element]:
        if len(table_lines) < 2:
            return None
        headers = [h.strip() for h in table_lines[0].split('|')[1:-1]]
        if not headers:
            return None
        data_lines = table_lines[2:]  # skip separator
        rows: List[List[str]] = []
        for ln in data_lines:
            if ln.strip():
                cells = [c.strip() for c in ln.split('|')[1:-1]]
                if cells:
                    rows.append(cells)

        n_cols = len(headers)
        col_w = str(int(9000 / n_cols))

        tbl = ET.Element(f'{NS_W}tbl')
        tblPr = ET.SubElement(tbl, f'{NS_W}tblPr')
        if self._table_style:
            ts = ET.SubElement(tblPr, f'{NS_W}tblStyle')
            ts.set(f'{NS_W}val', self._table_style)
        tW = ET.SubElement(tblPr, f'{NS_W}tblW')
        tW.set(f'{NS_W}w', '5000')
        tW.set(f'{NS_W}type', 'pct')
        jc = ET.SubElement(tblPr, f'{NS_W}jc')
        jc.set(f'{NS_W}val', 'center')
        layout = ET.SubElement(tblPr, f'{NS_W}tblLayout')
        layout.set(f'{NS_W}type', 'fixed')

        # Add table borders
        tblBorders = ET.SubElement(tblPr, f'{NS_W}tblBorders')
        for pos in ('top', 'left', 'bottom', 'right', 'insideH', 'insideV'):
            b = ET.SubElement(tblBorders, f'{NS_W}{pos}')
            b.set(f'{NS_W}val', 'single')
            b.set(f'{NS_W}sz', '4')
            b.set(f'{NS_W}space', '0')
            b.set(f'{NS_W}color', 'auto')

        grid = ET.SubElement(tbl, f'{NS_W}tblGrid')
        for _ in headers:
            gc = ET.SubElement(grid, f'{NS_W}gridCol')
            gc.set(f'{NS_W}w', col_w)

        # Header row
        self._add_table_row(tbl, headers, col_w, bold=True)
        # Data rows
        for row in rows:
            self._add_table_row(tbl, row, col_w, bold=False)
        return tbl

    def _add_table_row(self, tbl, cells, col_w, bold=False):
        tr = ET.SubElement(tbl, f'{NS_W}tr')
        for cell_text in cells:
            tc = ET.SubElement(tr, f'{NS_W}tc')
            tcPr = ET.SubElement(tc, f'{NS_W}tcPr')
            tcW = ET.SubElement(tcPr, f'{NS_W}tcW')
            tcW.set(f'{NS_W}w', col_w)
            tcW.set(f'{NS_W}type', 'dxa')
            p = ET.SubElement(tc, f'{NS_W}p')
            p.set(f'{NS_W14}paraId', self._next_para_id())
            pPr = ET.SubElement(p, f'{NS_W}pPr')
            jc_el = ET.SubElement(pPr, f'{NS_W}jc')
            jc_el.set(f'{NS_W}val', 'center')
            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            rF = ET.SubElement(rPr, f'{NS_W}rFonts')
            rF.set(f'{NS_W}hint', 'eastAsia')
            if bold:
                ET.SubElement(rPr, f'{NS_W}b')
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
        jc = ET.SubElement(pPr, f'{NS_W}jc')
        jc.set(f'{NS_W}val', 'center')

        if not os.path.isfile(img_path):
            r = ET.SubElement(p, f'{NS_W}r')
            t = ET.SubElement(r, f'{NS_W}t')
            t.text = f'[Image not found: {img_path}]'
            return p

        w_emu, h_emu = _image_size_emu(img_path)
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

    # Override convert to also copy images
    def convert(self, markdown_content: str, output_path: str):
        temp_dir = '/tmp/docx_ew_' + hashlib.md5(output_path.encode()).hexdigest()[:8]
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)

        with zipfile.ZipFile(self.template_path, 'r') as z:
            z.extractall(temp_dir)

        sx = StyleExtractor(temp_dir)
        self._heading_ids = sx.heading_ids()
        self._body_id = sx.body_style_id()
        self._formula_id = sx.find_style_id('formula', '公式', 'Equation')
        self._table_style = sx.find_style_id('Normal Table', '普通表格',
                                              'Table Grid', '网格型')
        self._sect_pr_xml = sx.extract_sect_pr_xml()
        self._image_rels = []

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
        print(f"Converted to: {output_path}")

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
