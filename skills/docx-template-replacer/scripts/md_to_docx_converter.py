#!/usr/bin/env python3
"""
Markdown to DOCX Converter for Thesis Template
- Converts Markdown to Word document following the template format exactly
- LaTeX formulas are converted to OMML matching template structure
"""

import re
import os
import shutil
import zipfile
import xml.etree.ElementTree as ET
import hashlib

STYLES = {
    'heading1': '2',
    'heading2': '3',
    'heading3': '4',
    'body': '29',
    'formula': '41',
    'formula_char': '43',
    'table': '46',
}

NS_W = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
NS_R = '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}'
NS_M = '{http://schemas.openxmlformats.org/officeDocument/2006/math}'
NS_W14 = '{http://schemas.microsoft.com/office/word/2010/wordml}'


class LatexToOmmlConverter:
    """Convert LaTeX to OMML exactly like the template"""
    
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
        """Create ctrlPr element like template - simplified"""
        ctrlPr = ET.Element(f'{NS_M}ctrlPr')
        rPr = ET.SubElement(ctrlPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        rFonts.set(f'{NS_W}ascii', 'Cambria Math')
        rFonts.set(f'{NS_W}hAnsi', 'Cambria Math')
        return ctrlPr

    def _math_run(self, text, hint='default'):
        """Create m:r element exactly like template"""
        r = ET.Element(f'{NS_M}r')
        # m:rPr with m:sty like template
        rPr = ET.SubElement(r, f'{NS_M}rPr')
        sty = ET.SubElement(rPr, f'{NS_M}sty')
        sty.set(f'{NS_M}val', 'p')
        # w:rPr with Cambria Math
        wrPr = ET.SubElement(r, f'{NS_W}rPr')
        rFonts = ET.SubElement(wrPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', hint)
        rFonts.set(f'{NS_W}ascii', 'Cambria Math')
        rFonts.set(f'{NS_W}hAnsi', 'Cambria Math')
        t = ET.SubElement(r, f'{NS_M}t')
        t.text = text
        return r

    def _math_run_with_hint(self, text, hint='eastAsia'):
        """Create m:r with hint attribute like template"""
        return self._math_run(text, hint)

    def convert(self, latex):
        latex = latex.strip()
        if not latex:
            return None
        if latex.startswith('$$') and latex.endswith('$$'):
            latex = latex[2:-2].strip()
        elif latex.startswith('$') and latex.endswith('$'):
            latex = latex[1:-1].strip()
        
        omath = ET.Element(f'{NS_M}oMath')
        
        # Process the entire latex
        self._parse_expr(omath, latex)
        
        return omath

    def _parse_expr(self, parent, expr):
        """Parse LaTeX expression and append to parent"""
        expr = expr.strip()
        if not expr:
            return
        
        # Handle parentheses with content: \func(...) or (...)
        # Check for \cos(, \sin(, etc. first
        func_paren_match = re.match(r'\\(cos|sin|tan|log|ln|exp)\s*\(([^)]+)\)', expr)
        if func_paren_match:
            func_name = func_paren_match.group(1)
            inner = func_paren_match.group(2)
            # Add function name as text
            parent.append(self._math_run(func_name))
            # Add delimiter for parentheses
            d = ET.SubElement(parent, f'{NS_M}d')
            dPr = ET.SubElement(d, f'{NS_M}dPr')
            begChr = ET.SubElement(dPr, f'{NS_M}begChr')
            begChr.set(f'{NS_M}val', '(')
            endChr = ET.SubElement(dPr, f'{NS_M}endChr')
            endChr.set(f'{NS_M}val', ')')
            dPr.append(self._ctrl_pr())
            e = ET.SubElement(d, f'{NS_M}e')
            self._parse_expr(e, inner)
            e.append(self._ctrl_pr())
            d.append(self._ctrl_pr())
            remaining = expr[func_paren_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle plain parentheses with nested fractions
        paren_match = re.match(r'\(([^)]+)\)', expr)
        if paren_match:
            inner = paren_match.group(1)
            # Check if inner contains complex content (frac, etc.)
            if '\\frac' in inner or '\\' in inner:
                d = ET.SubElement(parent, f'{NS_M}d')
                dPr = ET.SubElement(d, f'{NS_M}dPr')
                begChr = ET.SubElement(dPr, f'{NS_M}begChr')
                begChr.set(f'{NS_M}val', '(')
                endChr = ET.SubElement(dPr, f'{NS_M}endChr')
                endChr.set(f'{NS_M}val', ')')
                dPr.append(self._ctrl_pr())
                e = ET.SubElement(d, f'{NS_M}e')
                self._parse_expr(e, inner)
                e.append(self._ctrl_pr())
                d.append(self._ctrl_pr())
            else:
                # Simple parentheses, just add as text
                parent.append(self._math_run('('))
                self._parse_expr(parent, inner)
                parent.append(self._math_run(')'))
            remaining = expr[paren_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle matrix first (before other patterns)
        matrix_match = re.match(r'\\begin\{bmatrix\}(.*?)\\end\{bmatrix\}', expr, re.DOTALL)
        if matrix_match:
            content = matrix_match.group(1).strip()
            rows = [r.strip() for r in content.split('\\\\') if r.strip()]
            # Wrap matrix in delimiter for brackets
            d = ET.SubElement(parent, f'{NS_M}d')
            dPr = ET.SubElement(d, f'{NS_M}dPr')
            begChr = ET.SubElement(dPr, f'{NS_M}begChr')
            begChr.set(f'{NS_M}val', '[')
            endChr = ET.SubElement(dPr, f'{NS_M}endChr')
            endChr.set(f'{NS_M}val', ']')
            dPr.append(self._ctrl_pr())
            e = ET.SubElement(d, f'{NS_M}e')
            m = ET.SubElement(e, f'{NS_M}m')
            mPr = ET.SubElement(m, f'{NS_M}mPr')
            # Add m:mcs for column definitions like template
            mcs = ET.SubElement(mPr, f'{NS_M}mcs')
            mc = ET.SubElement(mcs, f'{NS_M}mc')
            mcPr = ET.SubElement(mc, f'{NS_M}mcPr')
            count = ET.SubElement(mcPr, f'{NS_M}count')
            count.set(f'{NS_M}val', '1')
            mcJc = ET.SubElement(mcPr, f'{NS_M}mcJc')
            mcJc.set(f'{NS_M}val', 'center')
            plcHide = ET.SubElement(mPr, f'{NS_M}plcHide')
            plcHide.set(f'{NS_M}val', '1')
            mPr.append(self._ctrl_pr())
            for row in rows:
                mr = ET.SubElement(m, f'{NS_M}mr')
                cells = [c.strip() for c in row.split('&')]
                for cell in cells:
                    e_cell = ET.SubElement(mr, f'{NS_M}e')
                    self._parse_expr(e_cell, cell)
                    e_cell.append(self._ctrl_pr())
            e.append(self._ctrl_pr())
            d.append(self._ctrl_pr())
            remaining = expr[matrix_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle cases environment - convert to single expression with braces
        cases_match = re.match(r'\\begin\{cases\}(.*?)\\end\{cases\}', expr, re.DOTALL)
        if cases_match:
            content = cases_match.group(1).strip()
            rows = [r.strip() for r in content.split('\\\\') if r.strip()]
            
            # Create a left brace delimiter containing all cases
            d = ET.SubElement(parent, f'{NS_M}d')
            dPr = ET.SubElement(d, f'{NS_M}dPr')
            begChr = ET.SubElement(dPr, f'{NS_M}begChr')
            begChr.set(f'{NS_M}val', '{')
            endChr = ET.SubElement(dPr, f'{NS_M}endChr')
            endChr.set(f'{NS_M}val', '')  # Empty end char
            dPr.append(self._ctrl_pr())
            
            e = ET.SubElement(d, f'{NS_M}e')
            
            # Add each row as separate expression with line break
            for i, row in enumerate(rows):
                if i > 0:
                    # Add line break character
                    e.append(self._math_run(' '))  # Space for separation
                if '&' in row:
                    # Only parse the left part (before &)
                    part = row.split('&')[0].strip()
                    self._parse_expr(e, part)
                else:
                    self._parse_expr(e, row)
            
            e.append(self._ctrl_pr())
            d.append(self._ctrl_pr())
            
            remaining = expr[cases_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle fractions (before subscripts to avoid conflicts with di_d)
        frac_match = re.match(r'\\frac\{([^}]+)\}\{([^}]+)\}', expr)
        if frac_match:
            f = ET.SubElement(parent, f'{NS_M}f')
            fPr = ET.SubElement(f, f'{NS_M}fPr')
            fPr.append(self._ctrl_pr())
            num = ET.SubElement(f, f'{NS_M}num')
            self._parse_expr(num, frac_match.group(1))
            num.append(self._ctrl_pr())
            den = ET.SubElement(f, f'{NS_M}den')
            self._parse_expr(den, frac_match.group(2))
            den.append(self._ctrl_pr())
            parent.append(self._ctrl_pr())
            remaining = expr[frac_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle combined subscript and superscript: x_d^* or x_{d}^{*}
        # Also supports Greek letters: \omega_e^*, i_\alpha^*, etc.
        combined_patterns = [
            r'\\([a-zA-Z]+)_\{([^}]+)\}\^\{([^}]+)\}',  # \omega_{e}^{*}
            r'\\([a-zA-Z]+)_([a-zA-Z0-9])\^([a-zA-Z0-9*])',  # \omega_e^*
            r'([a-zA-Z0-9])_\\([a-zA-Z]+)\^\{([^}]+)\}',  # i_\alpha^{*}
            r'([a-zA-Z0-9])_\\([a-zA-Z]+)\^([a-zA-Z0-9*])',  # i_\alpha^*
            r'([a-zA-Z0-9])_\{([^}]+)\}\^\{([^}]+)\}',  # u_{d}^{*}
            r'([a-zA-Z0-9])_([a-zA-Z0-9])\^([a-zA-Z0-9*])',  # u_d^*
        ]
        combined_match = None
        for pattern in combined_patterns:
            combined_match = re.match(pattern, expr)
            if combined_match:
                break
        if combined_match:
            ssubsup = ET.SubElement(parent, f'{NS_M}sSubSup')
            ssubsupPr = ET.SubElement(ssubsup, f'{NS_M}sSubSupPr')
            ssubsupPr.append(self._ctrl_pr())
            e = ET.SubElement(ssubsup, f'{NS_M}e')
            # Handle base character (could be Greek letter or regular)
            base = combined_match.group(1)
            if base in self.greek_map:
                e.append(self._math_run_with_hint(self.greek_map[base], 'eastAsia'))
            else:
                e.append(self._math_run(base))
            e.append(self._ctrl_pr())
            sub = ET.SubElement(ssubsup, f'{NS_M}sub')
            # Handle subscript (could be Greek letter or regular)
            sub_val = combined_match.group(2)
            if sub_val in self.greek_map:
                sub.append(self._math_run_with_hint(self.greek_map[sub_val], 'eastAsia'))
            else:
                sub.append(self._math_run(sub_val))
            sub.append(self._ctrl_pr())
            sup = ET.SubElement(ssubsup, f'{NS_M}sup')
            # Handle superscript (could be Greek letter or regular)
            sup_val = combined_match.group(3)
            if sup_val in self.greek_map:
                sup.append(self._math_run_with_hint(self.greek_map[sup_val], 'eastAsia'))
            else:
                sup.append(self._math_run(sup_val))
            sup.append(self._ctrl_pr())
            ssubsup.append(self._ctrl_pr())
            remaining = expr[combined_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle subscripts: x_d, x_{abc}, i_\alpha
        sub_patterns = [
            r'\\([a-zA-Z]+)_\{([^}]+)\}',  # \omega_{e}
            r'\\([a-zA-Z]+)_([a-zA-Z0-9])',  # \omega_e
            r'([a-zA-Z0-9])_\\([a-zA-Z]+)',  # i_\alpha
            r'([a-zA-Z0-9])_\{([^}]+)\}',  # u_{d}
            r'([a-zA-Z0-9])_([a-zA-Z0-9])',  # u_d
        ]
        sub_match = None
        for pattern in sub_patterns:
            sub_match = re.match(pattern, expr)
            if sub_match:
                break
        if sub_match:
            ssub = ET.SubElement(parent, f'{NS_M}sSub')
            ssubPr = ET.SubElement(ssub, f'{NS_M}sSubPr')
            ssubPr.append(self._ctrl_pr())
            e = ET.SubElement(ssub, f'{NS_M}e')
            base = sub_match.group(1)
            if base in self.greek_map:
                e.append(self._math_run_with_hint(self.greek_map[base], 'eastAsia'))
            else:
                e.append(self._math_run(base))
            e.append(self._ctrl_pr())
            sub = ET.SubElement(ssub, f'{NS_M}sub')
            sub_val = sub_match.group(2)
            if sub_val in self.greek_map:
                sub.append(self._math_run_with_hint(self.greek_map[sub_val], 'eastAsia'))
            else:
                sub.append(self._math_run(sub_val))
            sub.append(self._ctrl_pr())
            ssub.append(self._ctrl_pr())
            remaining = expr[sub_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle superscripts: x^2, x^{n}, \omega^*
        sup_patterns = [
            r'\\([a-zA-Z]+)\^\{([^}]+)\}',  # \omega^{*}
            r'\\([a-zA-Z]+)\^([a-zA-Z0-9*])',  # \omega^*
            r'([a-zA-Z0-9])\^\{([^}]+)\}',  # x^{n}
            r'([a-zA-Z0-9])\^([a-zA-Z0-9*])',  # x^2
        ]
        sup_match = None
        for pattern in sup_patterns:
            sup_match = re.match(pattern, expr)
            if sup_match:
                break
        if sup_match:
            ssup = ET.SubElement(parent, f'{NS_M}sSup')
            ssupPr = ET.SubElement(ssup, f'{NS_M}sSupPr')
            ssupPr.append(self._ctrl_pr())
            e = ET.SubElement(ssup, f'{NS_M}e')
            base = sup_match.group(1)
            if base in self.greek_map:
                e.append(self._math_run_with_hint(self.greek_map[base], 'eastAsia'))
            else:
                e.append(self._math_run(base))
            e.append(self._ctrl_pr())
            sup = ET.SubElement(ssup, f'{NS_M}sup')
            sup_val = sup_match.group(2)
            if sup_val in self.greek_map:
                sup.append(self._math_run_with_hint(self.greek_map[sup_val], 'eastAsia'))
            else:
                sup.append(self._math_run(sup_val))
            sup.append(self._ctrl_pr())
            ssup.append(self._ctrl_pr())
            remaining = expr[sup_match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Handle Greek letters
        greek_match = re.match(r'\\([a-zA-Z]+)', expr)
        if greek_match:
            name = greek_match.group(1)
            if name in self.greek_map:
                parent.append(self._math_run_with_hint(self.greek_map[name], 'eastAsia'))
                remaining = expr[greek_match.end():]
                if remaining:
                    self._parse_expr(parent, remaining)
                return
        
        # Handle operators and symbols
        if expr[0] in '=+-*/()[]{},;: ':
            parent.append(self._math_run(expr[0]))
            if len(expr) > 1:
                self._parse_expr(parent, expr[1:])
            return
        
        # Handle regular text (variables, numbers)
        match = re.match(r'([a-zA-Z0-9\.]+)', expr)
        if match:
            parent.append(self._math_run(match.group(1)))
            remaining = expr[match.end():]
            if remaining:
                self._parse_expr(parent, remaining)
            return
        
        # Skip unknown characters
        if len(expr) > 1:
            self._parse_expr(parent, expr[1:])


class MarkdownToDocxConverter:
    def __init__(self, template_path):
        self.template_path = template_path
        self.latex_converter = LatexToOmmlConverter()
        self.para_id_counter = 0x10000000

    def _generate_para_id(self):
        pid = hex(self.para_id_counter)[2:].upper().zfill(8)
        self.para_id_counter += 1
        return pid

    def convert(self, markdown_content, output_path):
        temp_dir = '/tmp/docx_temp_' + hashlib.md5(output_path.encode()).hexdigest()[:8]
        if os.path.exists(temp_dir):
            shutil.rmtree(temp_dir)
        with zipfile.ZipFile(self.template_path, 'r') as zip_ref:
            zip_ref.extractall(temp_dir)
        document_xml = self._generate_document_xml(markdown_content)
        doc_path = os.path.join(temp_dir, 'word', 'document.xml')
        with open(doc_path, 'w', encoding='utf-8') as f:
            f.write(document_xml)
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as zipf:
            for root, dirs, files in os.walk(temp_dir):
                for file in files:
                    file_path = os.path.join(root, file)
                    arcname = os.path.relpath(file_path, temp_dir)
                    zipf.write(file_path, arcname)
        shutil.rmtree(temp_dir)
        print(f"Converted to: {output_path}")

    def _generate_document_xml(self, markdown):
        lines = markdown.split('\n')
        root = ET.Element(f'{NS_W}document')
        root.set('xmlns:wpc', 'http://schemas.microsoft.com/office/word/2010/wordprocessingCanvas')
        root.set('xmlns:mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        root.set('xmlns:o', 'urn:schemas-microsoft-com:office:office')
        root.set('xmlns:v', 'urn:schemas-microsoft-com:vml')
        root.set('xmlns:wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')
        root.set('xmlns:w10', 'urn:schemas-microsoft-com:office:word')
        root.set('xmlns:w15', 'http://schemas.microsoft.com/office/word/2012/wordml')
        root.set('xmlns:wpg', 'http://schemas.microsoft.com/office/word/2010/wordprocessingGroup')
        root.set('xmlns:wpi', 'http://schemas.microsoft.com/office/word/2010/wordprocessingInk')
        root.set('xmlns:wne', 'http://schemas.microsoft.com/office/word/2006/wordml')
        root.set('xmlns:wps', 'http://schemas.microsoft.com/office/word/2010/wordprocessingShape')
        root.set('xmlns:wpsCustomData', 'http://www.wps.cn/officeDocument/2013/wpsCustomData')
        root.set('mc:Ignorable', 'w14 w15 wp14')
        body = ET.SubElement(root, f'{NS_W}body')
        body.append(self._create_empty_paragraph())
        i = 0
        while i < len(lines):
            line = lines[i].rstrip()
            stripped = line.strip()
            if not stripped:
                i += 1
                continue
            if stripped.startswith('# '):
                body.append(self._create_heading1(stripped[2:].strip()))
                i += 1
            elif stripped.startswith('## '):
                body.append(self._create_heading2(stripped[3:].strip()))
                i += 1
            elif stripped.startswith('### '):
                body.append(self._create_heading3(stripped[4:].strip()))
                i += 1
            elif stripped.startswith('|'):
                table_lines = []
                while i < len(lines) and lines[i].strip().startswith('|'):
                    table_lines.append(lines[i])
                    i += 1
                if len(table_lines) >= 2:
                    tbl = self._create_table(table_lines)
                    if tbl:
                        body.append(tbl)
            elif stripped.startswith('```'):
                code_lines = []
                i += 1
                while i < len(lines) and not lines[i].strip().startswith('```'):
                    code_lines.append(lines[i])
                    i += 1
                i += 1
                body.append(self._create_code_paragraph('\n'.join(code_lines)))
            elif stripped.startswith('$$'):
                formula_lines = [stripped]
                if not stripped.endswith('$$') or stripped == '$$':
                    i += 1
                    while i < len(lines) and not lines[i].strip().endswith('$$'):
                        formula_lines.append(lines[i])
                        i += 1
                    if i < len(lines):
                        formula_lines.append(lines[i].strip())
                        i += 1
                else:
                    i += 1
                body.append(self._create_formula_paragraph('\n'.join(formula_lines)))
            else:
                para_lines = [line]
                i += 1
                while i < len(lines):
                    next_line = lines[i].rstrip()
                    if not next_line.strip():
                        break
                    if next_line.strip().startswith(('#', '|', '$$', '```')):
                        break
                    para_lines.append(next_line)
                    i += 1
                body.append(self._create_body_paragraph('\n'.join(para_lines)))
        body.append(self._create_section_properties())
        xml_str = ET.tostring(root, encoding='unicode')
        # Fix namespace prefixes
        xml_str = xml_str.replace('xmlns:ns0=', 'xmlns:w=')
        xml_str = xml_str.replace('xmlns:ns1=', 'xmlns:w14=')
        xml_str = xml_str.replace('xmlns:ns2=', 'xmlns:m=')
        xml_str = xml_str.replace('xmlns:ns3=', 'xmlns:r=')
        xml_str = xml_str.replace('<ns0:', '<w:')
        xml_str = xml_str.replace('</ns0:', '</w:')
        xml_str = xml_str.replace(' ns0:', ' w:')
        xml_str = xml_str.replace('<ns1:', '<w14:')
        xml_str = xml_str.replace('</ns1:', '</w14:')
        xml_str = xml_str.replace(' ns1:', ' w14:')
        xml_str = xml_str.replace('<ns2:', '<m:')
        xml_str = xml_str.replace('</ns2:', '</m:')
        xml_str = xml_str.replace(' ns2:', ' m:')
        xml_str = xml_str.replace('<ns3:', '<r:')
        xml_str = xml_str.replace('</ns3:', '</r:')
        xml_str = xml_str.replace(' ns3:', ' r:')
        # Fix self-closing tags to match Word format
        xml_str = xml_str.replace(' />', '/>')
        return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' + xml_str

    def _create_empty_paragraph(self):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        return p

    def _create_heading1(self, text):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
        pStyle.set(f'{NS_W}val', STYLES['heading1'])
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

    def _create_heading2(self, text):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
        pStyle.set(f'{NS_W}val', STYLES['heading2'])
        numPr = ET.SubElement(pPr, f'{NS_W}numPr')
        ilvl = ET.SubElement(numPr, f'{NS_W}ilvl')
        ilvl.set(f'{NS_W}val', '1')
        numId = ET.SubElement(numPr, f'{NS_W}numId')
        numId.set(f'{NS_W}val', '0')
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'default')
        lang = ET.SubElement(rPr, f'{NS_W}lang')
        lang.set(f'{NS_W}val', 'en-US')
        lang.set(f'{NS_W}eastAsia', 'zh-CN')
        r = ET.SubElement(p, f'{NS_W}r')
        rPr2 = ET.SubElement(r, f'{NS_W}rPr')
        rFonts2 = ET.SubElement(rPr2, f'{NS_W}rFonts')
        rFonts2.set(f'{NS_W}hint', 'eastAsia')
        t = ET.SubElement(r, f'{NS_W}t')
        t.text = text
        return p

    def _create_heading3(self, text):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
        pStyle.set(f'{NS_W}val', STYLES['heading3'])
        numPr = ET.SubElement(pPr, f'{NS_W}numPr')
        ilvl = ET.SubElement(numPr, f'{NS_W}ilvl')
        ilvl.set(f'{NS_W}val', '1')
        numId = ET.SubElement(numPr, f'{NS_W}numId')
        numId.set(f'{NS_W}val', '0')
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'default')
        lang = ET.SubElement(rPr, f'{NS_W}lang')
        lang.set(f'{NS_W}val', 'en-US')
        lang.set(f'{NS_W}eastAsia', 'zh-CN')
        r = ET.SubElement(p, f'{NS_W}r')
        rPr2 = ET.SubElement(r, f'{NS_W}rPr')
        rFonts2 = ET.SubElement(rPr2, f'{NS_W}rFonts')
        rFonts2.set(f'{NS_W}hint', 'eastAsia')
        t = ET.SubElement(r, f'{NS_W}t')
        t.text = text
        return p

    def _create_body_paragraph(self, text):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
        pStyle.set(f'{NS_W}val', STYLES['body'])
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'eastAsia')
        lang = ET.SubElement(rPr, f'{NS_W}lang')
        lang.set(f'{NS_W}eastAsia', 'zh-CN')
        # Split inline formulas
        parts = re.split(r'(\$[^$]+\$)', text)
        for part in parts:
            if part.startswith('$') and part.endswith('$'):
                latex = part[1:-1]
                omml = self.latex_converter.convert(latex)
                if omml is not None:
                    p.append(omml)
            else:
                bold_parts = re.split(r'(\*\*[^*]+\*\*)', part)
                for bp in bold_parts:
                    if bp.startswith('**') and bp.endswith('**'):
                        r = ET.SubElement(p, f'{NS_W}r')
                        rPr2 = ET.SubElement(r, f'{NS_W}rPr')
                        rFonts2 = ET.SubElement(rPr2, f'{NS_W}rFonts')
                        rFonts2.set(f'{NS_W}hint', 'eastAsia')
                        b = ET.SubElement(rPr2, f'{NS_W}b')
                        lang2 = ET.SubElement(rPr2, f'{NS_W}lang')
                        lang2.set(f'{NS_W}val', 'en-US')
                        lang2.set(f'{NS_W}eastAsia', 'zh-CN')
                        t = ET.SubElement(r, f'{NS_W}t')
                        t.text = bp[2:-2]
                    elif bp.strip():
                        r = ET.SubElement(p, f'{NS_W}r')
                        rPr2 = ET.SubElement(r, f'{NS_W}rPr')
                        rFonts2 = ET.SubElement(rPr2, f'{NS_W}rFonts')
                        rFonts2.set(f'{NS_W}hint', 'eastAsia')
                        t = ET.SubElement(r, f'{NS_W}t')
                        t.text = bp
        return p

    def _create_formula_paragraph(self, formula):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
        pStyle.set(f'{NS_W}val', STYLES['formula'])
        bidi = ET.SubElement(pPr, f'{NS_W}bidi')
        bidi.set(f'{NS_W}val', '0')
        rPr = ET.SubElement(pPr, f'{NS_W}rPr')
        rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
        rFonts.set(f'{NS_W}hint', 'default')
        rFonts.set(f'{NS_W}eastAsia', '宋体')
        lang = ET.SubElement(rPr, f'{NS_W}lang')
        lang.set(f'{NS_W}val', 'en-US')
        lang.set(f'{NS_W}eastAsia', 'zh-CN')
        # Tab before formula - match template structure
        r_tab = ET.SubElement(p, f'{NS_W}r')
        rPr_tab = ET.SubElement(r_tab, f'{NS_W}rPr')
        rFonts_tab = ET.SubElement(rPr_tab, f'{NS_W}rFonts')
        rFonts_tab.set(f'{NS_W}hint', 'eastAsia')
        rFonts_tab.set(f'{NS_W}hAnsi', 'Cambria Math')
        b_elem = ET.SubElement(rPr_tab, f'{NS_W}b')
        b_elem.set(f'{NS_W}val', '0')
        i_elem = ET.SubElement(rPr_tab, f'{NS_W}i')
        i_elem.set(f'{NS_W}val', '0')
        lang_tab = ET.SubElement(rPr_tab, f'{NS_W}lang')
        lang_tab.set(f'{NS_W}val', 'en-US')
        lang_tab.set(f'{NS_W}eastAsia', 'zh-CN')
        tab_elem = ET.SubElement(r_tab, f'{NS_W}tab')
        # Formula
        omml = self.latex_converter.convert(formula)
        if omml is not None:
            p.append(omml)
        # Tab after formula - match template structure
        r_tab2 = ET.SubElement(p, f'{NS_W}r')
        rPr_tab2 = ET.SubElement(r_tab2, f'{NS_W}rPr')
        rFonts_tab2 = ET.SubElement(rPr_tab2, f'{NS_W}rFonts')
        rFonts_tab2.set(f'{NS_W}hint', 'eastAsia')
        rFonts_tab2.set(f'{NS_W}hAnsi', 'Cambria Math')
        i_elem2 = ET.SubElement(rPr_tab2, f'{NS_W}i')
        i_elem2.set(f'{NS_W}val', '0')
        lang_tab2 = ET.SubElement(rPr_tab2, f'{NS_W}lang')
        lang_tab2.set(f'{NS_W}val', 'en-US')
        lang_tab2.set(f'{NS_W}eastAsia', 'zh-CN')
        tab_elem2 = ET.SubElement(r_tab2, f'{NS_W}tab')
        return p

    def _create_code_paragraph(self, code_text):
        p = ET.Element(f'{NS_W}p')
        p.set(f'{NS_W14}paraId', self._generate_para_id())
        pPr = ET.SubElement(p, f'{NS_W}pPr')
        pStyle = ET.SubElement(pPr, f'{NS_W}pStyle')
        pStyle.set(f'{NS_W}val', '58')
        lines = code_text.split('\n')
        for i, line in enumerate(lines):
            if i > 0:
                br = ET.SubElement(p, f'{NS_W}r')
                br_elem = ET.SubElement(br, f'{NS_W}br')
            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
            rFonts.set(f'{NS_W}ascii', 'Courier New')
            rFonts.set(f'{NS_W}hAnsi', 'Courier New')
            t = ET.SubElement(r, f'{NS_W}t')
            if line.startswith(' ') or line.endswith(' '):
                t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = line if line else ''
        return p

    def _create_table(self, table_lines):
        if len(table_lines) < 2:
            return None
        header_line = table_lines[0]
        separator_line = table_lines[1]
        data_lines = table_lines[2:]
        headers = [h.strip() for h in header_line.split('|')[1:-1]]
        rows = []
        for line in data_lines:
            if line.strip():
                cells = [c.strip() for c in line.split('|')[1:-1]]
                if cells:
                    rows.append(cells)
        tbl = ET.Element(f'{NS_W}tbl')
        tblPr = ET.SubElement(tbl, f'{NS_W}tblPr')
        tblStyle = ET.SubElement(tblPr, f'{NS_W}tblStyle')
        tblStyle.set(f'{NS_W}val', '46')
        tblW = ET.SubElement(tblPr, f'{NS_W}tblW')
        tblW.set(f'{NS_W}w', '5000')
        tblW.set(f'{NS_W}type', 'pct')
        jc = ET.SubElement(tblPr, f'{NS_W}jc')
        jc.set(f'{NS_W}val', 'center')
        tblLayout = ET.SubElement(tblPr, f'{NS_W}tblLayout')
        tblLayout.set(f'{NS_W}type', 'fixed')
        tblGrid = ET.SubElement(tbl, f'{NS_W}tblGrid')
        for _ in headers:
            gridCol = ET.SubElement(tblGrid, f'{NS_W}gridCol')
            gridCol.set(f'{NS_W}w', str(int(9000 / len(headers))))
        tr_header = ET.SubElement(tbl, f'{NS_W}tr')
        for header in headers:
            tc = ET.SubElement(tr_header, f'{NS_W}tc')
            tcPr = ET.SubElement(tc, f'{NS_W}tcPr')
            tcW = ET.SubElement(tcPr, f'{NS_W}tcW')
            tcW.set(f'{NS_W}w', str(int(9000 / len(headers))))
            tcW.set(f'{NS_W}type', 'dxa')
            p = ET.SubElement(tc, f'{NS_W}p')
            p.set(f'{NS_W14}paraId', self._generate_para_id())
            pPr = ET.SubElement(p, f'{NS_W}pPr')
            jc_elem = ET.SubElement(pPr, f'{NS_W}jc')
            jc_elem.set(f'{NS_W}val', 'center')
            r = ET.SubElement(p, f'{NS_W}r')
            rPr = ET.SubElement(r, f'{NS_W}rPr')
            rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
            rFonts.set(f'{NS_W}hint', 'eastAsia')
            b = ET.SubElement(rPr, f'{NS_W}b')
            t = ET.SubElement(r, f'{NS_W}t')
            t.text = header
        for row_data in rows:
            tr = ET.SubElement(tbl, f'{NS_W}tr')
            for cell_text in row_data:
                tc = ET.SubElement(tr, f'{NS_W}tc')
                tcPr = ET.SubElement(tc, f'{NS_W}tcPr')
                tcW = ET.SubElement(tcPr, f'{NS_W}tcW')
                tcW.set(f'{NS_W}w', str(int(9000 / len(headers))))
                tcW.set(f'{NS_W}type', 'dxa')
                p = ET.SubElement(tc, f'{NS_W}p')
                p.set(f'{NS_W14}paraId', self._generate_para_id())
                pPr = ET.SubElement(p, f'{NS_W}pPr')
                jc_elem = ET.SubElement(pPr, f'{NS_W}jc')
                jc_elem.set(f'{NS_W}val', 'center')
                r = ET.SubElement(p, f'{NS_W}r')
                rPr = ET.SubElement(r, f'{NS_W}rPr')
                rFonts = ET.SubElement(rPr, f'{NS_W}rFonts')
                rFonts.set(f'{NS_W}hint', 'eastAsia')
                t = ET.SubElement(r, f'{NS_W}t')
                t.text = cell_text
        return tbl

    def _create_section_properties(self):
        sectPr = ET.Element(f'{NS_W}sectPr')
        header_ref = ET.SubElement(sectPr, f'{NS_W}headerReference')
        header_ref.set(f'{NS_R}id', 'rId3')
        header_ref.set(f'{NS_W}type', 'default')
        footer_ref = ET.SubElement(sectPr, f'{NS_W}footerReference')
        footer_ref.set(f'{NS_R}id', 'rId4')
        footer_ref.set(f'{NS_W}type', 'default')
        pgSz = ET.SubElement(sectPr, f'{NS_W}pgSz')
        pgSz.set(f'{NS_W}w', '11906')
        pgSz.set(f'{NS_W}h', '16838')
        pgMar = ET.SubElement(sectPr, f'{NS_W}pgMar')
        pgMar.set(f'{NS_W}top', '1474')
        pgMar.set(f'{NS_W}right', '1531')
        pgMar.set(f'{NS_W}bottom', '1474')
        pgMar.set(f'{NS_W}left', '1531')
        pgMar.set(f'{NS_W}header', '851')
        pgMar.set(f'{NS_W}footer', '992')
        pgMar.set(f'{NS_W}gutter', '0')
        pgBorders = ET.SubElement(sectPr, f'{NS_W}pgBorders')
        for border_pos in ['top', 'left', 'bottom', 'right']:
            border = ET.SubElement(pgBorders, f'{NS_W}{border_pos}')
            border.set(f'{NS_W}val', 'none')
            border.set(f'{NS_W}sz', '0')
            border.set(f'{NS_W}space', '0')
        pgNumType = ET.SubElement(sectPr, f'{NS_W}pgNumType')
        pgNumType.set(f'{NS_W}start', '1')
        cols = ET.SubElement(sectPr, f'{NS_W}cols')
        cols.set(f'{NS_W}space', '720')
        cols.set(f'{NS_W}num', '1')
        docGrid = ET.SubElement(sectPr, f'{NS_W}docGrid')
        docGrid.set(f'{NS_W}type', 'lines')
        docGrid.set(f'{NS_W}linePitch', '312')
        docGrid.set(f'{NS_W}charSpace', '0')
        return sectPr


def main():
    template_path = 'test_cases/Southwest Jiaotong University Thesis Template/正文.docx'
    markdown_path = 'test_cases/正文内容.md'
    output_path = 'test_cases/output.docx'
    with open(markdown_path, 'r', encoding='utf-8') as f:
        markdown_content = f.read()
    converter = MarkdownToDocxConverter(template_path)
    converter.convert(markdown_content, output_path)
    print(f"Conversion complete: {output_path}")


if __name__ == '__main__':
    main()
