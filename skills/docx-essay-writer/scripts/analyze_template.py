#!/usr/bin/env python3
"""
DOCX Template Structure Analyzer

Parses any .docx template to extract:
- Heading hierarchy (outline / writing structure)
- Style ID mapping (style name → styleId)
- Page layout (margins, page size, header/footer refs)
- Special elements inventory (tables, formulas, images, references)

Output: JSON that guides AI content generation and format-preserving conversion.
"""

import argparse
import json
import os
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from typing import Any, Dict, List, Optional

NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'pic': 'http://schemas.openxmlformats.org/drawingml/2006/picture',
    'v': 'urn:schemas-microsoft-com:vml',
    'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
}


class TemplateAnalyzer:
    """Analyze a DOCX template and extract its structure."""

    def __init__(self, docx_path: str):
        if not os.path.isfile(docx_path):
            raise FileNotFoundError(f"Template not found: {docx_path}")
        self.docx_path = docx_path
        self.zip = zipfile.ZipFile(docx_path, 'r')
        self._doc_tree: Optional[ET.Element] = None
        self._styles_tree: Optional[ET.Element] = None

    def close(self):
        self.zip.close()

    def __enter__(self):
        return self

    def __exit__(self, *args):
        self.close()

    # ------------------------------------------------------------------
    # XML helpers
    # ------------------------------------------------------------------

    def _read_xml(self, inner_path: str) -> Optional[ET.Element]:
        try:
            data = self.zip.read(inner_path)
            return ET.fromstring(data)
        except (KeyError, ET.ParseError):
            return None

    @property
    def doc_root(self) -> Optional[ET.Element]:
        if self._doc_tree is None:
            self._doc_tree = self._read_xml('word/document.xml')
        return self._doc_tree

    @property
    def styles_root(self) -> Optional[ET.Element]:
        if self._styles_tree is None:
            self._styles_tree = self._read_xml('word/styles.xml')
        return self._styles_tree

    # ------------------------------------------------------------------
    # Style mapping
    # ------------------------------------------------------------------

    def extract_style_map(self) -> Dict[str, Dict[str, str]]:
        """Return {styleName: {id, type, name}} for every style in styles.xml."""
        result: Dict[str, Dict[str, str]] = {}
        root = self.styles_root
        if root is None:
            return result

        for style_el in root.findall('.//w:style', NS):
            style_id = style_el.get(f'{{{NS["w"]}}}styleId') or style_el.get('w:styleId', '')
            style_type = style_el.get(f'{{{NS["w"]}}}type') or style_el.get('w:type', '')

            name_el = style_el.find('w:name', NS)
            name_val = ''
            if name_el is not None:
                name_val = name_el.get(f'{{{NS["w"]}}}val', '')

            if style_id:
                result[name_val or style_id] = {
                    'id': style_id,
                    'type': style_type,
                    'name': name_val,
                }
        return result

    def _heading_style_ids(self, style_map: Dict) -> Dict[int, str]:
        """Map heading level (1-9) → styleId."""
        heading_ids: Dict[int, str] = {}
        for name, info in style_map.items():
            lower = name.lower()
            for lvl in range(1, 10):
                if lower in (f'heading {lvl}', f'heading{lvl}',
                             f'标题 {lvl}', f'标题{lvl}'):
                    heading_ids[lvl] = info['id']
        return heading_ids

    # ------------------------------------------------------------------
    # Outline extraction
    # ------------------------------------------------------------------

    def extract_outline(self) -> List[Dict[str, Any]]:
        """Walk document body and return heading hierarchy with section info."""
        root = self.doc_root
        if root is None:
            return []

        body = root.find('w:body', NS)
        if body is None:
            return []

        style_map = self.extract_style_map()
        heading_ids = self._heading_style_ids(style_map)
        id_to_level = {v: k for k, v in heading_ids.items()}

        outline: List[Dict[str, Any]] = []
        current_section_elements: List[str] = []
        para_index = 0

        for child in body:
            tag = child.tag.split('}')[-1] if '}' in child.tag else child.tag

            if tag == 'tbl':
                current_section_elements.append('table')
                para_index += 1
                continue

            if tag != 'p':
                para_index += 1
                continue

            pPr = child.find('w:pPr', NS)
            style_id = None
            if pPr is not None:
                pStyle = pPr.find('w:pStyle', NS)
                if pStyle is not None:
                    style_id = pStyle.get(f'{{{NS["w"]}}}val', '')

            has_formula = child.find('.//m:oMath', NS) is not None
            has_image = (child.find('.//wp:inline', NS) is not None or
                         child.find('.//wp:anchor', NS) is not None or
                         child.find('.//v:imagedata', NS) is not None or
                         child.find('.//w:drawing', NS) is not None or
                         child.find('.//w:pict', NS) is not None)

            if has_formula:
                current_section_elements.append('formula')
            if has_image:
                current_section_elements.append('image')

            text = self._paragraph_text(child)

            if re.match(r'^\[\d+\]', text.strip()):
                current_section_elements.append('reference')

            if style_id and style_id in id_to_level:
                level = id_to_level[style_id]
                if outline:
                    outline[-1]['contains'] = list(set(current_section_elements))
                    current_section_elements = []
                outline.append({
                    'level': level,
                    'title': text.strip(),
                    'style_id': style_id,
                    'para_index': para_index,
                    'contains': [],
                })

            para_index += 1

        if outline:
            outline[-1]['contains'] = list(set(current_section_elements))

        return outline

    @staticmethod
    def _paragraph_text(p_el: ET.Element) -> str:
        """Concatenate all w:t text in a paragraph."""
        parts = []
        for t in p_el.iter(f'{{{NS["w"]}}}t'):
            if t.text:
                parts.append(t.text)
        return ''.join(parts)

    # ------------------------------------------------------------------
    # Page layout
    # ------------------------------------------------------------------

    def extract_page_layout(self) -> Dict[str, Any]:
        """Extract page margins, size, header/footer refs from sectPr."""
        root = self.doc_root
        if root is None:
            return {}

        body = root.find('w:body', NS)
        if body is None:
            return {}

        sect_pr = body.find('w:sectPr', NS)
        if sect_pr is None:
            for child in reversed(list(body)):
                sect_pr = child.find('w:sectPr', NS)
                if sect_pr is not None:
                    break

        if sect_pr is None:
            return {}

        layout: Dict[str, Any] = {}

        pgSz = sect_pr.find('w:pgSz', NS)
        if pgSz is not None:
            layout['page_size'] = {
                'width': pgSz.get(f'{{{NS["w"]}}}w', ''),
                'height': pgSz.get(f'{{{NS["w"]}}}h', ''),
            }
            orient = pgSz.get(f'{{{NS["w"]}}}orient', '')
            if orient:
                layout['page_size']['orient'] = orient

        pgMar = sect_pr.find('w:pgMar', NS)
        if pgMar is not None:
            layout['margins'] = {}
            for attr in ('top', 'right', 'bottom', 'left', 'header', 'footer', 'gutter'):
                val = pgMar.get(f'{{{NS["w"]}}}' + attr, '')
                if val:
                    layout['margins'][attr] = val

        layout['header_refs'] = []
        for href in sect_pr.findall('w:headerReference', NS):
            layout['header_refs'].append({
                'type': href.get(f'{{{NS["w"]}}}type', ''),
                'rId': href.get(f'{{{NS["r"]}}}id', ''),
            })

        layout['footer_refs'] = []
        for fref in sect_pr.findall('w:footerReference', NS):
            layout['footer_refs'].append({
                'type': fref.get(f'{{{NS["w"]}}}type', ''),
                'rId': fref.get(f'{{{NS["r"]}}}id', ''),
            })

        return layout

    # ------------------------------------------------------------------
    # Special element inventory
    # ------------------------------------------------------------------

    def extract_special_elements(self) -> Dict[str, int]:
        """Count tables, formulas, images, and reference-like paragraphs."""
        root = self.doc_root
        if root is None:
            return {}

        body = root.find('w:body', NS)
        if body is None:
            return {}

        counts = {
            'tables': 0,
            'formulas': 0,
            'images': 0,
            'references': 0,
        }

        for tbl in body.iter(f'{{{NS["w"]}}}tbl'):
            counts['tables'] += 1

        for omath in body.iter(f'{{{NS["m"]}}}oMath'):
            counts['formulas'] += 1

        for drawing in body.iter(f'{{{NS["w"]}}}drawing'):
            counts['images'] += 1
        for pict in body.iter(f'{{{NS["w"]}}}pict'):
            counts['images'] += 1

        for p in body.findall('w:p', NS):
            text = self._paragraph_text(p).strip()
            if re.match(r'^\[\d+\]', text):
                counts['references'] += 1

        return counts

    # ------------------------------------------------------------------
    # File inventory
    # ------------------------------------------------------------------

    def list_template_files(self) -> List[str]:
        """List all files inside the DOCX ZIP archive."""
        return self.zip.namelist()

    # ------------------------------------------------------------------
    # Full analysis
    # ------------------------------------------------------------------

    def analyze(self) -> Dict[str, Any]:
        """Run full analysis and return combined JSON-serializable dict."""
        style_map = self.extract_style_map()
        heading_ids = self._heading_style_ids(style_map)

        return {
            'template_path': os.path.abspath(self.docx_path),
            'outline': self.extract_outline(),
            'style_map': style_map,
            'heading_style_ids': {str(k): v for k, v in heading_ids.items()},
            'page_layout': self.extract_page_layout(),
            'special_elements': self.extract_special_elements(),
            'files': self.list_template_files(),
        }


def main():
    parser = argparse.ArgumentParser(
        description='Analyze a DOCX template and output its structure as JSON'
    )
    parser.add_argument('template', help='Path to the .docx template file')
    parser.add_argument('--output', '-o', help='Output JSON file (default: stdout)')
    parser.add_argument('--pretty', action='store_true', default=True,
                        help='Pretty-print JSON (default: True)')
    args = parser.parse_args()

    with TemplateAnalyzer(args.template) as analyzer:
        result = analyzer.analyze()

    json_str = json.dumps(result, ensure_ascii=False,
                          indent=2 if args.pretty else None)

    if args.output:
        with open(args.output, 'w', encoding='utf-8') as f:
            f.write(json_str + '\n')
        print(f"Analysis saved to: {args.output}", file=sys.stderr)
    else:
        print(json_str)


if __name__ == '__main__':
    main()
