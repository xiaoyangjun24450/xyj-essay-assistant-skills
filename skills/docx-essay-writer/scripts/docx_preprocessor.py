#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
docx预处理器：从docx生成unzipped、chunks和format_registry.json
输入：docx文件
输出：
  - unzipped/: docx解压后的原始结构
  - chunks/: 每个chunk文件包含带[PARA_ID]标记的段落文字（带格式类别ID标签）
  - origin_chunks/: chunks的备份
  - format_registry.json: 格式类别注册表（类别ID -> 规范化rPr XML）

新格式标记规则（格式类别注册 + ID 标签架构）：
  - 遍历文档所有run，按主要格式属性归纳为格式类别
  - 每个类别分配一个短ID（如F0、F1、F2）
  - 全局出现最多的类别作为基准（无需标签）
  - 其他run用 ‹ID:可读提示›...‹/› 标记
"""

import hashlib
import json
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
from typing import List, Dict, Any, Tuple, Optional, Set
from dataclasses import dataclass, asdict


@dataclass
class RunFormat:
    """Run的格式属性（仅主要属性，用于格式分类）"""
    font_ascii: str = ""          # 西文字体
    font_east_asia: str = ""      # 中文字体
    size: int = 0                 # 字号（半磅值，如24表示12pt）
    bold: bool = False            # 加粗
    italic: bool = False          # 斜体
    highlight: str = ""           # 高亮颜色
    color: str = ""               # 字体颜色
    underline: str = ""           # 下划线样式
    strike: bool = False          # 单删除线
    dstrike: bool = False         # 双删除线
    vert_align: str = ""          # 上下标 (superscript/subscript)
    small_caps: bool = False      # 小型大写字母
    
    def to_category_key(self) -> Tuple:
        """生成格式类别键（用于分类）"""
        # 使用有效字体（优先中文字体）
        effective_font = self.font_east_asia or self.font_ascii or "宋体"
        return (
            effective_font,
            self.size,
            self.bold,
            self.italic,
            self.highlight,
            self.color,
            self.underline,
            self.strike or self.dstrike,  # 删除线统一处理
            self.vert_align,
            self.small_caps,
        )
    
    def get_label(self) -> str:
        """生成可读标签"""
        parts = []
        font = self.font_east_asia or self.font_ascii or "宋体"
        if font:
            # 简化字体名
            font_short = font.replace('Times New Roman', 'TNR').replace('宋体', 'SimSun')
            parts.append(font_short)
        if self.size:
            # 转换为pt并简化
            pt = self.size / 2
            if pt == 12:
                parts.append('小四')
            elif pt == 14:
                parts.append('四号')
            elif pt == 16:
                parts.append('三号')
            elif pt == 18:
                parts.append('小二')
            else:
                parts.append(f'{int(pt)}pt')
        if self.bold:
            parts.append('加粗')
        if self.italic:
            parts.append('斜体')
        if self.underline:
            parts.append('下划线')
        if self.strike or self.dstrike:
            parts.append('删除线')
        if self.vert_align == 'superscript':
            parts.append('上标')
        if self.vert_align == 'subscript':
            parts.append('下标')
        if self.small_caps:
            parts.append('小型大写')
        if self.highlight:
            parts.append(f'高亮{self.highlight}')
        if self.color and self.color != '000000':
            parts.append(f'色{self.color}')
        
        return '/'.join(parts) if parts else '默认'
    
    def get_short_hint(self) -> str:
        """生成短格式提示（用于chunk标签）"""
        hints = []
        if self.bold:
            hints.append('b')
        if self.italic:
            hints.append('i')
        if self.underline:
            hints.append('u')
        if self.strike or self.dstrike:
            hints.append('s')
        if self.vert_align:
            hints.append(self.vert_align[:3])
        
        # 如果字体与常见基准不同，添加字体提示
        font = self.font_east_asia or self.font_ascii or "宋体"
        if font and font not in ['宋体', 'SimSun']:
            font_short = font.replace('Times New Roman', 'TNR')
            if not hints:
                hints.append(font_short)
            else:
                hints.append(font_short)
        
        return ','.join(hints) if hints else 'fmt'


class DocxPreprocessor:
    """Word文档预处理器（格式类别注册架构）"""

    NS = {
        'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
        'w14': 'http://schemas.microsoft.com/office/word/2010/wordml',
        'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    }

    def __init__(self, docx_path: str):
        self.docx_path = Path(docx_path)
        self.paragraphs: List[Dict[str, Any]] = []
        # 格式类别注册表
        self.format_categories: Dict[Tuple, Dict[str, Any]] = {}  # key -> {id, rpr_xml, count, label}
        self.category_counter = 0
        # 存储每个段落的run信息（用于后续生成chunks）
        self.para_runs_data: Dict[str, List[Dict[str, Any]]] = {}  # para_id -> runs_data

    def process(self, output_dir: str):
        """执行预处理"""
        output_path = Path(output_dir)
        output_path.mkdir(parents=True, exist_ok=True)

        # 1. 解压docx到unzipped/
        unzipped_dir = output_path / 'unzipped'
        self._unzip_docx(unzipped_dir)

        # 2. 计算文档哈希
        doc_hash = self._compute_doc_hash()
        print(f"文档哈希: {doc_hash[:16]}...")

        # 3. 遍历所有段落，为每个段落单独收集格式类别
        print("收集段落格式类别...")
        self._collect_para_format_categories()
        
        print(f"共处理 {len(self.para_format_data)} 个段落")

        # 4. 生成chunks/
        chunks_dir = output_path / 'chunks'
        self._generate_chunks(chunks_dir)

        # 5. 生成每段落独立的format_registry.json
        self._generate_format_registries(output_path / 'format_registry.json', doc_hash)

        print(f"处理完成！")
        print(f"  - unzipped/: {unzipped_dir}")
        print(f"  - chunks/: {chunks_dir}")
        print(f"  - format_registry.json: {output_path / 'format_registry.json'}")

    def _compute_doc_hash(self) -> str:
        """计算文档哈希"""
        with open(self.docx_path, 'rb') as f:
            return hashlib.sha256(f.read()).hexdigest()

    def _unzip_docx(self, output_dir: Path):
        """解压docx文件"""
        output_dir.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
        print(f"已解压到: {output_dir}")

    def _collect_para_format_categories(self):
        """为每个段落单独收集格式类别"""
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            doc_xml = zip_ref.read('word/document.xml')

        root = ET.fromstring(doc_xml)
        paragraphs = root.findall('.//w:p', self.NS)

        # 存储每个段落的格式数据
        self.para_format_data: Dict[str, Dict[str, Any]] = {}  # para_id -> {categories, baseline_id, runs_data, text}

        for para in paragraphs:
            para_id = para.get(f'{{{self.NS["w14"]}}}paraId', 'unknown')
            
            # 提取段落中的所有run格式
            runs_data = self._extract_run_formats_from_para(para)
            
            if not runs_data:
                continue
            
            # 为该段落收集格式类别
            para_categories: Dict[Tuple, Dict[str, Any]] = {}
            
            for run in runs_data:
                fmt = run['format']
                rpr_xml = run['rpr_xml']
                
                if fmt is None:
                    continue
                
                key = fmt.to_category_key()
                
                if key not in para_categories:
                    para_categories[key] = {
                        'id': None,
                        'rpr_xml': rpr_xml,
                        'count': 0,
                        'label': fmt.get_label(),
                        'format': fmt,
                    }
                
                para_categories[key]['count'] += 1
            
            if not para_categories:
                continue
            
            # 为段落内格式分配ID，选择出现最多的作为基准
            sorted_cats = sorted(
                para_categories.items(),
                key=lambda x: x[1]['count'],
                reverse=True
            )
            
            for idx, (key, data) in enumerate(sorted_cats):
                cat_id = f"F{idx}"
                para_categories[key]['id'] = cat_id
            
            baseline_id = sorted_cats[0][1]['id'] if sorted_cats else None
            
            # 生成段落文本
            para_text = self._process_formatting_with_para_categories(runs_data, para_categories, baseline_id)
            
            if para_text.strip():
                para_text = self._convert_to_latex(para_text)
                self.para_format_data[para_id] = {
                    'categories': para_categories,
                    'baseline_id': baseline_id,
                    'runs_data': runs_data,
                    'text': para_text,
                }
                self.paragraphs.append({
                    'para_id': para_id,
                    'text': para_text
                })

    def _extract_run_formats_from_para(self, para) -> List[Dict[str, Any]]:
        """从段落中提取所有run的格式信息"""
        runs_data = []
        
        for child in para:
            tag = child.tag
            local_name = tag.split('}')[-1] if '}' in tag else tag
            
            if local_name == 'r':
                fmt, rpr_xml = self._extract_run_format_and_xml(child)
                text, formula_parts = self._extract_run_text_with_formula(child)
                if text or formula_parts:
                    runs_data.append({
                        'format': fmt,
                        'rpr_xml': rpr_xml,
                        'text': text,
                        'formulas': formula_parts
                    })
            elif local_name in ('oMath', 'oMathPara'):
                formula_text = self._extract_formula(child)
                if formula_text:
                    runs_data.append({
                        'format': None,
                        'rpr_xml': None,
                        'text': '',
                        'formulas': [f'${formula_text}$']
                    })
        
        return runs_data

    def _extract_run_format_and_xml(self, run) -> Tuple[Optional[RunFormat], Optional[str]]:
        """从run元素提取格式属性和原始rPr XML"""
        fmt = RunFormat()
        rpr_xml = None
        
        rpr = run.find('.//w:rPr', self.NS)
        if rpr is not None:
            # 保存原始rPr XML（用于注册表）
            rpr_xml = ET.tostring(rpr, encoding='unicode')
            
            # 字体
            rfonts = rpr.find('.//w:rFonts', self.NS)
            if rfonts is not None:
                fmt.font_ascii = rfonts.get(f'{{{self.NS["w"]}}}ascii', '')
                fmt.font_east_asia = rfonts.get(f'{{{self.NS["w"]}}}eastAsia', '')
            
            # 字号
            sz = rpr.find('.//w:sz', self.NS)
            if sz is not None:
                val = sz.get(f'{{{self.NS["w"]}}}val')
                if val:
                    try:
                        fmt.size = int(val)
                    except:
                        pass
            
            # 加粗
            b = rpr.find('.//w:b', self.NS)
            if b is not None:
                val = b.get(f'{{{self.NS["w"]}}}val', 'true')
                fmt.bold = val.lower() != 'false'
            
            # 斜体
            i = rpr.find('.//w:i', self.NS)
            if i is not None:
                val = i.get(f'{{{self.NS["w"]}}}val', 'true')
                fmt.italic = val.lower() != 'false'
            
            # 高亮
            highlight = rpr.find('.//w:highlight', self.NS)
            if highlight is not None:
                fmt.highlight = highlight.get(f'{{{self.NS["w"]}}}val', '')
            
            # 字体颜色
            color = rpr.find('.//w:color', self.NS)
            if color is not None:
                fmt.color = color.get(f'{{{self.NS["w"]}}}val', '')
            
            # 下划线
            u = rpr.find('.//w:u', self.NS)
            if u is not None:
                val = u.get(f'{{{self.NS["w"]}}}val', '')
                if not val or val == 'none':
                    fmt.underline = 'single'  # 元素存在即表示有下划线
                else:
                    fmt.underline = val
            
            # 单删除线
            strike = rpr.find('.//w:strike', self.NS)
            if strike is not None:
                val = strike.get(f'{{{self.NS["w"]}}}val', 'true')
                fmt.strike = val.lower() != 'false'
            
            # 双删除线
            dstrike = rpr.find('.//w:dstrike', self.NS)
            if dstrike is not None:
                val = dstrike.get(f'{{{self.NS["w"]}}}val', 'true')
                fmt.dstrike = val.lower() != 'false'
            
            # 上下标
            vert_align = rpr.find('.//w:vertAlign', self.NS)
            if vert_align is not None:
                fmt.vert_align = vert_align.get(f'{{{self.NS["w"]}}}val', '')
            
            # 小型大写字母
            small_caps = rpr.find('.//w:smallCaps', self.NS)
            if small_caps is not None:
                val = small_caps.get(f'{{{self.NS["w"]}}}val', 'true')
                fmt.small_caps = val.lower() != 'false'
        
        return fmt, rpr_xml

    def _assign_category_ids(self) -> str:
        """分配类别ID，返回基准类别ID"""
        # 按出现次数排序
        sorted_categories = sorted(
            self.format_categories.items(),
            key=lambda x: x[1]['count'],
            reverse=True
        )
        
        # 分配ID（F0, F1, F2...）
        for idx, (key, data) in enumerate(sorted_categories):
            cat_id = f"F{idx}"
            self.format_categories[key]['id'] = cat_id
        
        # 返回出现次数最多的类别ID作为基准
        return sorted_categories[0][1]['id'] if sorted_categories else "F0"

    def _get_key_by_id(self, cat_id: str) -> Optional[Tuple]:
        """根据类别ID获取key"""
        for key, data in self.format_categories.items():
            if data['id'] == cat_id:
                return key
        return None

    def _process_formatting_with_para_categories(self, runs_data: List[Dict[str, Any]], 
                                                    para_categories: Dict[Tuple, Dict], 
                                                    baseline_id: str) -> str:
        """处理段落内所有run的格式标记（使用段落内类别ID标签）"""
        if not runs_data:
            return ''
        
        result_parts = []
        current_cat_id = None
        current_text_parts = []
        
        def flush_current():
            """刷新当前标记块"""
            nonlocal current_cat_id, current_text_parts
            if current_text_parts:
                text = ''.join(current_text_parts)
                if current_cat_id and current_cat_id != baseline_id:
                    # 非基准格式，添加类别ID标签
                    # 获取类别信息用于生成提示
                    key = self._get_key_by_id_from_para(para_categories, current_cat_id)
                    if key:
                        hint = para_categories[key]['format'].get_short_hint()
                        result_parts.append(f"‹{current_cat_id}:{hint}›{text}‹/›")
                    else:
                        result_parts.append(f"‹{current_cat_id}:fmt›{text}‹/›")
                else:
                    # 基准格式，直接输出纯文本
                    result_parts.append(text)
                current_text_parts = []
                current_cat_id = None
        
        for run in runs_data:
            # 处理公式
            if run['format'] is None:
                flush_current()
                for formula in run['formulas']:
                    result_parts.append(formula)
                continue
            
            fmt = run['format']
            text = run['text']
            formulas = run['formulas']
            
            # 获取类别ID
            key = fmt.to_category_key()
            cat_id = para_categories[key]['id'] if key in para_categories else baseline_id
            
            # 纯空格也要添加格式标签，保留空格格式
            if not text.strip() and not formulas:
                if text:
                    if cat_id != current_cat_id:
                        flush_current()
                        current_cat_id = cat_id
                    current_text_parts.append(text)
                continue
            
            # 如果有公式，需要分段处理
            if formulas:
                # 先处理公式前的文本
                if text:
                    if cat_id != current_cat_id:
                        flush_current()
                        current_cat_id = cat_id
                    current_text_parts.append(text)
                    flush_current()
                
                # 添加公式（公式不加标记）
                for formula in formulas:
                    result_parts.append(formula)
            else:
                # 无公式，正常处理
                if cat_id != current_cat_id:
                    flush_current()
                    current_cat_id = cat_id
                current_text_parts.append(text)
        
        flush_current()
        return ''.join(result_parts)

    def _get_key_by_id_from_para(self, para_categories: Dict[Tuple, Dict], cat_id: str) -> Optional[Tuple]:
        """根据段落内类别ID获取key"""
        for key, data in para_categories.items():
            if data['id'] == cat_id:
                return key
        return None

    def _convert_to_latex(self, text: str) -> str:
        """将 Unicode 数学字符转换为 LaTeX 格式"""
        char_map = {
            '∈': r'\in', '∉': r'\notin', '⊂': r'\subset', '⊆': r'\subseteq',
            '∪': r'\cup', '∩': r'\cap', '∅': r'\emptyset', '∞': r'\infty',
            '∑': r'\sum', '∏': r'\prod', '∫': r'\int', '∂': r'\partial',
            '∇': r'\nabla', '√': r'\sqrt', '∆': r'\Delta', '∏': r'\Pi',
            'Σ': r'\Sigma', 'Ω': r'\Omega', 'α': r'\alpha', 'β': r'\beta',
            'γ': r'\gamma', 'δ': r'\delta', 'ε': r'\epsilon', 'θ': r'\theta',
            'λ': r'\lambda', 'μ': r'\mu', 'π': r'\pi', 'ρ': r'\rho',
            'σ': r'\sigma', 'τ': r'\tau', 'φ': r'\phi', 'ω': r'\omega',
            '→': r'\rightarrow', '←': r'\leftarrow', '↔': r'\leftrightarrow',
            '⇒': r'\Rightarrow', '⇐': r'\Leftarrow', '⇔': r'\Leftrightarrow',
            '≤': r'\leq', '≥': r'\geq', '≠': r'\neq', '≈': r'\approx',
            '≡': r'\equiv', '±': r'\pm', '∓': r'\mp', '×': r'\times ',
            '÷': r'\div ', '·': r'\cdot ', 'ℝ': r'\mathbb{R}', 'ℕ': r'\mathbb{N}',
            'ℤ': r'\mathbb{Z}', 'ℂ': r'\mathbb{C}', '°': r'^\circ',
            '′': r"'", '″': r"\prime\prime",
        }
        for unicode_char, latex_char in char_map.items():
            text = text.replace(unicode_char, latex_char)
        
        import re
        text = re.sub(r'\\times\s+', r'\\times ', text)
        text = re.sub(r'\\div\s+', r'\\div ', text)
        text = re.sub(r'\\cdot\s+', r'\\cdot ', text)
        
        return text

    def _extract_run_text_with_formula(self, run) -> Tuple[str, List[str]]:
        """提取run中的文本和公式，返回(文本, 公式列表)"""
        text_parts = []
        formula_parts = []
        
        for child in run:
            tag = child.tag
            local_name = tag.split('}')[-1] if '}' in tag else tag
            
            if local_name == 't':
                if child.text:
                    text_parts.append(child.text)
            elif local_name in ('oMath', 'oMathPara'):
                formula_text = self._extract_formula(child)
                if formula_text:
                    formula_parts.append(f'${formula_text}$')
            elif local_name == 'object':
                text_parts.append(self._extract_object_text(child))
        
        return ''.join(text_parts), formula_parts

    def _extract_object_text(self, obj) -> str:
        """提取object元素中的文本"""
        text_parts = []
        for child in obj:
            tag = child.tag
            local_name = tag.split('}')[-1] if '}' in tag else tag
            if local_name in ('oMath', 'oMathPara'):
                formula_text = self._extract_formula(child)
                if formula_text:
                    text_parts.append(f'${formula_text}$')
        return ''.join(text_parts)

    def _extract_formula(self, elem) -> str:
        """从OMML元素提取公式文本"""
        formula_parts = []
        self._extract_omml_text(elem, formula_parts)
        return ''.join(formula_parts)

    def _extract_omml_text(self, elem, parts):
        """递归提取 OMML 元素中的文本，输出 LaTeX 格式"""
        m_ns = self.NS['m']

        if elem.tag.endswith('}t') or elem.tag == f'{{{m_ns}}}t':
            if elem.text:
                text = elem.text.replace('{', r'\{').replace('}', r'\}')
                parts.append(text)
        elif elem.tag.endswith('}sub') or elem.tag == f'{{{m_ns}}}sub':
            sub_parts = []
            for child in elem:
                self._extract_omml_text(child, sub_parts)
            if sub_parts:
                parts.append(f'_{{{ "".join(sub_parts) }}}')
        elif elem.tag.endswith('}sup') or elem.tag == f'{{{m_ns}}}sup':
            sup_parts = []
            for child in elem:
                self._extract_omml_text(child, sup_parts)
            if sup_parts:
                parts.append(f'^{{{ "".join(sup_parts) }}}')
        elif elem.tag.endswith('}f') or elem.tag == f'{{{m_ns}}}f':
            num_parts = []
            den_parts = []
            num = elem.find(f'{{{m_ns}}}num')
            if num is not None:
                for child in num:
                    self._extract_omml_text(child, num_parts)
            den = elem.find(f'{{{m_ns}}}den')
            if den is not None:
                for child in den:
                    self._extract_omml_text(child, den_parts)
            num_str = ''.join(num_parts) if num_parts else ''
            den_str = ''.join(den_parts) if den_parts else ''
            parts.append(f'\\frac{{{num_str}}}{{{den_str}}}')
        elif elem.tag.endswith('}rad') or elem.tag == f'{{{m_ns}}}rad':
            e = elem.find(f'{{{m_ns}}}e')
            e_parts = []
            if e is not None:
                for child in e:
                    self._extract_omml_text(child, e_parts)
            e_str = ''.join(e_parts) if e_parts else ''
            deg = elem.find(f'{{{m_ns}}}deg')
            if deg is not None:
                deg_parts = []
                for child in deg:
                    self._extract_omml_text(child, deg_parts)
                deg_str = ''.join(deg_parts) if deg_parts else ''
                if deg_str:
                    parts.append(f'\\sqrt[{deg_str}]{{{e_str}}}')
                else:
                    parts.append(f'\\sqrt{{{e_str}}}')
            else:
                parts.append(f'\\sqrt{{{e_str}}}')
        else:
            for child in elem:
                self._extract_omml_text(child, parts)

    def _generate_chunks(self, chunks_dir: Path, max_chars: int = 1000):
        """生成chunks文件"""
        chunks_dir.mkdir(parents=True, exist_ok=True)
        
        origin_chunks_dir = chunks_dir.parent / 'origin_chunks'
        origin_chunks_dir.mkdir(parents=True, exist_ok=True)

        chunk_idx = 0
        current_chunk_lines = []
        current_chunk_chars = 0

        for para in self.paragraphs:
            para_id = para['para_id']
            text = para['text']
            text_len = len(text)

            if text_len > max_chars:
                if current_chunk_lines:
                    self._save_chunk(chunks_dir, chunk_idx, current_chunk_lines)
                    chunk_idx += 1
                    current_chunk_lines = []
                    current_chunk_chars = 0

                line = f"[{para_id}] {text}"
                self._save_chunk(chunks_dir, chunk_idx, [line])
                chunk_idx += 1
                continue

            if current_chunk_chars + text_len > max_chars:
                if current_chunk_lines:
                    self._save_chunk(chunks_dir, chunk_idx, current_chunk_lines)
                    chunk_idx += 1
                    current_chunk_lines = []
                    current_chunk_chars = 0

            line = f"[{para_id}] {text}"
            current_chunk_lines.append(line)
            current_chunk_chars += text_len

        if current_chunk_lines:
            self._save_chunk(chunks_dir, chunk_idx, current_chunk_lines)

        self._copy_chunks_to_origin(chunks_dir, origin_chunks_dir)

        print(f"已生成 {chunk_idx + 1} 个chunk文件")
        print(f"  - chunks/: {chunks_dir}")
        print(f"  - origin_chunks/: {origin_chunks_dir}")

    def _copy_chunks_to_origin(self, chunks_dir: Path, origin_chunks_dir: Path):
        """将chunks复制到origin_chunks目录"""
        import shutil
        for chunk_file in sorted(chunks_dir.glob('chunk_*.md')):
            dest_file = origin_chunks_dir / chunk_file.name
            shutil.copy2(chunk_file, dest_file)

    def _save_chunk(self, chunks_dir: Path, idx: int, lines: List[str]):
        """保存单个chunk文件"""
        chunk_file = chunks_dir / f"chunk_{idx}.md"
        with open(chunk_file, 'w', encoding='utf-8') as f:
            f.write('\n'.join(lines))

    def _generate_format_registries(self, output_path: Path, doc_hash: str):
        """生成每个段落独立的格式注册表JSON文件"""
        registry = {
            'doc_hash': doc_hash,
            'paragraphs': {}
        }
        
        for para_id, para_data in self.para_format_data.items():
            categories = para_data['categories']
            baseline_id = para_data['baseline_id']
            
            para_categories = {}
            for key, data in categories.items():
                cat_id = data['id']
                para_categories[cat_id] = {
                    'label': data['label'],
                    'rpr_xml': data['rpr_xml'],
                    'count': data['count']
                }
            
            registry['paragraphs'][para_id] = {
                'baseline': baseline_id,
                'categories': para_categories
            }
        
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(registry, f, ensure_ascii=False, indent=2)
        
        print(f"已生成格式注册表: {output_path}")
        print(f"  共 {len(registry['paragraphs'])} 个段落")


def main():
    import sys

    if len(sys.argv) < 2:
        print("用法: python docx_preprocessor.py <docx文件> [输出目录]")
        print("示例: python docx_preprocessor.py template.docx output/")
        sys.exit(1)

    docx_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else 'output'

    print(f"正在处理: {docx_path}")
    print(f"输出目录: {output_dir}")

    processor = DocxPreprocessor(docx_path)
    processor.process(output_dir)


if __name__ == '__main__':
    main()
