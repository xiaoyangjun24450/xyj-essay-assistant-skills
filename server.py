#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文改写助手 - 本地 Web 服务
仅使用标准库：通过 HTTP 提供 HTML 界面与 API，读写 config、提示词，并执行改写流程。
"""

import json
import io
import importlib
import re
import sys
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, unquote
import xml.etree.ElementTree as ET
import zipfile

# 项目根目录
if getattr(sys, 'frozen', False):
    SCRIPT_DIR = Path(sys.executable).parent
else:
    SCRIPT_DIR = Path(__file__).resolve().parent

sys.path.insert(0, str(SCRIPT_DIR / 'scripts'))


def get_config_path():
    return SCRIPT_DIR / 'config.json'


def get_prompts_dir():
    return SCRIPT_DIR / 'prompts'


def load_config():
    p = get_config_path()
    if not p.exists():
        return {}
    with open(p, 'r', encoding='utf-8') as f:
        return json.load(f)


def save_config(data):
    with open(get_config_path(), 'w', encoding='utf-8') as f:
        json.dump(data, f, ensure_ascii=False, indent=2)


def load_prompt(name):
    """name: 'system_prompt' | 'fix_prompt'"""
    path = get_prompts_dir() / f'{name}.txt'
    if not path.exists():
        return ''
    return path.read_text(encoding='utf-8')


def save_prompt(name, content):
    path = get_prompts_dir() / f'{name}.txt'
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(content, encoding='utf-8')


DOCX_NS = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
}


def extract_docx_text(docx_data):
    try:
        with zipfile.ZipFile(io.BytesIO(docx_data), 'r') as zip_ref:
            doc_xml = zip_ref.read('word/document.xml')
    except zipfile.BadZipFile as e:
        raise ValueError('上传的 docx 文件无效') from e
    except KeyError as e:
        raise ValueError('docx 文件缺少 document.xml') from e

    try:
        root = ET.fromstring(doc_xml)
    except ET.ParseError as e:
        raise ValueError('docx 内容解析失败') from e

    word_ns = DOCX_NS['w']
    text_tag = f'{{{word_ns}}}t'
    tab_tag = f'{{{word_ns}}}tab'
    br_tags = {f'{{{word_ns}}}br', f'{{{word_ns}}}cr'}
    paragraphs = []

    for para in root.findall('.//w:p', DOCX_NS):
        parts = []
        for node in para.iter():
            if node.tag == text_tag and node.text:
                parts.append(node.text)
            elif node.tag == tab_tag:
                parts.append('\t')
            elif node.tag in br_tags:
                parts.append('\n')
        para_text = ''.join(parts).strip()
        if para_text:
            paragraphs.append(para_text)

    return '\n'.join(paragraphs).strip()


def extract_material_text(filename, data):
    ext = Path(filename).suffix.lower()
    if ext == '.docx':
        return extract_docx_text(data)
    if ext not in {'.txt', '.md'}:
        raise ValueError(f'语料素材仅支持 .docx、.txt、.md：{filename}')
    return data.decode('utf-8-sig', errors='replace').strip()


def merge_requirement_with_materials(requirement, material_texts):
    valid_materials = [
        {'name': item.get('name', ''), 'text': item.get('text', '').strip()}
        for item in material_texts
        if item.get('text', '').strip()
    ]
    if not valid_materials:
        return requirement

    prompt_parts = [requirement.strip(), "", "## 参考语料素材"]
    for idx, item in enumerate(valid_materials, start=1):
        prompt_parts.extend([
            "",
            f"### 素材{idx}：{item.get('name', f'素材{idx}')}",
            item['text'],
        ])

    return "\n".join(prompt_parts).strip()


def run_pipeline(docx_path: str, requirement: str, material_paths=None):
    """在调用线程中执行预处理 -> 分段 -> 切块 -> 检索改写 -> 重建。"""
    import datetime

    DocxPreprocessor = importlib.import_module('1_docx_preprocessor').DocxPreprocessor
    SegmentAnalyzer = importlib.import_module('2_segment_analyzer').SegmentAnalyzer
    SegmentValidator = importlib.import_module('3_segment_validator').SegmentValidator
    ChunkGenerator = importlib.import_module('4_chunk_generator').ChunkGenerator
    MaterialProcessor = importlib.import_module('5_material_processor').MaterialProcessor
    run_rewrite = importlib.import_module('7_chunks_rewriter').run_rewrite
    DocxChunksRestorer = importlib.import_module('8_docx_chunks_restorer').DocxChunksRestorer

    docx_path = Path(docx_path).expanduser().resolve()
    material_paths = [str(Path(path).expanduser().resolve()) for path in (material_paths or [])]
    if not docx_path.exists():
        return False, f'文件不存在: {docx_path}', None

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = SCRIPT_DIR / "output" / timestamp
    work_dir.mkdir(parents=True, exist_ok=True)

    try:
        config = load_config()
        max_chars = config.get('chunk_max_chars', 1000)
        concurrency = config.get('concurrency', 3)
        segment_max_window_chars = config.get('segment_max_window_chars', 8000)
        material_segment_max_window_chars = config.get('material_segment_max_window_chars', 8000)
        retrieval_limit = config.get('retrieval_limit', 3)

        processor = DocxPreprocessor(str(docx_path))
        processor.process(str(work_dir), max_chars=max_chars)

        logs_dir = work_dir / "logs"
        stage2_logs_dir = logs_dir / "stage2"
        stage3_logs_dir = logs_dir / "stage3"
        logs_dir.mkdir(parents=True, exist_ok=True)
        stage2_logs_dir.mkdir(parents=True, exist_ok=True)
        stage3_logs_dir.mkdir(parents=True, exist_ok=True)

        full_paragraphs_path = work_dir / "full_paragraphs.txt"
        stage2_segments_path = work_dir / "template_segments.stage2.json"
        stage3_segments_path = work_dir / "template_segments.json"
        analyzer = SegmentAnalyzer(
            full_paragraphs_path=str(full_paragraphs_path),
            output_path=str(stage2_segments_path),
            max_window_chars=segment_max_window_chars,
            log_dir=str(stage2_logs_dir),
        )
        analyzer.analyze()

        validator = SegmentValidator(
            full_paragraphs_path=str(full_paragraphs_path),
            input_segments_path=str(stage2_segments_path),
            output_path=str(stage3_segments_path),
            log_dir=str(stage3_logs_dir),
        )
        validator.run()

        chunks_dir = work_dir / "chunks"
        chunks_metadata_path = work_dir / "chunks_metadata.json"
        chunk_generator = ChunkGenerator(
            full_paragraphs_path=str(full_paragraphs_path),
            segments_path=str(stage3_segments_path),
            output_dir=str(chunks_dir),
            metadata_path=str(chunks_metadata_path),
            max_chars=max_chars,
        )
        chunk_generator.generate()

        chunks_new_dir = work_dir / "chunks_new"
        material_numbered_path = None
        material_segments_path = None
        if material_paths:
            materials_dir = work_dir / "materials"
            material_result = MaterialProcessor(
                material_paths=material_paths,
                output_dir=str(materials_dir),
                max_window_chars=material_segment_max_window_chars,
                log_dir=str(logs_dir),
            ).process()
            material_numbered_path = material_result["numbered_output_path"]
            material_segments_path = material_result["segments_output_path"]

        run_rewrite(
            str(chunks_dir),
            str(chunks_new_dir),
            requirement,
            str(logs_dir),
            concurrency=concurrency,
            chunks_metadata_path=str(chunks_metadata_path) if material_numbered_path else None,
            material_numbered_path=material_numbered_path,
            material_segments_path=material_segments_path,
            retrieval_limit=retrieval_limit,
        )

        output_docx = work_dir / "output.docx"
        restorer = DocxChunksRestorer(str(work_dir / "unzipped"), str(chunks_new_dir))
        restorer.restore(str(output_docx))

        return True, f'完成，输出: {output_docx}', str(output_docx)
    except Exception as e:
        return False, str(e), None


class Handler(BaseHTTPRequestHandler):
    def do_GET(self):
        path = urlparse(self.path).path
        if path == '/' or path == '/index.html':
            self.serve_html()
        elif path == '/api/config':
            self.serve_get_config()
        elif path == '/api/prompts/system':
            self.serve_get_prompt('system_prompt')
        elif path == '/api/prompts/fix':
            self.serve_get_prompt('fix_prompt')
        else:
            self.send_error(404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path == '/api/config':
            self.save_config()
        elif path == '/api/prompts/system':
            self.save_prompt('system_prompt')
        elif path == '/api/prompts/fix':
            self.save_prompt('fix_prompt')
        elif path == '/api/run':
            self.api_run()
        else:
            self.send_error(404)

    def read_body_json(self):
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length == 0:
            return None
        body = self.rfile.read(content_length)
        return json.loads(body.decode('utf-8'))

    def read_body_text(self):
        content_length = int(self.headers.get('Content-Length', 0))
        if content_length == 0:
            return ''
        return self.rfile.read(content_length).decode('utf-8')

    def send_json(self, data, status=200):
        self.send_response(status)
        self.send_header('Content-Type', 'application/json; charset=utf-8')
        self.end_headers()
        self.wfile.write(json.dumps(data, ensure_ascii=False).encode('utf-8'))

    def send_text(self, text, content_type='text/plain; charset=utf-8'):
        self.send_response(200)
        self.send_header('Content-Type', content_type)
        self.end_headers()
        self.wfile.write(text.encode('utf-8'))

    def serve_html(self):
        html_path = SCRIPT_DIR / 'index.html'
        if not html_path.exists():
            self.send_error(404, 'index.html not found')
            return
        self.send_response(200)
        self.send_header('Content-Type', 'text/html; charset=utf-8')
        self.end_headers()
        self.wfile.write(html_path.read_bytes())

    def serve_get_config(self):
        try:
            data = load_config()
            self.send_json(data)
        except Exception as e:
            self.send_json({'error': str(e)}, 500)

    def serve_get_prompt(self, name):
        try:
            text = load_prompt(name)
            self.send_text(text)
        except Exception as e:
            self.send_json({'error': str(e)}, 500)

    def save_config(self):
        try:
            data = self.read_body_json()
            if data is None:
                self.send_json({'error': '需要 JSON 请求体'}, 400)
                return
            current = load_config()
            if 'api' in data:
                current['api'] = {**current.get('api', {}), **data['api']}
            if 'concurrency' in data:
                current['concurrency'] = int(data['concurrency'])
            if 'chunk_max_chars' in data:
                current['chunk_max_chars'] = int(data['chunk_max_chars'])
            if 'segment_max_window_chars' in data:
                current['segment_max_window_chars'] = int(data['segment_max_window_chars'])
            if 'material_segment_max_window_chars' in data:
                current['material_segment_max_window_chars'] = int(data['material_segment_max_window_chars'])
            save_config(current)
            self.send_json({'ok': True})
        except Exception as e:
            self.send_json({'error': str(e)}, 500)

    def save_prompt(self, name):
        try:
            content = self.read_body_text()
            save_prompt(name, content)
            self.send_json({'ok': True})
        except Exception as e:
            self.send_json({'error': str(e)}, 500)

    def api_run(self):
        try:
            content_type = self.headers.get('Content-Type', '')
            material_paths = []
            if 'multipart/form-data' in content_type:
                docx_path, requirement, material_paths = self._parse_run_multipart(content_type)
            else:
                data = self.read_body_json()
                if not data:
                    self.send_json({'ok': False, 'error': '需要上传文件或 JSON'}, 400)
                    return
                docx_path = data.get('docx_path', '').strip()
                requirement = data.get('requirement', '').strip()
                if not docx_path or not requirement:
                    self.send_json({'ok': False, 'error': '需要 docx_path 与 requirement'}, 400)
                    return

            if not requirement:
                self.send_json({'ok': False, 'error': '请填写写作任务'}, 400)
                return
            if not docx_path:
                self.send_json({'ok': False, 'error': '请先上传格式模板 docx'}, 400)
                return

            ok, message, output_path = run_pipeline(docx_path, requirement, material_paths=material_paths)
            self.send_json({
                'ok': ok,
                'message': message,
                'output_docx': output_path,
            })
        except Exception as e:
            self.send_json({'ok': False, 'error': str(e)}, 500)

    def _parse_run_multipart(self, content_type):
        """解析 multipart/form-data（不依赖 cgi），返回 (docx_path, requirement, material_paths)。"""
        import datetime
        # 取 boundary
        m = re.search(r'boundary=([^;\s]+)', content_type)
        boundary = m.group(1).strip('"').encode() if m else None
        if not boundary:
            raise ValueError('multipart 缺少 boundary')
        content_length = int(self.headers.get('Content-Length', 0))
        body = self.rfile.read(content_length)
        # 按 boundary 拆成 part（首段可能是空或 preamble）
        sep = b'\r\n--' + boundary
        parts = body.split(sep)
        requirement = ''
        docx_data, docx_filename = None, 'upload.docx'
        material_files = []
        for part in parts:
            if part.strip() in (b'', b'--'):
                continue
            head, _, raw = part.partition(b'\r\n\r\n')
            value = raw.rstrip(b'\r\n')
            # 从头部解析 name / filename
            disp = None
            for line in head.split(b'\r\n'):
                if line.lower().startswith(b'content-disposition:'):
                    disp = line.decode('utf-8', errors='replace')
                    break
            if not disp:
                continue
            name_m = re.search(r'name="([^"]*)"', disp)
            filename_m = re.search(r'filename="([^"]*)"', disp)
            name = name_m.group(1) if name_m else ''
            if name == 'requirement':
                requirement = value.decode('utf-8', errors='replace').strip()
            elif name == 'docx' and filename_m:
                docx_filename = unquote(filename_m.group(1))
                docx_data = value
            elif name == 'materials' and filename_m:
                material_filename = unquote(filename_m.group(1))
                material_files.append((material_filename, value))
        if not docx_data:
            return None, requirement, []
        if Path(docx_filename).suffix.lower() != '.docx':
            raise ValueError('格式模板仅支持 .docx')
        safe_name = re.sub(r'[^\w\.\-]', '_', docx_filename)
        upload_dir = SCRIPT_DIR / 'output' / 'uploads'
        upload_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        save_path = upload_dir / f'{timestamp}_{safe_name}'
        save_path.write_bytes(docx_data)

        material_paths = []
        for index, (material_filename, material_data) in enumerate(material_files, start=1):
            safe_material_name = re.sub(r'[^\w\.\-]', '_', material_filename)
            material_path = upload_dir / f'{timestamp}_material_{index:02d}_{safe_material_name}'
            material_path.write_bytes(material_data)
            material_paths.append(str(material_path))

        return str(save_path), requirement, material_paths

    def log_message(self, format, *args):
        print(format % args)


def main():
    port = 8765
    if len(sys.argv) > 1:
        try:
            port = int(sys.argv[1])
        except ValueError:
            print('用法: python3 server.py [端口号]')
            sys.exit(1)
    try:
        server = HTTPServer(('127.0.0.1', port), Handler)
    except OSError as e:
        if e.errno == 98:  # Address already in use
            print(f'端口 {port} 已被占用。可改用其它端口，例如: python3 server.py {port + 1}')
            print('或先结束占用该端口的进程: lsof -i :%d' % port)
        else:
            raise
        sys.exit(1)
    print(f'论文改写助手 Web 服务: http://127.0.0.1:{port}/')
    print('按 Ctrl+C 停止')
    server.serve_forever()


if __name__ == '__main__':
    main()
