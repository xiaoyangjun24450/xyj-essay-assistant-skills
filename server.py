#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
论文改写助手 - 本地 Web 服务
仅使用标准库：通过 HTTP 提供 HTML 界面与 API，读写 config、提示词，并执行改写流程。
"""

import json
import re
import sys
from http.server import HTTPServer, BaseHTTPRequestHandler
from pathlib import Path
from urllib.parse import urlparse, unquote

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


def run_pipeline(docx_path: str, requirement: str):
    """在调用线程中执行预处理 -> 改写 -> 重建，返回 (success, message, output_path_or_error)。"""
    import datetime
    from docx_preprocessor import DocxPreprocessor
    from chunks_rewriter import run_rewrite
    from docx_chunks_restorer import DocxChunksRestorer

    docx_path = Path(docx_path).expanduser().resolve()
    if not docx_path.exists():
        return False, f'文件不存在: {docx_path}', None

    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    work_dir = SCRIPT_DIR / "output" / timestamp
    work_dir.mkdir(parents=True, exist_ok=True)

    try:
        config = load_config()
        max_chars = config.get('chunk_max_chars', 1000)
        concurrency = config.get('concurrency', 3)

        processor = DocxPreprocessor(str(docx_path))
        processor.process(str(work_dir), max_chars=max_chars)

        chunks_dir = work_dir / "chunks"
        chunks_new_dir = work_dir / "chunks_new"
        logs_dir = work_dir / "logs"
        run_rewrite(str(chunks_dir), str(chunks_new_dir), requirement, str(logs_dir), concurrency=concurrency)

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
            if 'multipart/form-data' in content_type:
                docx_path, requirement = self._parse_run_multipart(content_type)
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
                self.send_json({'ok': False, 'error': '请填写改写需求'}, 400)
                return
            if not docx_path:
                self.send_json({'ok': False, 'error': '请选择要上传的 docx 文件'}, 400)
                return

            ok, message, output_path = run_pipeline(docx_path, requirement)
            self.send_json({
                'ok': ok,
                'message': message,
                'output_docx': output_path,
            })
        except Exception as e:
            self.send_json({'ok': False, 'error': str(e)}, 500)

    def _parse_run_multipart(self, content_type):
        """解析 multipart/form-data（不依赖 cgi），保存上传的 docx，返回 (保存后的路径, requirement)。"""
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
        if not docx_data:
            return None, requirement
        safe_name = re.sub(r'[^\w\.\-]', '_', docx_filename)
        upload_dir = SCRIPT_DIR / 'output' / 'uploads'
        upload_dir.mkdir(parents=True, exist_ok=True)
        timestamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
        save_path = upload_dir / f'{timestamp}_{safe_name}'
        save_path.write_bytes(docx_data)
        return str(save_path), requirement

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
