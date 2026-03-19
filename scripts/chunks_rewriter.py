#!/usr/bin/env python3
"""Rewrite chunks using Moonshot API."""

import os
import sys
import re
import json
from pathlib import Path
from typing import List, Tuple, Optional
import httpx
from concurrent.futures import ThreadPoolExecutor, as_completed


def load_config() -> dict:
    """加载配置文件"""
    # 判断是否为打包后的 exe 环境
    if getattr(sys, 'frozen', False):
        # 打包后的 exe 环境，配置文件在 exe 同级目录
        config_path = Path(sys.executable).parent / 'config.json'
    else:
        # 开发环境，配置文件在项目根目录
        config_path = Path(__file__).parent.parent / 'config.json'
    
    if not config_path.exists():
        raise FileNotFoundError(f"配置文件未找到: {config_path}")
    
    with open(config_path, 'r', encoding='utf-8') as f:
        return json.load(f)


# 加载配置
_config = load_config()
API_KEY = _config["api"]["api_key"]
MODEL = _config["api"]["model"]
BASE_URL = _config["api"]["base_url"]


def read_file(path: str) -> str:
    """Read file contents."""
    fp = Path(path).expanduser().resolve()
    if not fp.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if not fp.is_file():
        raise ValueError(f"Not a file: {path}")
    return fp.read_text(encoding="utf-8")


def write_file(path: str, content: str) -> None:
    """Write content to a file."""
    fp = Path(path).expanduser().resolve()
    fp.parent.mkdir(parents=True, exist_ok=True)
    fp.write_text(content, encoding="utf-8")


def load_system_prompt() -> str:
    """Load rewrite system prompt from file."""
    if getattr(sys, 'frozen', False):
        prompt_path = Path(sys.executable).parent / 'prompts' / 'system_prompt.txt'
    else:
        prompt_path = Path(__file__).parent.parent / 'prompts' / 'system_prompt.txt'

    if not prompt_path.exists():
        raise FileNotFoundError(f"系统提示词文件未找到: {prompt_path}")

    with open(prompt_path, 'r', encoding='utf-8') as f:
        return f.read().strip()


def load_fix_prompt() -> str:
    """Load fix system prompt for correcting validation errors."""
    if getattr(sys, 'frozen', False):
        prompt_path = Path(sys.executable).parent / 'prompts' / 'fix_prompt.txt'
    else:
        prompt_path = Path(__file__).parent.parent / 'prompts' / 'fix_prompt.txt'

    if not prompt_path.exists():
        raise FileNotFoundError(f"修复提示词文件未找到: {prompt_path}")

    with open(prompt_path, 'r', encoding='utf-8') as f:
        return f.read().strip()


def extract_para_id(line: str) -> Optional[str]:
    """从行中提取 PARA_ID，格式为 [XXXXXXXX]"""
    match = re.match(r'^\[([0-9A-Fa-f]{8})\]', line)
    if match:
        return match.group(1).upper()
    return None


def extract_tag_types(line: str) -> set:
    """从行中提取所有格式标签的种类"""
    tag_types = set()
    tag_pair_pattern = r'<([fisbzchuv,=\w]+)>([^<]*)</\1>'
    
    for match in re.finditer(tag_pair_pattern, line):
        attrs = match.group(1)
        for attr in attrs.split(','):
            if '=' in attr:
                tag_types.add(attr.strip())
    
    return tag_types


def verify_format_tags(line: str) -> Tuple[bool, Optional[str]]:
    """
    校验行中的格式标签是否匹配
    返回: (是否通过, 错误信息)
    """
    tag_pattern = r'<([fisbzchuv,=\w]*)>|</([fisbzchuv,=\w]*)>'
    tags = list(re.finditer(tag_pattern, line))
    
    if not tags:
        return True, None
    
    stack = []
    
    for match in tags:
        open_tag = match.group(1)
        close_tag = match.group(2)
        
        if open_tag is not None:
            stack.append((open_tag, match.start()))
        elif close_tag is not None:
            if not stack:
                return False, f"多余的闭标签 </{close_tag}>"
            
            expected_open, open_pos = stack.pop()
            if expected_open != close_tag:
                return False, f"标签不匹配: <{expected_open}> 与 </{close_tag}>"
    
    if stack:
        remaining_tag, pos = stack[0]
        return False, f"标签未闭合: <{remaining_tag}>"
    
    return True, None


def validate_rewritten_content(
    original_lines: List[str],
    rewritten_lines: List[str],
    filename: str
) -> List[Tuple[int, str, str]]:
    """
    校验改写后的内容
    返回: [(行号, 错误类型, 错误信息), ...]
    """
    errors = []
    
    # 检查行数
    if len(rewritten_lines) != len(original_lines):
        errors.append((
            0,
            "行数不匹配",
            f"行数不一致: 原文={len(original_lines)}, 改写后={len(rewritten_lines)}"
        ))
        return errors
    
    for i, (orig_line, rewritten_line) in enumerate(zip(original_lines, rewritten_lines)):
        line_num = i + 1
        
        # 检查 PARA_ID
        orig_para_id = extract_para_id(orig_line)
        rewritten_para_id = extract_para_id(rewritten_line)
        
        if orig_para_id is None:
            continue  # 原文没有 PARA_ID，跳过
        
        if rewritten_para_id is None:
            errors.append((line_num, "缺少PARA_ID", f"第{line_num}行缺少PARA_ID"))
        elif orig_para_id != rewritten_para_id:
            errors.append((
                line_num,
                "PARA_ID不匹配",
                f"PARA_ID不匹配: 原文={orig_para_id}, 改写后={rewritten_para_id}"
            ))
        
        # 检查格式标签
        passed, error_msg = verify_format_tags(rewritten_line)
        if not passed:
            errors.append((line_num, "格式标签错误", error_msg))
        
        # 检查是否缺少格式标签类型
        orig_tag_types = extract_tag_types(orig_line)
        rewritten_tag_types = extract_tag_types(rewritten_line)
        missing_types = orig_tag_types - rewritten_tag_types
        
        for missing in sorted(missing_types):
            errors.append((line_num, "缺少格式类型", f"缺少格式类型: {missing}"))
    
    return errors


def call_llm_api(prompt: str, filename: str, log_dir: str = None) -> str:
    """Send request to Moonshot API and return the response content. 每次调用时从 config 读取最新 api 配置。"""
    import json
    from datetime import datetime

    cfg = load_config()
    api_cfg = cfg.get("api", {})
    api_key = api_cfg.get("api_key") or API_KEY
    model = api_cfg.get("model") or MODEL
    base_url = api_cfg.get("base_url") or BASE_URL

    headers = {
        "Authorization": f"Bearer {api_key}",
        "Content-Type": "application/json",
    }

    data = {
        "model": model,
        "messages": [{"role": "user", "content": prompt}],
        "max_tokens": 1024 * 32,
    }

    response = httpx.post(
        f"{base_url.rstrip('/')}/chat/completions",
        headers=headers,
        json=data,
        timeout=240,
    )

    # 输出详细的错误信息
    if response.status_code != 200:
        print(f"  [DEBUG] Request data: {data}")
        print(f"  [DEBUG] Response status: {response.status_code}")
        print(f"  [DEBUG] Response body: {response.text}")

    response.raise_for_status()

    result = response.json()

    # 检查 API 响应
    if not result.get("choices"):
        raise ValueError(f"API 返回空响应，filename={filename}")

    choice = result["choices"][0]
    if choice["message"]["content"] is None:
        finish_reason = choice.get("finish_reason", "unknown")
        raise ValueError(f"API 返回 content 为 None，finish_reason={finish_reason}，filename={filename}")

    content = choice["message"]["content"]

    # 保存请求日志
    if log_dir:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            log_filename = f"{filename}_{timestamp}.json"
            log_path = Path(log_dir) / log_filename

            log_data = {
                "timestamp": timestamp,
                "filename": filename,
                "request": {
                    "model": model,
                    "max_tokens": data["max_tokens"],
                    "prompt": prompt,
                },
                "response": {
                    "status_code": response.status_code,
                    "content": content,
                    "finish_reason": choice.get("finish_reason"),
                    "usage": result.get("usage", {}),
                }
            }

            with open(log_path, 'w', encoding='utf-8') as f:
                json.dump(log_data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            print(f"  ! 日志保存失败: {e}")

    return content


def rewrite_chunk(filename: str, content: str, requirement: str, system_prompt: str, log_dir: str = None) -> str:
    """Call Moonshot API to rewrite a chunk."""
    # 将 system_prompt 和 user_prompt 合并到 user message 中
    prompt_parts = [
        system_prompt,
        "",
        "## 改写需求",
        requirement,
        "",
        "## 原始文档",
        content,
    ]
    prompt = "\n".join(prompt_parts)

    return call_llm_api(prompt, filename, log_dir)


def fix_chunk_with_llm(
    filename: str,
    original_content: str,
    rewritten_content: str,
    errors: List[Tuple[int, str, str]],
    fix_prompt: str,
    log_dir: str = None
) -> str:
    """调用 LLM 修复校验错误"""
    error_details = []
    for line_num, error_type, error_msg in errors:
        error_details.append(f"- 第{line_num}行 [{error_type}]: {error_msg}")

    prompt_parts = [
        fix_prompt,
        "",
        "## 错误信息",
        "\n".join(error_details),
        "",
        "## 原文内容",
        original_content,
        "",
        "## 需要修复的内容",
        rewritten_content,
    ]
    prompt = "\n".join(prompt_parts)

    return call_llm_api(prompt, filename, log_dir)


def process_chunk(args: tuple[str, str, str, str, str, str, str]) -> tuple[str, str | None]:
    """Process a single chunk file.

    Args:
        chunks_dir: Source chunks directory
        chunks_new_dir: Destination chunks directory
        filename: File name to process
        requirement: Rewrite requirement
        system_prompt: System prompt for rewriting
        fix_prompt: System prompt for fixing errors
        log_dir: Directory for saving request logs

    Returns:
        (filename, error_message or None)
    """
    chunks_dir, chunks_new_dir, filename, requirement, system_prompt, fix_prompt, log_dir = args

    try:
        # 读取原始文件
        src_path = os.path.join(chunks_dir, filename)
        content = read_file(src_path)
        original_lines = content.strip().split('\n')

        # 调用 Moonshot API 重写
        rewritten = rewrite_chunk(filename, content, requirement, system_prompt, log_dir)

        # 校验并重试（最多10次）
        max_retries = 10
        for attempt in range(max_retries):
            rewritten_lines = rewritten.strip().split('\n')
            errors = validate_rewritten_content(original_lines, rewritten_lines, filename)
            
            if not errors:
                print(f"  ✓ {filename}: 校验通过")
                break  # 校验通过
            
            if attempt < max_retries - 1:
                print(f"  ! {filename}: 校验失败，尝试修复 (第{attempt + 1}次)")
                for line_num, error_type, error_msg in errors[:3]:  # 只显示前3个错误
                    print(f"    - 第{line_num}行 [{error_type}]: {error_msg}")
                if len(errors) > 3:
                    print(f"    ... 还有 {len(errors) - 3} 个错误")
                
                # 调用 LLM 修复
                rewritten = fix_chunk_with_llm(
                    filename, content, rewritten, errors, fix_prompt, log_dir
                )
            else:
                # 最后一次尝试仍然失败
                error_details = "\n".join([
                    f"第{line_num}行 [{error_type}]: {error_msg}"
                    for line_num, error_type, error_msg in errors
                ])
                raise ValueError(f"校验失败（已重试{max_retries}次）:\n{error_details}")

        # 写入新文件
        dst_path = os.path.join(chunks_new_dir, filename)
        write_file(dst_path, rewritten)

        print(f"✓ 完成: {filename}")
        return filename, None

    except Exception as e:
        error_msg = f"✗ 失败: {filename} - {str(e)}"
        print(error_msg)
        return filename, str(e)


def run_rewrite(chunks_dir: str, chunks_new_dir: str, requirement: str, log_dir: str = None, concurrency: int = None):
    """运行改写流程（供外部调用）。concurrency 为 None 时从 config 读取，超过并发数时由线程池排队等待下一轮。"""
    # 确保输出目录存在
    os.makedirs(chunks_new_dir, exist_ok=True)

    # 创建日志目录
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)

    # 获取 chunks 目录下的所有 .md 文件
    chunks_path = Path(chunks_dir)
    files = sorted([f.name for f in chunks_path.glob("*.md")])

    if not files:
        print(f"警告: {chunks_dir} 目录下没有找到 .md 文件")
        return

    if concurrency is None:
        concurrency = _config.get("concurrency", 3)
    concurrency = max(1, int(concurrency))

    print(f"找到 {len(files)} 个 chunk 文件，并发数={concurrency}，开始处理...")

    # 加载提示词
    system_prompt = load_system_prompt()
    fix_prompt = load_fix_prompt()

    # 多线程处理，限制并发数
    results = []
    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        futures = {}
        for filename in files:
            task = (chunks_dir, chunks_new_dir, filename, requirement, system_prompt, fix_prompt, log_dir)
            future = executor.submit(process_chunk, task)
            futures[future] = filename

        for future in as_completed(futures):
            result = future.result()
            results.append(result)

    # 统计结果
    errors = [r for r in results if r[1] is not None]
    success = len(results) - len(errors)

    print(f"\n处理完成: 成功 {success}/{len(files)}")

    if errors:
        print("\n失败的文件:")
        for filename, error in errors:
            print(f"  - {filename}: {error}")
        raise RuntimeError("部分文件改写失败")


def main():
    if len(sys.argv) < 4 or len(sys.argv) > 5:
        print(f"用法: {sys.argv[0]} <chunks_dir> <chunks_new_dir> <requirement> [log_dir]")
        print(f"示例: {sys.argv[0]} output/chunks output/chunks_new \"请改写\" output/logs")
        sys.exit(1)

    chunks_dir = sys.argv[1]
    chunks_new_dir = sys.argv[2]
    requirement = sys.argv[3]
    log_dir = sys.argv[4] if len(sys.argv) > 4 else None

    try:
        run_rewrite(chunks_dir, chunks_new_dir, requirement, log_dir)
    except RuntimeError as e:
        print(f"错误: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
