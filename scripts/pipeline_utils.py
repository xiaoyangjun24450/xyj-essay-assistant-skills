#!/usr/bin/env python3
"""Shared utilities for segmented RAG pipeline scripts."""

import json
import re
import sys
from datetime import datetime
from functools import lru_cache
from pathlib import Path
from typing import List

from openai import OpenAI


def get_runtime_base_dir() -> Path:
    """Return the directory that contains runtime assets such as config/prompts."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def load_config() -> dict:
    """Load project config from the runtime base directory."""
    config_path = get_runtime_base_dir() / "config.json"
    if not config_path.exists():
        raise FileNotFoundError(f"配置文件未找到: {config_path}")
    return json.loads(config_path.read_text(encoding="utf-8"))


def read_file(path: str) -> str:
    """Read a UTF-8 text file."""
    fp = Path(path).expanduser().resolve()
    if not fp.exists():
        raise FileNotFoundError(f"File not found: {path}")
    if not fp.is_file():
        raise ValueError(f"Not a file: {path}")
    return fp.read_text(encoding="utf-8")


def write_file(path: str, content: str) -> None:
    """Write UTF-8 text content to a file."""
    fp = Path(path).expanduser().resolve()
    fp.parent.mkdir(parents=True, exist_ok=True)
    fp.write_text(content, encoding="utf-8")


def load_prompt(name: str) -> str:
    """Load a prompt file from prompts/<name>.txt."""
    prompt_path = get_runtime_base_dir() / "prompts" / f"{name}.txt"
    if not prompt_path.exists():
        raise FileNotFoundError(f"提示词文件未找到: {prompt_path}")
    return prompt_path.read_text(encoding="utf-8").strip()


def normalize_hex_para_id(para_id: str) -> str:
    """Normalize a template PARA_ID to 8 uppercase hex chars.

    LLMs sometimes drop the leading zero of values like 03922BB3 and return 3922BB3.
    This helper preserves valid 8-char ids and left-pads 1-7 char hex ids.
    """
    normalized = str(para_id).strip().upper()
    if not normalized:
        raise ValueError("para_id 不能为空")
    if re.fullmatch(r"[0-9A-F]{1,8}", normalized):
        return normalized.zfill(8)
    return normalized


@lru_cache(maxsize=8)
def _get_openai_client(api_key: str, base_url: str) -> OpenAI:
    """Create and cache an OpenAI-compatible client for the configured provider."""
    return OpenAI(
        api_key=api_key,
        base_url=base_url.rstrip("/"),
        timeout=240,
    )


def _call_llm_with_messages(
    messages: List[dict[str, str]],
    filename: str,
    log_dir: str = None,
) -> str:
    """Send chat messages through the OpenAI SDK to the configured provider."""
    cfg = load_config()
    api_cfg = cfg.get("api", {})
    api_key = api_cfg.get("api_key")
    model = api_cfg.get("model")
    base_url = api_cfg.get("base_url")

    if not api_key or not model or not base_url:
        raise ValueError("config.json 中缺少 api_key、model 或 base_url 配置")

    client = _get_openai_client(api_key, base_url)
    request_data = {
        "model": model,
        "messages": messages,
        "max_tokens": 1024 * 32,
    }

    try:
        completion = client.chat.completions.create(**request_data)
    except Exception:
        print(f"  [DEBUG] Request data: {request_data}")
        raise

    result = completion.model_dump()
    choices = result.get("choices") or []
    if not choices:
        raise ValueError(f"API 返回空响应，filename={filename}")

    choice = choices[0]
    content = choice.get("message", {}).get("content")
    if content is None:
        finish_reason = choice.get("finish_reason", "unknown")
        raise ValueError(
            f"API 返回 content 为 None，finish_reason={finish_reason}，filename={filename}"
        )

    if log_dir:
        try:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]
            log_path = Path(log_dir) / f"{filename}_{timestamp}.json"
            log_path.parent.mkdir(parents=True, exist_ok=True)
            log_data = {
                "timestamp": timestamp,
                "filename": filename,
                "request": {
                    "model": model,
                    "max_tokens": request_data["max_tokens"],
                    "messages": messages,
                    "base_url": base_url,
                },
                "response": {
                    "content": content,
                    "finish_reason": choice.get("finish_reason"),
                    "usage": result.get("usage", {}),
                    "id": result.get("id"),
                    "model": result.get("model"),
                },
            }
            log_path.write_text(
                json.dumps(log_data, ensure_ascii=False, indent=2),
                encoding="utf-8",
            )
        except Exception as exc:
            print(f"  ! 日志保存失败: {exc}")

    return content


def call_llm_messages(
    messages: List[dict[str, str]],
    filename: str,
    log_dir: str = None,
) -> str:
    """Send multiple chat messages to the configured endpoint."""
    if not messages:
        raise ValueError("messages 不能为空")
    return _call_llm_with_messages(messages, filename, log_dir)


def call_llm_api(prompt: str, filename: str, log_dir: str = None) -> str:
    """Send a single user prompt to the configured chat-completions endpoint."""
    if not prompt.strip():
        raise ValueError("prompt 不能为空")
    return _call_llm_with_messages(
        [{"role": "user", "content": prompt}],
        filename,
        log_dir,
    )
