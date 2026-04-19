#!/usr/bin/env python3
"""Rewrite chunks with optional material retrieval."""

from __future__ import annotations

import importlib
import json
import os
import re
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path
from typing import Any, Callable, List, Optional, Tuple

from pipeline_utils import (
    call_llm_messages,
    load_config,
    load_prompt,
    read_file,
    write_file,
)


_config = load_config()
_HEX_PARA_ID_RE = re.compile(r"^\[([0-9A-Fa-f]{8})\]")
MaterialRetrievalManager = importlib.import_module("6_material_retriever").MaterialRetrievalManager


def is_blank_line(line: str) -> bool:
    """Return True when the line only contains whitespace characters."""
    return line.strip() == ""


def is_empty_para_line(line: str) -> bool:
    """Return True when the line only contains a PARA_ID and blank body."""
    match = _HEX_PARA_ID_RE.match(line)
    if not match:
        return False
    return line[match.end() :].strip() == ""


def should_skip_line(line: str) -> bool:
    """Return True for lines that should be hidden from the LLM and validation."""
    return is_blank_line(line) or is_empty_para_line(line)


def build_compact_line_view(lines: List[str]) -> Tuple[List[str], List[int]]:
    """Return non-skipped lines and their original zero-based indexes."""
    compact_lines: List[str] = []
    line_map: List[int] = []

    for idx, line in enumerate(lines):
        if should_skip_line(line):
            continue
        compact_lines.append(line)
        line_map.append(idx)

    return compact_lines, line_map


def compact_content_for_llm(content: str) -> str:
    """Drop skipped lines before sending content to the LLM."""
    compact_lines, _ = build_compact_line_view(content.splitlines())
    return "\n".join(compact_lines)


def restore_blank_lines(
    original_lines: List[str],
    rewritten_lines: List[str],
    preserve_trailing_newline: bool = False,
) -> str:
    """Reinsert rewritten non-blank lines into the original blank-line skeleton."""
    original_compact_lines, original_line_map = build_compact_line_view(original_lines)
    rewritten_compact_lines, _ = build_compact_line_view(rewritten_lines)

    if len(rewritten_compact_lines) != len(original_compact_lines):
        raise ValueError(
            "恢复空白行失败: "
            f"原文非空白行={len(original_compact_lines)}, "
            f"改写后非空白行={len(rewritten_compact_lines)}"
        )

    restored_lines = list(original_lines)
    for compact_idx, original_line_idx in enumerate(original_line_map):
        restored_lines[original_line_idx] = rewritten_compact_lines[compact_idx]

    restored_content = "\n".join(restored_lines)
    if preserve_trailing_newline:
        restored_content += "\n"
    return restored_content


def extract_para_ids(content: str) -> list[str]:
    para_ids: list[str] = []
    for raw_line in content.splitlines():
        match = _HEX_PARA_ID_RE.match(raw_line)
        if match:
            para_ids.append(match.group(1).upper())
    return para_ids


def extract_para_id(line: str) -> Optional[str]:
    """Extract PARA_ID in the form of [XXXXXXXX] from a line."""
    match = _HEX_PARA_ID_RE.match(line)
    if match:
        return match.group(1).upper()
    return None


def extract_tag_types(line: str) -> set[str]:
    """Extract format tag attribute types from a line."""
    tag_types = set()
    tag_pair_pattern = r"<([fisbzchuv,=\w]+)>([^<]*)</\1>"

    for match in re.finditer(tag_pair_pattern, line):
        attrs = match.group(1)
        for attr in attrs.split(","):
            if "=" in attr:
                tag_types.add(attr.strip())

    return tag_types


def verify_format_tags(line: str) -> Tuple[bool, Optional[str]]:
    """Verify whether markup tags in a line are balanced and properly nested."""
    tag_pattern = r"<([fisbzchuv,=\w]*)>|</([fisbzchuv,=\w]*)>"
    tags = list(re.finditer(tag_pattern, line))

    if not tags:
        return True, None

    stack = []
    for match in tags:
        open_tag = match.group(1)
        close_tag = match.group(2)

        if open_tag is not None:
            stack.append(open_tag)
        elif close_tag is not None:
            if not stack:
                return False, f"多余的闭标签 </{close_tag}>"

            expected_open = stack.pop()
            if expected_open != close_tag:
                return False, f"标签不匹配: <{expected_open}> 与 </{close_tag}>"

    if stack:
        return False, f"标签未闭合: <{stack[0]}>"

    return True, None


def validate_rewritten_content(
    original_lines: List[str],
    rewritten_lines: List[str],
    filename: str,
) -> List[Tuple[int, str, str]]:
    """Validate line count, PARA_ID, and format-tag consistency."""
    del filename
    errors = []
    original_compact_lines, original_line_map = build_compact_line_view(original_lines)
    rewritten_compact_lines, _ = build_compact_line_view(rewritten_lines)

    if len(rewritten_compact_lines) != len(original_compact_lines):
        errors.append(
            (
                0,
                "行数不匹配",
                (
                    "非空白行数不一致: "
                    f"原文={len(original_compact_lines)}, "
                    f"改写后={len(rewritten_compact_lines)}"
                ),
            )
        )
        return errors

    for orig_line, rewritten_line, original_idx in zip(
        original_compact_lines,
        rewritten_compact_lines,
        original_line_map,
    ):
        line_num = original_idx + 1

        orig_para_id = extract_para_id(orig_line)
        rewritten_para_id = extract_para_id(rewritten_line)

        if orig_para_id is None:
            continue

        if rewritten_para_id is None:
            errors.append((line_num, "缺少PARA_ID", f"第{line_num}行缺少PARA_ID"))
        elif orig_para_id != rewritten_para_id:
            errors.append(
                (
                    line_num,
                    "PARA_ID不匹配",
                    f"PARA_ID不匹配: 原文={orig_para_id}, 改写后={rewritten_para_id}",
                )
            )

        passed, error_msg = verify_format_tags(rewritten_line)
        if not passed:
            errors.append((line_num, "格式标签错误", error_msg or "未知标签错误"))

        orig_tag_types = extract_tag_types(orig_line)
        rewritten_tag_types = extract_tag_types(rewritten_line)
        missing_types = orig_tag_types - rewritten_tag_types

        for missing in sorted(missing_types):
            errors.append((line_num, "缺少格式类型", f"缺少格式类型: {missing}"))

    return errors


def rewrite_chunk(
    filename: str,
    content: str,
    requirement: str,
    system_prompt: str,
    log_dir: Optional[str] = None,
    rewrite_context: Optional[dict[str, Any]] = None,
    llm_caller: Optional[Callable[[List[dict[str, str]], str, Optional[str]], str]] = None,
) -> str:
    """Call LLM to rewrite a chunk."""
    llm = llm_caller or call_llm_messages
    compact_content = compact_content_for_llm(content)
    prompt_parts = [
        "## 改写需求",
        requirement.strip(),
    ]

    if rewrite_context:
        chunk_metadata = rewrite_context["chunk_metadata"]
        retrieval = rewrite_context["retrieval"]
        material_lines = rewrite_context["material_lines"]

        prompt_parts.extend(
            [
                "",
                "## 当前 chunk 的语义信息",
                json.dumps(
                    {
                        "chunk_id": chunk_metadata["chunk_id"],
                        "segment_id": chunk_metadata["segment_id"],
                        "segment_summary": chunk_metadata["segment_summary"],
                        "segment_purpose": chunk_metadata["segment_purpose"],
                        "para_ids": chunk_metadata["para_ids"],
                        "is_split_chunk": chunk_metadata.get("is_split_chunk", False),
                        "split_index": chunk_metadata.get("split_index", 1),
                        "split_count": chunk_metadata.get("split_count", 1),
                    },
                    ensure_ascii=False,
                    indent=2,
                ),
                "",
                "## 改写大纲",
                retrieval["outline"],
                "",
                "## 检索到的相关素材段落",
                "\n".join(material_lines),
                "",
                "## 素材选择理由",
                retrieval["reason"],
            ]
        )

    prompt_parts.extend(
        [
            "",
            "## 原始文档",
            compact_content,
        ]
    )

    return llm(
        [
            {"role": "system", "content": system_prompt.strip()},
            {"role": "user", "content": "\n".join(prompt_parts).strip()},
        ],
        filename,
        log_dir,
    )


def fix_chunk_with_llm(
    filename: str,
    original_content: str,
    rewritten_content: str,
    errors: List[Tuple[int, str, str]],
    fix_prompt: str,
    log_dir: Optional[str] = None,
    llm_caller: Optional[Callable[[List[dict[str, str]], str, Optional[str]], str]] = None,
) -> str:
    """Call LLM to fix validation errors."""
    llm = llm_caller or call_llm_messages
    compact_original_content = compact_content_for_llm(original_content)
    compact_rewritten_content = compact_content_for_llm(rewritten_content)
    error_details = []
    for line_num, error_type, error_msg in errors:
        error_details.append(f"- 第{line_num}行 [{error_type}]: {error_msg}")

    prompt_parts = [
        "## 错误信息",
        "\n".join(error_details),
        "",
        "## 原文内容",
        compact_original_content,
        "",
        "## 需要修复的内容",
        compact_rewritten_content,
    ]

    return llm(
        [
            {"role": "system", "content": fix_prompt.strip()},
            {"role": "user", "content": "\n".join(prompt_parts).strip()},
        ],
        filename,
        log_dir,
    )


def process_chunk(
    args: tuple[
        str,
        str,
        str,
        str,
        str,
        str,
        Optional[str],
        MaterialRetrievalManager,
    ],
) -> tuple[str, str | None, dict[str, Any] | None]:
    """Process a single chunk file."""
    (
        chunks_dir,
        chunks_new_dir,
        filename,
        requirement,
        system_prompt,
        fix_prompt,
        log_dir,
        retrieval_manager,
    ) = args

    try:
        src_path = os.path.join(chunks_dir, filename)
        content = read_file(src_path)
        original_lines = content.splitlines()
        preserve_trailing_newline = content.endswith("\n")
        original_compact_lines, _ = build_compact_line_view(original_lines)

        dst_path = os.path.join(chunks_new_dir, filename)
        if not original_compact_lines:
            write_file(dst_path, content)
            print(f"  ✓ {filename}: 全为空白行，跳过 LLM")
            print(f"✓ 完成: {filename}")
            return filename, None, None

        rewrite_context = retrieval_manager.prepare_rewrite_context(filename, content, requirement)
        rewritten = rewrite_chunk(
            filename,
            content,
            requirement,
            system_prompt,
            log_dir,
            rewrite_context=rewrite_context,
        )

        max_retries = 10
        restored_rewritten = ""
        for attempt in range(max_retries):
            rewritten_lines = rewritten.splitlines()
            errors = validate_rewritten_content(original_lines, rewritten_lines, filename)

            if not errors:
                restored_rewritten = restore_blank_lines(
                    original_lines,
                    rewritten_lines,
                    preserve_trailing_newline=preserve_trailing_newline,
                )
                print(f"  ✓ {filename}: 校验通过")
                break

            if attempt < max_retries - 1:
                print(f"  ! {filename}: 校验失败，尝试修复 (第{attempt + 1}次)")
                for line_num, error_type, error_msg in errors[:3]:
                    print(f"    - 第{line_num}行 [{error_type}]: {error_msg}")
                if len(errors) > 3:
                    print(f"    ... 还有 {len(errors) - 3} 个错误")

                rewritten = fix_chunk_with_llm(
                    filename,
                    content,
                    rewritten,
                    errors,
                    fix_prompt,
                    log_dir,
                )
            else:
                error_details = "\n".join(
                    f"第{line_num}行 [{error_type}]: {error_msg}"
                    for line_num, error_type, error_msg in errors
                )
                raise ValueError(f"校验失败（已重试{max_retries}次）:\n{error_details}")

        write_file(dst_path, restored_rewritten)

        retrieval_payload = None
        if rewrite_context:
            retrieval_payload = {
                "filename": filename,
                "chunk_id": rewrite_context["chunk_metadata"]["chunk_id"],
                "segment_id": rewrite_context["chunk_metadata"]["segment_id"],
                "chunk_para_ids": extract_para_ids(content),
                "material_para_ids": rewrite_context["retrieval"]["material_para_ids"],
                "outline": rewrite_context["retrieval"]["outline"],
                "reason": rewrite_context["retrieval"]["reason"],
            }

        print(f"✓ 完成: {filename}")
        return filename, None, retrieval_payload

    except Exception as exc:
        error_msg = f"✗ 失败: {filename} - {str(exc)}"
        print(error_msg)
        return filename, str(exc), None


def run_rewrite(
    chunks_dir: str,
    chunks_new_dir: str,
    requirement: str,
    log_dir: Optional[str] = None,
    concurrency: Optional[int] = None,
    chunks_metadata_path: Optional[str] = None,
    material_numbered_path: Optional[str] = None,
    material_segments_path: Optional[str] = None,
    retrieval_limit: int = 3,
):
    """Run rewrite pipeline with optional material retrieval."""
    os.makedirs(chunks_new_dir, exist_ok=True)
    if log_dir:
        os.makedirs(log_dir, exist_ok=True)

    chunks_path = Path(chunks_dir)
    files = sorted(f.name for f in chunks_path.glob("*.md"))
    if not files:
        print(f"警告: {chunks_dir} 目录下没有找到 .md 文件")
        return

    if concurrency is None:
        concurrency = _config.get("concurrency", 3)
    concurrency = max(1, int(concurrency))

    retrieval_manager = MaterialRetrievalManager(
        chunks_metadata_path=chunks_metadata_path,
        material_numbered_path=material_numbered_path,
        material_segments_path=material_segments_path,
        retrieval_limit=retrieval_limit,
        log_dir=log_dir,
        llm_caller=call_llm_messages,
    )

    print(f"找到 {len(files)} 个 chunk 文件，并发数={concurrency}，开始处理...")
    if retrieval_manager.enabled:
        print("已启用素材检索流程")

    system_prompt = load_prompt("system_prompt")
    fix_prompt = load_prompt("fix_prompt")

    results = []
    with ThreadPoolExecutor(max_workers=concurrency) as executor:
        futures = {}
        for filename in files:
            task = (
                chunks_dir,
                chunks_new_dir,
                filename,
                requirement,
                system_prompt,
                fix_prompt,
                log_dir,
                retrieval_manager,
            )
            futures[executor.submit(process_chunk, task)] = filename

        for future in as_completed(futures):
            results.append(future.result())

    errors = [result for result in results if result[1] is not None]
    success = len(results) - len(errors)

    print(f"\n处理完成: 成功 {success}/{len(files)}")

    retrieval_results = [result[2] for result in results if result[2] is not None]
    if retrieval_results:
        retrieval_results_path = Path(chunks_new_dir).expanduser().resolve().parent / "retrieval_results.json"
        retrieval_results_path.write_text(
            json.dumps(
                {
                    "chunks_dir": str(Path(chunks_dir).expanduser().resolve()),
                    "chunks_metadata_path": str(Path(chunks_metadata_path).expanduser().resolve()),
                    "material_numbered_path": str(Path(material_numbered_path).expanduser().resolve()),
                    "material_segments_path": str(Path(material_segments_path).expanduser().resolve()),
                    "retrieval_limit": retrieval_limit,
                    "results": sorted(retrieval_results, key=lambda item: item["filename"]),
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )
        print(f"已生成检索结果: {retrieval_results_path}")

    if errors:
        print("\n失败的文件:")
        for filename, error, _ in errors:
            print(f"  - {filename}: {error}")
        raise RuntimeError("部分文件改写失败")


def main() -> None:
    chunks_dir = "tests/阶段四输出/output_chunks"
    chunks_new_dir = "tests/阶段七输出"
    requirement = "请把内容改为基于三维激光雷达的智能移动充电桩的自动驾驶方法研究"
    log_dir = "tests/阶段七输出/logs"
    concurrency = _config.get("concurrency", 3)
    chunks_metadata_path = "tests/阶段四输出/chunks_metadata.json"
    material_numbered_path = "tests/阶段五输出/material_numbered.txt"
    material_segments_path = "tests/阶段五输出/material_segments.json"
    retrieval_limit = 3

    try:
        run_rewrite(
            chunks_dir,
            chunks_new_dir,
            requirement,
            log_dir,
            concurrency=concurrency,
            chunks_metadata_path=chunks_metadata_path,
            material_numbered_path=material_numbered_path,
            material_segments_path=material_segments_path,
            retrieval_limit=retrieval_limit,
        )
    except RuntimeError as exc:
        print(f"错误: {exc}")
        raise SystemExit(1) from exc


if __name__ == "__main__":
    main()
