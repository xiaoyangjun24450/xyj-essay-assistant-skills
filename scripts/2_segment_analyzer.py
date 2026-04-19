#!/usr/bin/env python3
"""Analyze semantic segments from full_paragraphs.txt using sliding windows."""

from __future__ import annotations

import json
import math
import re
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

from pipeline_utils import call_llm_messages, get_runtime_base_dir, normalize_hex_para_id


@dataclass(frozen=True)
class Paragraph:
    para_id: str
    text: str

    @property
    def char_count(self) -> int:
        return len(self.text)

    def render(self) -> str:
        return f"[{self.para_id}] {self.text}"


class SegmentAnalyzer:
    """Semantic segment analyzer for stage-2 template segmentation."""

    def __init__(
        self,
        full_paragraphs_path: str,
        output_path: Optional[str] = None,
        max_window_chars: int = 8000,
        max_llm_attempts: int = 3,
        log_dir: Optional[str] = None,
        runtime_base_dir: Optional[Path] = None,
        llm_caller: Optional[Callable[[list[dict[str, str]], str, Optional[str]], str]] = None,
        system_prompt_name: str = "segment_system_prompt.txt",
        user_prompt_name: str = "segment_user_prompt.txt",
        paragraph_line_regex: str = r"^\[([0-9A-F]{8})\]\s?(.*)$",
        para_id_normalizer: Optional[Callable[[str], str]] = None,
        segment_id_prefix: str = "seg",
    ):
        self.full_paragraphs_path = Path(full_paragraphs_path).expanduser().resolve()
        if output_path is None:
            output_path = str(self.full_paragraphs_path.with_name("template_segments.json"))
        self.output_path = Path(output_path).expanduser().resolve()
        self.max_window_chars = int(max_window_chars)
        self.max_llm_attempts = max(1, int(max_llm_attempts))
        self.log_dir = log_dir
        self.runtime_base_dir = runtime_base_dir or get_runtime_base_dir()
        self.prompts_dir = self.runtime_base_dir / "prompts"
        self.llm_caller = llm_caller or call_llm_messages
        self.system_prompt_name = system_prompt_name
        self.user_prompt_name = user_prompt_name
        self.paragraph_line_regex = re.compile(paragraph_line_regex)
        self.para_id_normalizer = para_id_normalizer or normalize_hex_para_id
        self.segment_id_prefix = segment_id_prefix.strip() or "seg"

        self.paragraphs = self._load_full_paragraphs()
        self.para_index = {para.para_id: idx for idx, para in enumerate(self.paragraphs)}
        self.system_prompt = self._load_prompt_template(self.system_prompt_name)
        self.user_prompt_template = self._load_prompt_template(self.user_prompt_name)

    def analyze(self) -> list[dict[str, Any]]:
        """Run sliding-window segmentation and write the final JSON output."""
        final_segments: list[dict[str, Any]] = []
        pending_segment: Optional[dict[str, Any]] = None
        previous_segment_for_prompt: Optional[dict[str, Any]] = None
        overlap_para_ids: list[str] = []
        cursor = 0
        window_index = 1

        # 按窗口推进；最后一个分段会留到下一轮继续衔接。
        while cursor < len(self.paragraphs) or overlap_para_ids:
            window = self._build_window(overlap_para_ids, cursor, window_index)
            window_segments = self._analyze_single_window(
                window_paragraphs=window["paragraphs"],
                window_index=window_index,
                total_windows=window["estimated_total_windows"],
                previous_segment=previous_segment_for_prompt,
                context_hint=window["context_hint"],
            )

            is_final_window = window["next_cursor"] >= len(self.paragraphs)
            if pending_segment is None:
                if is_final_window:
                    final_segments.extend(window_segments)
                    overlap_para_ids = []
                elif len(window_segments) == 1:
                    pending_segment = window_segments[0]
                    overlap_para_ids = list(pending_segment["para_ids"])
                    previous_segment_for_prompt = deepcopy(pending_segment)
                else:
                    final_segments.extend(window_segments[:-1])
                    pending_segment = window_segments[-1]
                    overlap_para_ids = list(pending_segment["para_ids"])
                    previous_segment_for_prompt = deepcopy(pending_segment)
            else:
                merged_first = self._merge_pending_segment(
                    pending_segment,
                    window_segments[0],
                )
                if is_final_window:
                    final_segments.append(merged_first)
                    final_segments.extend(window_segments[1:])
                    overlap_para_ids = []
                    pending_segment = None
                elif len(window_segments) == 1:
                    pending_segment = merged_first
                    overlap_para_ids = list(pending_segment["para_ids"])
                    previous_segment_for_prompt = deepcopy(pending_segment)
                else:
                    final_segments.append(merged_first)
                    final_segments.extend(window_segments[1:-1])
                    pending_segment = window_segments[-1]
                    overlap_para_ids = list(pending_segment["para_ids"])
                    previous_segment_for_prompt = deepcopy(pending_segment)

            cursor = window["next_cursor"]
            window_index += 1

            if is_final_window:
                break

        normalized_segments = self._renumber_segments(final_segments)
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self.output_path.write_text(
            json.dumps(normalized_segments, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"已生成分段结果: {self.output_path}")
        return normalized_segments

    def _load_full_paragraphs(self) -> list[Paragraph]:
        lines = self.full_paragraphs_path.read_text(encoding="utf-8").splitlines()
        paragraphs: list[Paragraph] = []

        # 每行格式固定为 [PARA_ID] 段落内容。
        for raw_line in lines:
            match = self.paragraph_line_regex.match(raw_line)
            if not match:
                continue

            paragraphs.append(
                Paragraph(
                    para_id=self._normalize_para_id_value(match.group(1)),
                    text=match.group(2),
                )
            )

        return paragraphs

    def _load_prompt_template(self, name: str) -> str:
        return (self.prompts_dir / name).read_text(encoding="utf-8")

    def _build_window(
        self,
        overlap_para_ids: list[str],
        cursor: int,
        window_index: int,
    ) -> dict[str, Any]:
        # 先把上一轮最后一个分段放进来，保证跨窗口可以续上。
        overlap_paragraphs = [self.paragraphs[self.para_index[para_id]] for para_id in overlap_para_ids]
        window_paragraphs = list(overlap_paragraphs)
        char_count = sum(para.char_count for para in window_paragraphs)
        next_cursor = cursor
        added_new_paragraph = False

        while next_cursor < len(self.paragraphs):
            para = self.paragraphs[next_cursor]
            next_char_count = char_count + para.char_count
            if window_paragraphs and next_char_count > self.max_window_chars and added_new_paragraph:
                break

            window_paragraphs.append(para)
            char_count = next_char_count
            next_cursor += 1
            added_new_paragraph = True

            if next_char_count > self.max_window_chars:
                break

        start_index = self.para_index[window_paragraphs[0].para_id]
        return {
            "paragraphs": window_paragraphs,
            "next_cursor": next_cursor,
            "context_hint": self._build_context_hint(start_index, next_cursor),
            "estimated_total_windows": self._estimate_total_windows(char_count, cursor, window_index),
        }

    def _estimate_total_windows(
        self,
        current_window_chars: int,
        cursor: int,
        window_index: int,
    ) -> int:
        remaining_chars = sum(para.char_count for para in self.paragraphs[cursor:])
        effective_window_chars = max(1, min(self.max_window_chars, max(current_window_chars, 1)))
        estimated_remaining_windows = max(1, math.ceil(remaining_chars / effective_window_chars))
        return max(window_index, window_index - 1 + estimated_remaining_windows)

    def _build_context_hint(self, start_index: int, next_cursor: int) -> str:
        previous_hint = "无前文"
        next_hint = "无后文"

        if start_index > 0:
            previous_hint = self.paragraphs[start_index - 1].render()
        if next_cursor < len(self.paragraphs):
            next_hint = self.paragraphs[next_cursor].render()

        return f"窗口外前文: {previous_hint}\n窗口外后文: {next_hint}"

    def _analyze_single_window(
        self,
        window_paragraphs: list[Paragraph],
        window_index: int,
        total_windows: int,
        previous_segment: Optional[dict[str, Any]],
        context_hint: str,
    ) -> list[dict[str, Any]]:
        # 把当前窗口和上一轮尾 segment 一起喂给模型。
        window_content = "\n".join(para.render() for para in window_paragraphs)
        previous_segment_json = (
            json.dumps(previous_segment, ensure_ascii=False, indent=2)
            if previous_segment
            else "null"
        )
        prompt = self.user_prompt_template.format(
            window_content=window_content,
            window_index=window_index,
            total_windows=max(window_index, total_windows),
            previous_segment_json=previous_segment_json,
            context_hint=context_hint,
            max_segment_chars=self.max_window_chars,
        )
        messages = [
            {"role": "system", "content": self.system_prompt.strip()},
            {"role": "user", "content": prompt.strip()},
        ]
        base_filename = f"segment_window_{window_index:03d}"
        last_error: Optional[ValueError] = None

        for attempt_index in range(1, self.max_llm_attempts + 1):
            filename = (
                base_filename
                if attempt_index == 1
                else f"{base_filename}_retry_{attempt_index:02d}"
            )
            response_text = self.llm_caller(messages, filename, self.log_dir)

            try:
                return self._normalize_window_segments(self._parse_response_json(response_text))
            except ValueError as exc:
                last_error = exc
                self._write_validation_error_log(
                    f"{filename}_validation_error.json",
                    {
                        "attempt": attempt_index,
                        "window_index": window_index,
                        "total_windows": max(window_index, total_windows),
                        "window_para_ids": [para.para_id for para in window_paragraphs],
                        "response_text": response_text,
                        "error": str(exc),
                    },
                )

        assert last_error is not None
        raise ValueError(
            "LLM 分段结果在多次重试后仍未通过校验。"
            f" attempts={self.max_llm_attempts} window_index={window_index} last_error={last_error}"
        )

    def _parse_response_json(self, response_text: str) -> Any:
        text = response_text.strip()
        candidates = [text]

        fenced = re.findall(r"```(?:json)?\s*(.*?)```", text, flags=re.S)
        candidates.extend(item.strip() for item in fenced if item.strip())

        array_match = re.search(r"(\[\s*\{.*\}\s*\])", text, flags=re.S)
        if array_match:
            candidates.append(array_match.group(1).strip())

        for candidate in candidates:
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue

        raise ValueError("LLM 返回的分段结果不是有效 JSON")

    def _normalize_window_segments(self, segments: Any) -> list[dict[str, Any]]:
        if not isinstance(segments, list):
            raise ValueError("LLM 返回的分段结果不是 JSON 数组")

        normalized_segments: list[dict[str, Any]] = []
        for idx, segment in enumerate(segments, start=1):
            if not isinstance(segment, dict):
                raise ValueError(f"LLM 返回的第 {idx} 个分段不是 JSON 对象")

            raw_para_ids = segment.get("para_ids", [])
            if not isinstance(raw_para_ids, list):
                raise ValueError(f"LLM 返回的第 {idx} 个分段缺少 para_ids 数组")

            normalized_segments.append(
                {
                    "segment_id": str(segment.get("segment_id") or f"window_seg_{idx}").strip(),
                    "summary": str(segment.get("summary", "")).strip(),
                    "purpose": str(segment.get("purpose", "")).strip(),
                    "para_ids": [
                        self._normalize_para_id_value(str(para_id))
                        for para_id in raw_para_ids
                    ],
                }
            )

        return normalized_segments

    def _write_validation_error_log(self, name: str, payload: dict[str, Any]) -> None:
        if not self.log_dir:
            return
        log_path = Path(self.log_dir).expanduser().resolve() / name
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _merge_pending_segment(
        self,
        pending_segment: dict[str, Any],
        current_first_segment: dict[str, Any],
    ) -> dict[str, Any]:
        # 新窗口首段默认视为上一窗口尾段的延续。
        merged = deepcopy(current_first_segment)
        merged["segment_id"] = pending_segment["segment_id"]
        return merged

    def _renumber_segments(self, segments: list[dict[str, Any]]) -> list[dict[str, Any]]:
        # 最终输出统一重排 segment_id，避免保留窗口内临时编号。
        return [
            {
                **deepcopy(segment),
                "segment_id": f"{self.segment_id_prefix}_{idx}",
            }
            for idx, segment in enumerate(segments, start=1)
        ]

    def _normalize_para_id_value(self, para_id: str) -> str:
        return self.para_id_normalizer(str(para_id).strip())


def main() -> None:
    # 这里保留默认测试路径，方便直接运行脚本调试。
    full_paragraphs_path = "tests/阶段一输出/full_paragraphs.txt"
    output_path = "tests/阶段二输出/template_segments.json"
    max_window_chars = 8000
    log_dir = "tests/阶段二输出/logs"

    analyzer = SegmentAnalyzer(
        full_paragraphs_path=full_paragraphs_path,
        output_path=output_path,
        max_window_chars=max_window_chars,
        log_dir=log_dir,
    )
    analyzer.analyze()


if __name__ == "__main__":
    main()
