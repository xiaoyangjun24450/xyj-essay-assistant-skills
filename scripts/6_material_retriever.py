#!/usr/bin/env python3
"""Stage-6 material retrieval for rewrite chunks."""

from __future__ import annotations

import json
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, List, Optional

from pipeline_utils import call_llm_messages, load_prompt, read_file


_NUMBERED_MATERIAL_LINE_RE = re.compile(r"^\[(\d+)\]\s?(.*)$")


@dataclass(frozen=True)
class MaterialParagraph:
    para_id: str
    text: str

    def render(self) -> str:
        return f"[{self.para_id}] {self.text}"


class MaterialRetrievalManager:
    """Load retrieval context and resolve related material paragraphs for each chunk."""

    def __init__(
        self,
        chunks_metadata_path: Optional[str] = None,
        material_numbered_path: Optional[str] = None,
        material_segments_path: Optional[str] = None,
        retrieval_limit: int = 3,
        max_llm_attempts: int = 3,
        log_dir: Optional[str] = None,
        llm_caller: Optional[Callable[[List[dict[str, str]], str, Optional[str]], str]] = None,
    ):
        self.enabled = False
        self.log_dir = log_dir
        self.llm_caller = llm_caller or call_llm_messages
        self.retrieval_limit = max(1, int(retrieval_limit))
        self.max_llm_attempts = max(1, int(max_llm_attempts))
        self.chunk_metadata_by_filename: dict[str, dict[str, Any]] = {}
        self.material_paragraphs: dict[str, MaterialParagraph] = {}
        self.material_segments: list[dict[str, Any]] = []
        self.system_prompt = ""
        self.user_prompt_template = ""

        provided_paths = [chunks_metadata_path, material_numbered_path, material_segments_path]
        if not any(provided_paths):
            return
        if not all(provided_paths):
            raise ValueError(
                "启用素材检索时，必须同时提供 chunks_metadata、material_numbered 和 material_segments"
            )

        self.chunk_metadata_by_filename = self._load_chunks_metadata(chunks_metadata_path or "")
        self.material_paragraphs = self._load_material_paragraphs(material_numbered_path or "")
        self.material_segments = self._load_material_segments(material_segments_path or "")
        self.system_prompt = load_prompt("material_retrieval_system_prompt")
        self.user_prompt_template = load_prompt("material_retrieval_user_prompt")
        self.enabled = True

    def prepare_rewrite_context(
        self,
        filename: str,
        chunk_content: str,
        requirement: str,
    ) -> Optional[dict[str, Any]]:
        if not self.enabled:
            return None

        chunk_metadata = self.chunk_metadata_by_filename.get(filename)
        if chunk_metadata is None:
            raise ValueError(f"未在 chunks_metadata.json 中找到 {filename} 的元数据")

        retrieval = self._retrieve_materials(filename, chunk_content, requirement, chunk_metadata)
        material_lines = [
            self.material_paragraphs[para_id].render()
            for para_id in retrieval["material_para_ids"]
        ]
        return {
            "chunk_metadata": chunk_metadata,
            "retrieval": retrieval,
            "material_lines": material_lines,
        }

    def _retrieve_materials(
        self,
        filename: str,
        chunk_content: str,
        requirement: str,
        chunk_metadata: dict[str, Any],
    ) -> dict[str, Any]:
        chunk_payload = {
            "chunk_id": chunk_metadata["chunk_id"],
            "segment_id": chunk_metadata["segment_id"],
            "segment_summary": chunk_metadata["segment_summary"],
            "segment_purpose": chunk_metadata["segment_purpose"],
            "para_ids": chunk_metadata["para_ids"],
            "is_split_chunk": chunk_metadata.get("is_split_chunk", False),
            "split_index": chunk_metadata.get("split_index", 1),
            "split_count": chunk_metadata.get("split_count", 1),
        }
        material_segment_payload = [
            self._build_material_segment_prompt_item(segment)
            for segment in self.material_segments
        ]
        user_prompt = self.user_prompt_template.format(
            retrieval_limit=self.retrieval_limit,
            requirement=requirement.strip(),
            chunk_metadata_json=json.dumps(chunk_payload, ensure_ascii=False, indent=2),
            chunk_content=chunk_content.strip(),
            material_segments_json=json.dumps(material_segment_payload, ensure_ascii=False, indent=2),
        )

        messages = [
            {"role": "system", "content": self.system_prompt.strip()},
            {"role": "user", "content": user_prompt.strip()},
        ]
        base_filename = f"{Path(filename).stem}_retrieval"
        last_error: Optional[ValueError] = None

        for attempt_index in range(1, self.max_llm_attempts + 1):
            llm_filename = (
                base_filename
                if attempt_index == 1
                else f"{base_filename}_retry_{attempt_index:02d}"
            )
            response_text = self.llm_caller(messages, llm_filename, self.log_dir)

            try:
                parsed = self._parse_response_json(response_text)
                return self._validate_retrieval_result(parsed)
            except ValueError as exc:
                last_error = exc
                self._write_json_log(
                    f"{llm_filename}_validation_error.json",
                    {
                        "attempt": attempt_index,
                        "filename": filename,
                        "chunk_metadata": chunk_payload,
                        "response_text": response_text,
                        "error": str(exc),
                    },
                )

        assert last_error is not None
        raise ValueError(
            "LLM 素材检索结果在多次重试后仍未通过校验。"
            f" attempts={self.max_llm_attempts} filename={filename} last_error={last_error}"
        )

    def _build_material_segment_prompt_item(self, segment: dict[str, Any]) -> dict[str, Any]:
        para_ids = segment["para_ids"]
        preview_parts = [self.material_paragraphs[para_id].text for para_id in para_ids[:2]]
        preview = " ".join(part.strip() for part in preview_parts if part.strip())
        if len(preview) > 160:
            preview = f"{preview[:157]}..."
        return {
            "segment_id": segment["segment_id"],
            "summary": segment["summary"],
            "purpose": segment["purpose"],
            "para_ids": para_ids,
            "preview": preview,
        }

    def _load_chunks_metadata(self, path: str) -> dict[str, dict[str, Any]]:
        payload = json.loads(read_file(path))
        chunks = payload.get("chunks")
        if not isinstance(chunks, list) or not chunks:
            raise ValueError("chunks_metadata.json 缺少非空 chunks 数组")

        metadata_by_filename: dict[str, dict[str, Any]] = {}
        for index, chunk in enumerate(chunks, start=1):
            if not isinstance(chunk, dict):
                raise ValueError(f"chunks_metadata.json 第 {index} 个 chunk 不是对象")

            filename = str(chunk.get("filename", "")).strip()
            chunk_id = str(chunk.get("chunk_id", "")).strip() or Path(filename).stem
            segment_id = str(chunk.get("segment_id", "")).strip()
            segment_summary = str(chunk.get("segment_summary", "")).strip()
            segment_purpose = str(chunk.get("segment_purpose", "")).strip()
            raw_para_ids = chunk.get("para_ids")

            if not filename:
                raise ValueError(f"chunks_metadata.json 第 {index} 个 chunk 缺少 filename")
            if filename in metadata_by_filename:
                raise ValueError(f"chunks_metadata.json 中存在重复 filename: {filename}")
            if not segment_id or not segment_summary or not segment_purpose:
                raise ValueError(f"chunk {filename} 缺少 segment 相关信息")
            if not isinstance(raw_para_ids, list) or not raw_para_ids:
                raise ValueError(f"chunk {filename} 的 para_ids 必须是非空数组")

            metadata_by_filename[filename] = {
                "filename": filename,
                "chunk_id": chunk_id,
                "segment_id": segment_id,
                "segment_summary": segment_summary,
                "segment_purpose": segment_purpose,
                "para_ids": [str(para_id).strip().upper() for para_id in raw_para_ids],
                "is_split_chunk": bool(chunk.get("is_split_chunk", False)),
                "split_index": int(chunk.get("split_index", 1)),
                "split_count": int(chunk.get("split_count", 1)),
            }
        return metadata_by_filename

    def _load_material_paragraphs(self, path: str) -> dict[str, MaterialParagraph]:
        paragraphs: dict[str, MaterialParagraph] = {}
        for line_no, raw_line in enumerate(read_file(path).splitlines(), start=1):
            match = _NUMBERED_MATERIAL_LINE_RE.match(raw_line)
            if not match:
                raise ValueError(
                    f"{path} 第 {line_no} 行格式错误，应为 [数字编号] 段落文本"
                )

            para_id = str(int(match.group(1)))
            text = match.group(2)
            if para_id in paragraphs:
                raise ValueError(f"{path} 中存在重复素材 para_id: {para_id}")
            paragraphs[para_id] = MaterialParagraph(para_id=para_id, text=text)

        if not paragraphs:
            raise ValueError(f"未在 {path} 中读取到任何素材段落")
        return paragraphs

    def _load_material_segments(self, path: str) -> list[dict[str, Any]]:
        payload = json.loads(read_file(path))
        if not isinstance(payload, list) or not payload:
            raise ValueError("material_segments.json 必须是非空 JSON 数组")

        known_para_ids = set(self.material_paragraphs)
        normalized_segments: list[dict[str, Any]] = []
        seen_segment_ids = set()
        for index, segment in enumerate(payload, start=1):
            if not isinstance(segment, dict):
                raise ValueError(f"material_segments.json 第 {index} 个 segment 不是对象")

            segment_id = str(segment.get("segment_id", "")).strip()
            summary = str(segment.get("summary", "")).strip()
            purpose = str(segment.get("purpose", "")).strip()
            raw_para_ids = segment.get("para_ids")

            if not segment_id:
                raise ValueError(f"material_segments.json 第 {index} 个 segment_id 为空")
            if segment_id in seen_segment_ids:
                raise ValueError(f"material_segments.json 中存在重复 segment_id: {segment_id}")
            seen_segment_ids.add(segment_id)
            if not summary or not purpose:
                raise ValueError(f"material segment {segment_id} 缺少 summary 或 purpose")
            if not isinstance(raw_para_ids, list) or not raw_para_ids:
                raise ValueError(f"material segment {segment_id} 的 para_ids 必须是非空数组")

            para_ids = [str(int(str(para_id).strip())) for para_id in raw_para_ids]
            unknown_para_ids = [para_id for para_id in para_ids if para_id not in known_para_ids]
            if unknown_para_ids:
                raise ValueError(
                    f"material segment {segment_id} 包含未知 para_id: {', '.join(unknown_para_ids)}"
                )

            normalized_segments.append(
                {
                    "segment_id": segment_id,
                    "summary": summary,
                    "purpose": purpose,
                    "para_ids": para_ids,
                }
            )
        return normalized_segments

    def _parse_response_json(self, response_text: str) -> Any:
        text = response_text.strip()
        candidates = [text]
        candidates.extend(
            item.strip()
            for item in re.findall(r"```(?:json)?\s*(.*?)```", text, flags=re.S)
        )

        object_match = re.search(r"(\{\s*\".*\})", text, flags=re.S)
        if object_match:
            candidates.append(object_match.group(1).strip())

        for candidate in candidates:
            if not candidate:
                continue
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue
        raise ValueError("LLM 返回的素材检索结果不是有效 JSON")

    def _validate_retrieval_result(self, payload: Any) -> dict[str, Any]:
        if not isinstance(payload, dict):
            raise ValueError("素材检索结果必须是 JSON 对象")

        raw_para_ids = payload.get("material_para_ids")
        outline = str(payload.get("outline", "")).strip()
        reason = str(payload.get("reason", "")).strip()

        if not isinstance(raw_para_ids, list):
            raise ValueError("素材检索结果缺少 material_para_ids 数组")

        normalized_para_ids: list[str] = []
        seen = set()
        for raw_para_id in raw_para_ids:
            para_id = str(int(str(raw_para_id).strip()))
            if para_id in seen:
                continue
            seen.add(para_id)
            normalized_para_ids.append(para_id)

        if not normalized_para_ids:
            raise ValueError("素材检索结果未返回任何 material_para_ids")

        unknown_para_ids = [
            para_id for para_id in normalized_para_ids if para_id not in self.material_paragraphs
        ]
        if unknown_para_ids:
            raise ValueError(
                f"素材检索结果包含未知 para_id: {', '.join(unknown_para_ids)}"
            )

        normalized_para_ids = normalized_para_ids[: self.retrieval_limit]

        if not outline:
            raise ValueError("素材检索结果缺少 outline")
        if not reason:
            raise ValueError("素材检索结果缺少 reason")

        return {
            "material_para_ids": normalized_para_ids,
            "outline": outline,
            "reason": reason,
        }

    def _write_json_log(self, name: str, payload: dict[str, Any]) -> None:
        if not self.log_dir:
            return
        log_path = Path(self.log_dir).expanduser().resolve() / name
        log_path.parent.mkdir(parents=True, exist_ok=True)
        log_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
