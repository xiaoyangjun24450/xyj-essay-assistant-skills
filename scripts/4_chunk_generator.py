#!/usr/bin/env python3
"""Generate rewrite chunks from full_paragraphs.txt and template_segments.json."""

from __future__ import annotations

import json
import re
import sys
from dataclasses import dataclass
from pathlib import Path
from typing import Any, List

from pipeline_utils import normalize_hex_para_id


SEGMENT_REQUIRED_KEYS = ("segment_id", "summary", "purpose", "para_ids")


@dataclass(frozen=True)
class Paragraph:
    para_id: str
    text: str

    @property
    def char_count(self) -> int:
        return len(self.text)

    def render(self) -> str:
        return f"[{self.para_id}] {self.text}"


class ChunkGenerator:
    """Generate chunk files while keeping paragraph boundaries inside each segment."""

    def __init__(
        self,
        full_paragraphs_path: str,
        segments_path: str,
        output_dir: str,
        metadata_path: str | None = None,
        max_chars: int = 3000,
    ):
        self.full_paragraphs_path = Path(full_paragraphs_path).expanduser().resolve()
        self.segments_path = Path(segments_path).expanduser().resolve()
        self.output_dir = Path(output_dir).expanduser().resolve()
        if metadata_path is None:
            metadata_path = str(self.output_dir.parent / "chunks_metadata.json")
        self.metadata_path = Path(metadata_path).expanduser().resolve()
        self.max_chars = int(max_chars)

        if self.max_chars <= 0:
            raise ValueError("max_chars 必须大于 0")

        self.paragraphs = self._load_full_paragraphs()
        self.paragraph_map = {para.para_id: para for para in self.paragraphs}
        self.para_index = {para.para_id: index for index, para in enumerate(self.paragraphs)}

    def generate(self) -> dict[str, Any]:
        segments = self._load_segments()
        normalized_segments = self._validate_segments(segments)

        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.metadata_path.parent.mkdir(parents=True, exist_ok=True)
        self._cleanup_existing_chunks()

        chunks: List[dict[str, Any]] = []
        chunk_counter = 1

        for segment in normalized_segments:
            segment_paragraphs = [self.paragraph_map[para_id] for para_id in segment["para_ids"]]
            sub_chunks = self._split_segment(segment_paragraphs)
            split_count = len(sub_chunks)

            for split_index, chunk_paragraphs in enumerate(sub_chunks, start=1):
                filename = f"chunk_{chunk_counter:04d}.md"
                chunk_path = self.output_dir / filename
                content = "\n".join(paragraph.render() for paragraph in chunk_paragraphs)
                chunk_path.write_text(content, encoding="utf-8")

                para_ids = [paragraph.para_id for paragraph in chunk_paragraphs]
                char_count = sum(paragraph.char_count for paragraph in chunk_paragraphs)
                chunks.append(
                    {
                        "chunk_id": f"chunk_{chunk_counter:04d}",
                        "filename": filename,
                        "segment_id": segment["segment_id"],
                        "segment_summary": segment["summary"],
                        "segment_purpose": segment["purpose"],
                        "para_ids": para_ids,
                        "paragraph_count": len(para_ids),
                        "char_count": char_count,
                        "max_chars": self.max_chars,
                        "exceeds_max_chars": char_count > self.max_chars,
                        "is_split_chunk": split_count > 1,
                        "split_index": split_index,
                        "split_count": split_count,
                    }
                )
                chunk_counter += 1

        metadata = {
            "full_paragraphs_path": str(self.full_paragraphs_path),
            "segments_path": str(self.segments_path),
            "chunks_dir": str(self.output_dir),
            "max_chars": self.max_chars,
            "chunk_count": len(chunks),
            "chunks": chunks,
        }
        self.metadata_path.write_text(
            json.dumps(metadata, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        print(f"已生成 {len(chunks)} 个 chunk: {self.output_dir}")
        print(f"已生成 chunk 元数据: {self.metadata_path}")
        return metadata

    def _load_full_paragraphs(self) -> List[Paragraph]:
        if not self.full_paragraphs_path.exists():
            raise FileNotFoundError(f"full_paragraphs 文件未找到: {self.full_paragraphs_path}")

        paragraphs: List[Paragraph] = []
        seen_para_ids = set()
        for line_no, raw_line in enumerate(
            self.full_paragraphs_path.read_text(encoding="utf-8").splitlines(),
            start=1,
        ):
            match = re.match(r"^\[([0-9A-F]{8})\]\s?(.*)$", raw_line)
            if not match:
                raise ValueError(
                    f"{self.full_paragraphs_path} 第 {line_no} 行格式错误，应为 [PARA_ID] 段落文本"
                )

            para_id = match.group(1).upper()
            text = match.group(2)
            if para_id in seen_para_ids:
                raise ValueError(f"full_paragraphs 中存在重复 para_id: {para_id}")
            seen_para_ids.add(para_id)
            paragraphs.append(Paragraph(para_id=para_id, text=text))

        if not paragraphs:
            raise ValueError(f"未在 {self.full_paragraphs_path} 中读取到任何段落")
        return paragraphs

    def _load_segments(self) -> Any:
        if not self.segments_path.exists():
            raise FileNotFoundError(f"segments 文件未找到: {self.segments_path}")
        return json.loads(self.segments_path.read_text(encoding="utf-8"))

    def _validate_segments(self, segments: Any) -> List[dict[str, Any]]:
        if not isinstance(segments, list) or not segments:
            raise ValueError("segments 必须是非空 JSON 数组")

        expected_para_ids = [paragraph.para_id for paragraph in self.paragraphs]
        collected_para_ids: List[str] = []
        normalized_segments: List[dict[str, Any]] = []
        seen_segment_ids = set()

        for index, segment in enumerate(segments, start=1):
            if not isinstance(segment, dict):
                raise ValueError(f"第 {index} 个 segment 不是 JSON 对象")

            missing_keys = [key for key in SEGMENT_REQUIRED_KEYS if key not in segment]
            if missing_keys:
                raise ValueError(
                    f"第 {index} 个 segment 缺少字段: {', '.join(missing_keys)}"
                )

            segment_id = str(segment["segment_id"]).strip()
            summary = str(segment["summary"]).strip()
            purpose = str(segment["purpose"]).strip()
            raw_para_ids = segment["para_ids"]

            if not segment_id:
                raise ValueError(f"第 {index} 个 segment 的 segment_id 不能为空")
            if segment_id in seen_segment_ids:
                raise ValueError(f"segments 中存在重复 segment_id: {segment_id}")
            seen_segment_ids.add(segment_id)

            if not summary:
                raise ValueError(f"segment {segment_id} 的 summary 不能为空")
            if not purpose:
                raise ValueError(f"segment {segment_id} 的 purpose 不能为空")
            if not isinstance(raw_para_ids, list) or not raw_para_ids:
                raise ValueError(f"segment {segment_id} 的 para_ids 必须是非空数组")

            para_ids = [normalize_hex_para_id(str(para_id)) for para_id in raw_para_ids]
            if len(set(para_ids)) != len(para_ids):
                raise ValueError(f"segment {segment_id} 内存在重复 para_id")

            unknown_para_ids = [para_id for para_id in para_ids if para_id not in self.para_index]
            if unknown_para_ids:
                raise ValueError(
                    f"segment {segment_id} 包含未知 para_id: {', '.join(unknown_para_ids)}"
                )

            indexes = [self.para_index[para_id] for para_id in para_ids]
            expected_indexes = list(range(indexes[0], indexes[0] + len(indexes)))
            if indexes != expected_indexes:
                raise ValueError(
                    f"segment {segment_id} 的 para_ids 不是连续段落: {para_ids}"
                )

            normalized_segments.append(
                {
                    "segment_id": segment_id,
                    "summary": summary,
                    "purpose": purpose,
                    "para_ids": para_ids,
                }
            )
            collected_para_ids.extend(para_ids)

        if len(set(collected_para_ids)) != len(collected_para_ids):
            duplicates = self._find_duplicates(collected_para_ids)
            raise ValueError(f"segments 中存在重复 para_id: {', '.join(duplicates)}")

        if collected_para_ids != expected_para_ids:
            missing_para_ids = [para_id for para_id in expected_para_ids if para_id not in collected_para_ids]
            if missing_para_ids:
                raise ValueError(f"segments 缺少 para_id: {', '.join(missing_para_ids)}")

            mismatch = self._find_first_mismatch(expected_para_ids, collected_para_ids)
            if mismatch is not None:
                position, expected_para_id, actual_para_id = mismatch
                raise ValueError(
                    "segments 的 para_id 顺序与原文不一致，"
                    f"第 {position} 个位置期望 {expected_para_id}，实际为 {actual_para_id}"
                )

        return normalized_segments

    def _split_segment(self, paragraphs: List[Paragraph]) -> List[List[Paragraph]]:
        chunks: List[List[Paragraph]] = []
        current_chunk: List[Paragraph] = []
        current_chars = 0

        for paragraph in paragraphs:
            next_chars = current_chars + paragraph.char_count
            if current_chunk and next_chars > self.max_chars:
                chunks.append(current_chunk)
                current_chunk = [paragraph]
                current_chars = paragraph.char_count
            else:
                current_chunk.append(paragraph)
                current_chars = next_chars

        if current_chunk:
            chunks.append(current_chunk)
        return chunks

    def _cleanup_existing_chunks(self) -> None:
        for chunk_path in self.output_dir.glob("chunk_*.md"):
            chunk_path.unlink()

    def _find_duplicates(self, para_ids: List[str]) -> List[str]:
        seen = set()
        duplicates: List[str] = []
        for para_id in para_ids:
            if para_id in seen and para_id not in duplicates:
                duplicates.append(para_id)
            seen.add(para_id)
        return duplicates

    def _find_first_mismatch(
        self,
        expected_para_ids: List[str],
        actual_para_ids: List[str],
    ) -> tuple[int, str, str] | None:
        compare_length = min(len(expected_para_ids), len(actual_para_ids))
        for index in range(compare_length):
            if expected_para_ids[index] != actual_para_ids[index]:
                return index + 1, expected_para_ids[index], actual_para_ids[index]
        if len(expected_para_ids) != len(actual_para_ids):
            expected_para_id = expected_para_ids[compare_length] if compare_length < len(expected_para_ids) else "<结束>"
            actual_para_id = actual_para_ids[compare_length] if compare_length < len(actual_para_ids) else "<结束>"
            return compare_length + 1, expected_para_id, actual_para_id
        return None


def main() -> None:
    
    full_paragraphs_path = "tests/阶段一输出/full_paragraphs.txt"
    segments_path = "tests/阶段三输出/template_segments.json"
    output_dir = "tests/阶段四输出/output_chunks"
    max_chars = 8000
    metadata_path = "tests/阶段四输出/chunks_metadata.json"

    generator = ChunkGenerator(
        full_paragraphs_path=full_paragraphs_path,
        segments_path=segments_path,
        output_dir=output_dir,
        metadata_path=metadata_path,
        max_chars=max_chars,
    )
    generator.generate()


if __name__ == "__main__":
    main()
