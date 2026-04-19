#!/usr/bin/env python3
"""Process user materials through stage-1/2/3 style flow and merge outputs."""

from __future__ import annotations

import argparse
import importlib
import json
import re
from pathlib import Path
from typing import Any, Callable, Optional

from pipeline_utils import call_llm_messages, get_runtime_base_dir


DocxPreprocessor = importlib.import_module("1_docx_preprocessor").DocxPreprocessor
SegmentAnalyzer = importlib.import_module("2_segment_analyzer").SegmentAnalyzer
SegmentValidator = importlib.import_module("3_segment_validator").SegmentValidator

_HEX_PARAGRAPH_RE = re.compile(r"^\[([0-9A-F]{8})\]\s?(.*)$")


class MaterialProcessor:
    """Stage-5 material processor for txt/md/docx inputs."""

    def __init__(
        self,
        material_paths: list[str],
        output_dir: str,
        max_window_chars: int = 8000,
        log_dir: Optional[str] = None,
        runtime_base_dir: Optional[Path] = None,
        llm_caller: Optional[Callable[[list[dict[str, str]], str, Optional[str]], str]] = None,
    ):
        if not material_paths:
            raise ValueError("material_paths 不能为空")

        self.material_paths = [Path(path).expanduser().resolve() for path in material_paths]
        self.output_dir = Path(output_dir).expanduser().resolve()
        self.numbered_output_path = Path(
            self.output_dir / "material_numbered.txt"
        ).expanduser().resolve()
        self.segments_output_path = Path(
            self.output_dir / "material_segments.json"
        ).expanduser().resolve()
        self.max_window_chars = int(max_window_chars)
        self.log_dir = Path(log_dir).expanduser().resolve() if log_dir else None
        self.runtime_base_dir = runtime_base_dir or get_runtime_base_dir()
        self.llm_caller = llm_caller or call_llm_messages

        if self.max_window_chars <= 0:
            raise ValueError("max_window_chars 必须大于 0")

    def process(self) -> dict[str, Any]:
        self.output_dir.mkdir(parents=True, exist_ok=True)
        self.numbered_output_path.parent.mkdir(parents=True, exist_ok=True)
        self.segments_output_path.parent.mkdir(parents=True, exist_ok=True)

        material_results = [
            self._process_single_material(material_path, index)
            for index, material_path in enumerate(self.material_paths, start=1)
        ]
        paragraphs, segments = self._merge_material_results(material_results)
        if not paragraphs:
            raise ValueError("未从素材中读取到任何有效段落")

        self.numbered_output_path.write_text("\n".join(paragraphs), encoding="utf-8")
        self.segments_output_path.write_text(
            json.dumps(segments, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

        print(f"已生成编号素材: {self.numbered_output_path}")
        print(f"已生成素材分段结果: {self.segments_output_path}")

        return {
            "material_paths": [str(path) for path in self.material_paths],
            "output_dir": str(self.output_dir),
            "numbered_output_path": str(self.numbered_output_path),
            "segments_output_path": str(self.segments_output_path),
            "paragraph_count": len(paragraphs),
            "segment_count": len(segments),
            "segments": segments,
        }

    def _process_single_material(self, material_path: Path, index: int) -> dict[str, Any]:
        if not material_path.exists():
            raise FileNotFoundError(f"素材文件未找到: {material_path}")
        if not material_path.is_file():
            raise ValueError(f"素材路径不是文件: {material_path}")

        material_dir = self.output_dir / f"material_{index:02d}"
        full_paragraphs_path = self._prepare_full_paragraphs(material_path, material_dir)
        segments = self._run_segmentation(full_paragraphs_path, material_dir)
        return {
            "material_path": str(material_path),
            "full_paragraphs_path": str(full_paragraphs_path),
            "segments": segments,
        }

    def _prepare_full_paragraphs(self, material_path: Path, material_dir: Path) -> Path:
        suffix = material_path.suffix.lower()
        if suffix == ".docx":
            DocxPreprocessor(str(material_path)).process(str(material_dir))
            return material_dir / "full_paragraphs.txt"
        if suffix not in {".txt", ".md", ".markdown"}:
            raise ValueError(f"不支持的素材格式: {material_path.suffix}，目前仅支持 txt、md、docx")

        paragraphs = self._read_text_like_file(material_path)
        if not paragraphs:
            raise ValueError(f"素材未解析出有效段落: {material_path}")

        output_path = material_dir / "full_paragraphs.txt"
        material_dir.mkdir(parents=True, exist_ok=True)
        output_path.write_text(
            "\n".join(f"[{index:08X}] {paragraph}" for index, paragraph in enumerate(paragraphs)),
            encoding="utf-8",
        )
        return output_path

    def _run_segmentation(self, full_paragraphs_path: Path, material_dir: Path) -> list[dict[str, Any]]:
        stage2_path = material_dir / "material_segments.stage2.json"
        validated_path = material_dir / "material_segments.validated.json"

        SegmentAnalyzer(
            full_paragraphs_path=str(full_paragraphs_path),
            output_path=str(stage2_path),
            max_window_chars=self.max_window_chars,
            log_dir=str(self._resolve_log_dir(material_dir, "stage2")),
            runtime_base_dir=self.runtime_base_dir,
            llm_caller=self.llm_caller,
            system_prompt_name="material_segment_system_prompt.txt",
            user_prompt_name="material_segment_user_prompt.txt",
            segment_id_prefix="mseg",
        ).analyze()

        SegmentValidator(
            full_paragraphs_path=str(full_paragraphs_path),
            input_segments_path=str(stage2_path),
            output_path=str(validated_path),
            log_dir=str(self._resolve_log_dir(material_dir, "stage3")),
            llm_caller=self.llm_caller,
        ).run()

        return json.loads(validated_path.read_text(encoding="utf-8"))

    def _merge_material_results(
        self,
        material_results: list[dict[str, Any]],
    ) -> tuple[list[str], list[dict[str, Any]]]:
        merged_paragraphs: list[str] = []
        merged_segments: list[dict[str, Any]] = []
        next_para_id = 1

        for result in material_results:
            paragraphs = self._load_full_paragraphs(Path(result["full_paragraphs_path"]))
            para_id_map: dict[str, str] = {}

            for paragraph in paragraphs:
                para_id_map[paragraph["para_id"]] = str(next_para_id)
                merged_paragraphs.append(f"[{next_para_id}] {paragraph['text']}")
                next_para_id += 1

            for segment in result["segments"]:
                merged_segments.append(
                    {
                        "segment_id": "",
                        "summary": str(segment.get("summary", "")).strip(),
                        "purpose": str(segment.get("purpose", "")).strip(),
                        "para_ids": [para_id_map[para_id] for para_id in segment.get("para_ids", [])],
                    }
                )

        for index, segment in enumerate(merged_segments, start=1):
            segment["segment_id"] = f"mseg_{index}"

        return merged_paragraphs, merged_segments

    def _load_full_paragraphs(self, path: Path) -> list[dict[str, str]]:
        paragraphs: list[dict[str, str]] = []
        for line_no, raw_line in enumerate(path.read_text(encoding="utf-8").splitlines(), start=1):
            match = _HEX_PARAGRAPH_RE.match(raw_line)
            if not match:
                raise ValueError(f"{path} 第 {line_no} 行格式错误，应为 [PARA_ID] 段落文本")
            paragraphs.append({"para_id": match.group(1), "text": match.group(2)})
        return paragraphs

    def _resolve_log_dir(self, material_dir: Path, stage_name: str) -> Path:
        base_dir = self.log_dir / material_dir.name if self.log_dir else material_dir / "logs"
        return base_dir / stage_name

    def _read_text_like_file(self, path: Path) -> list[str]:
        return self._split_text_content(path.read_text(encoding="utf-8"))

    def _split_text_content(self, content: str) -> list[str]:
        blocks = re.split(r"\n\s*\n+", content.replace("\r\n", "\n").replace("\r", "\n"))
        paragraphs: list[str] = []

        for block in blocks:
            lines = [line.strip() for line in block.split("\n") if line.strip()]
            if not lines:
                continue
            if len(lines) == 1:
                paragraphs.append(lines[0])
            elif self._should_split_multiline_block(lines):
                paragraphs.extend(lines)
            else:
                paragraphs.append(" ".join(lines))

        return paragraphs

    def _should_split_multiline_block(self, lines: list[str]) -> bool:
        structured_patterns = (
            r"^#{1,6}\s",
            r"^[-*+]\s",
            r"^\d+[.)]\s",
            r"^>\s",
            r"^\|",
            r"^```",
        )
        return any(re.match(pattern, line) for line in lines for pattern in structured_patterns)



def main() -> None:
    processor = MaterialProcessor(
        material_paths=["tests/material/material1.docx", "tests/material/material2.docx"],
        output_dir="tests/阶段五输出",
        max_window_chars=8000,
        log_dir="tests/阶段五输出/logs",
    )
    processor.process()


if __name__ == "__main__":
    main()
