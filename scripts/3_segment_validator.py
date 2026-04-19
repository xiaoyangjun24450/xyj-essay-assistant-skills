#!/usr/bin/env python3
"""Validate and repair template segment outputs."""

from __future__ import annotations

import json
import re
from collections import Counter
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Callable, Optional

from pipeline_utils import call_llm_messages, load_prompt, normalize_hex_para_id


@dataclass(frozen=True)
class Paragraph:
    para_id: str
    text: str

    def render(self) -> str:
        return f"[{self.para_id}] {self.text}"


class SegmentValidator:
    """Validate stage-2 segment output and repair issues with local rules + LLM."""

    def __init__(
        self,
        full_paragraphs_path: str,
        input_segments_path: str,
        output_path: str,
        log_dir: str,
        max_rounds: int = 5,
        max_repair_batch_size: int = 15,
        max_llm_repair_attempts: int = 3,
        llm_caller: Optional[Callable[[list[dict[str, str]], str, Optional[str]], str]] = None,
    ):
        self.full_paragraphs_path = Path(full_paragraphs_path).expanduser().resolve()
        self.input_segments_path = Path(input_segments_path).expanduser().resolve()
        self.output_path = Path(output_path).expanduser().resolve()
        self.log_dir = Path(log_dir).expanduser().resolve()
        self.max_rounds = max_rounds
        self.max_repair_batch_size = max(1, max_repair_batch_size)
        self.max_llm_repair_attempts = max(1, max_llm_repair_attempts)
        self.llm_caller = llm_caller or call_llm_messages

        self.paragraphs = self._load_full_paragraphs()
        self.all_para_ids = [paragraph.para_id for paragraph in self.paragraphs]
        self.para_index = {paragraph.para_id: index for index, paragraph in enumerate(self.paragraphs)}
        self.paragraph_map = {paragraph.para_id: paragraph for paragraph in self.paragraphs}

        self.system_prompt = load_prompt("segment_repair_system_prompt")
        self.user_prompt_template = load_prompt("segment_repair_user_prompt")

    def run(self) -> dict[str, Any]:
        segments = self._cleanup_and_renumber(self._load_segments(self.input_segments_path))
        self._write_segments(segments)

        for round_index in range(1, self.max_rounds + 1):
            report = self._validate_segments(segments, round_index, "initial")
            self._write_json_log(f"validation_round_{round_index:03d}.json", report)
            if report["passed"]:
                return {
                    "passed": True,
                    "rounds": round_index,
                    "output_path": str(self.output_path),
                    "report": report,
                }

            if report["unknown_para_ids"]:
                segments, unknown_log = self._remove_unknown_para_ids(segments, report["unknown_para_ids"])
                segments = self._cleanup_and_renumber(segments)
                self._write_segments(segments)
                self._write_json_log(
                    f"repair_round_{round_index:03d}_remove_unknown.json",
                    unknown_log,
                )

            post_unknown_report = self._validate_segments(segments, round_index, "post_unknown_cleanup")
            duplicate_para_ids = post_unknown_report["duplicate_para_ids"]
            if duplicate_para_ids:
                segments, duplicate_log = self._remove_duplicate_para_ids(segments, duplicate_para_ids)
                segments = self._cleanup_and_renumber(segments)
                self._write_segments(segments)
                self._write_json_log(
                    f"repair_round_{round_index:03d}_remove_duplicate.json",
                    duplicate_log,
                )

            repair_report = self._validate_segments(segments, round_index, "pre_reinsert")
            if repair_report["missing_para_ids"]:
                before_segments = deepcopy(segments)
                decisions, llm_log = self._request_batch_repair_decisions(
                    segments,
                    repair_report["missing_para_ids"],
                    round_index,
                )
                segments, apply_log = self._apply_batch_repair_decisions(
                    segments,
                    repair_report["missing_para_ids"],
                    decisions,
                )
                segments = self._cleanup_and_renumber(segments)
                self._write_segments(segments)
                self._write_json_log(
                    f"repair_round_{round_index:03d}_reinsert_batch.json",
                    {
                        "round": round_index,
                        "para_ids": repair_report["missing_para_ids"],
                        "llm": llm_log,
                        "apply": apply_log,
                        "before_segments": before_segments,
                        "after_segments": segments,
                    },
                )

        final_report = self._validate_segments(segments, self.max_rounds, "final")
        self._write_json_log("validation_failed_final.json", final_report)
        raise ValueError(
            "segment 修复未在限定轮数内完成。"
            f" 输出文件: {self.output_path} 日志目录: {self.log_dir}"
        )

    def _load_full_paragraphs(self) -> list[Paragraph]:
        if not self.full_paragraphs_path.exists():
            raise FileNotFoundError(f"full_paragraphs 文件未找到: {self.full_paragraphs_path}")

        paragraphs: list[Paragraph] = []
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

            para_id = normalize_hex_para_id(match.group(1))
            if para_id in seen_para_ids:
                raise ValueError(f"full_paragraphs 中存在重复 para_id: {para_id}")

            seen_para_ids.add(para_id)
            paragraphs.append(Paragraph(para_id=para_id, text=match.group(2)))

        if not paragraphs:
            raise ValueError(f"未在 {self.full_paragraphs_path} 中读取到任何段落")
        return paragraphs

    def _load_segments(self, path: Path) -> list[dict[str, Any]]:
        if not path.exists():
            raise FileNotFoundError(f"segments 文件未找到: {path}")

        payload = json.loads(path.read_text(encoding="utf-8"))
        if not isinstance(payload, list):
            raise ValueError("segments 文件必须是 JSON 数组")

        normalized_segments: list[dict[str, Any]] = []
        for index, segment in enumerate(payload, start=1):
            if not isinstance(segment, dict):
                raise ValueError(f"第 {index} 个 segment 不是 JSON 对象")

            raw_para_ids = segment.get("para_ids", [])
            if not isinstance(raw_para_ids, list):
                raw_para_ids = []

            normalized_segments.append(
                {
                    "segment_id": str(segment.get("segment_id") or f"seg_{index}").strip(),
                    "summary": str(segment.get("summary", "")).strip(),
                    "purpose": str(segment.get("purpose", "")).strip(),
                    "para_ids": [
                        normalize_hex_para_id(str(para_id))
                        for para_id in raw_para_ids
                        if str(para_id).strip()
                    ],
                }
            )

        return normalized_segments

    def _validate_segments(
        self,
        segments: list[dict[str, Any]],
        round_index: int,
        phase: str,
    ) -> dict[str, Any]:
        flattened_para_ids: list[str] = []
        para_id_locations: dict[str, list[dict[str, Any]]] = {}
        unknown_para_ids: list[str] = []
        seen_unknown = set()

        for segment_index, segment in enumerate(segments):
            for position, para_id in enumerate(segment["para_ids"]):
                flattened_para_ids.append(para_id)
                para_id_locations.setdefault(para_id, []).append(
                    {
                        "segment_index": segment_index,
                        "segment_id": segment["segment_id"],
                        "position": position,
                    }
                )
                if para_id not in self.para_index and para_id not in seen_unknown:
                    unknown_para_ids.append(para_id)
                    seen_unknown.add(para_id)

        known_counter = Counter(para_id for para_id in flattened_para_ids if para_id in self.para_index)
        missing_para_ids = [
            para_id
            for para_id in self.all_para_ids
            if known_counter.get(para_id, 0) == 0
        ]
        duplicate_para_ids = [
            para_id
            for para_id in self.all_para_ids
            if known_counter.get(para_id, 0) > 1
        ]

        return {
            "round": round_index,
            "phase": phase,
            "passed": not missing_para_ids and not duplicate_para_ids and not unknown_para_ids,
            "input_segments_path": str(self.input_segments_path),
            "output_path": str(self.output_path),
            "expected_paragraph_count": len(self.all_para_ids),
            "segment_count": len(segments),
            "flattened_paragraph_count": len(flattened_para_ids),
            "missing_para_ids": missing_para_ids,
            "duplicate_para_ids": duplicate_para_ids,
            "unknown_para_ids": unknown_para_ids,
            "para_id_locations": para_id_locations,
            "message": (
                f"missing={len(missing_para_ids)} "
                f"duplicate={len(duplicate_para_ids)} "
                f"unknown={len(unknown_para_ids)}"
            ),
        }

    def _remove_unknown_para_ids(
        self,
        segments: list[dict[str, Any]],
        unknown_para_ids: list[str],
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        unknown_set = set(unknown_para_ids)
        updated_segments = deepcopy(segments)
        segment_changes: list[dict[str, Any]] = []

        for segment in updated_segments:
            before_para_ids = list(segment["para_ids"])
            after_para_ids = [para_id for para_id in before_para_ids if para_id not in unknown_set]
            removed_para_ids = [para_id for para_id in before_para_ids if para_id in unknown_set]
            if removed_para_ids:
                segment["para_ids"] = after_para_ids
                segment_changes.append(
                    {
                        "segment_id": segment["segment_id"],
                        "removed_para_ids": removed_para_ids,
                        "before_para_ids": before_para_ids,
                        "after_para_ids": after_para_ids,
                    }
                )

        return updated_segments, {
            "action": "remove_unknown_para_ids",
            "unknown_para_ids": unknown_para_ids,
            "segment_changes": segment_changes,
        }

    def _remove_duplicate_para_ids(
        self,
        segments: list[dict[str, Any]],
        duplicate_para_ids: list[str],
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        duplicate_set = set(duplicate_para_ids)
        updated_segments = deepcopy(segments)
        segment_changes: list[dict[str, Any]] = []

        for segment in updated_segments:
            before_para_ids = list(segment["para_ids"])
            after_para_ids = [para_id for para_id in before_para_ids if para_id not in duplicate_set]
            removed_para_ids = [para_id for para_id in before_para_ids if para_id in duplicate_set]
            if removed_para_ids:
                segment["para_ids"] = after_para_ids
                segment_changes.append(
                    {
                        "segment_id": segment["segment_id"],
                        "removed_para_ids": removed_para_ids,
                        "before_para_ids": before_para_ids,
                        "after_para_ids": after_para_ids,
                    }
                )

        return updated_segments, {
            "action": "remove_duplicate_para_ids",
            "duplicate_para_ids": duplicate_para_ids,
            "segment_changes": segment_changes,
        }

    def _request_batch_repair_decisions(
        self,
        segments: list[dict[str, Any]],
        para_ids: list[str],
        round_index: int,
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        all_decisions: list[dict[str, Any]] = []
        batch_logs: list[dict[str, Any]] = []

        for batch_index, batch_para_ids in enumerate(
            self._chunk_para_ids(para_ids, self.max_repair_batch_size),
            start=1,
        ):
            decisions, batch_log = self._request_single_batch_repair_decisions(
                segments=segments,
                para_ids=batch_para_ids,
                round_index=round_index,
                batch_index=batch_index,
            )
            all_decisions.extend(decisions)
            batch_logs.append(batch_log)

        return all_decisions, {
            "round": round_index,
            "batch_size": self.max_repair_batch_size,
            "total_para_ids": len(para_ids),
            "batch_count": len(batch_logs),
            "batches": batch_logs,
        }

    def _request_single_batch_repair_decisions(
        self,
        segments: list[dict[str, Any]],
        para_ids: list[str],
        round_index: int,
        batch_index: int,
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        repair_items: list[dict[str, Any]] = []
        for para_id in para_ids:
            context = self._build_repair_context(segments, para_id)
            paragraph = self.paragraph_map[para_id]
            repair_items.append(
                {
                    "para_id": para_id,
                    "paragraph": {"para_id": paragraph.para_id, "text": paragraph.text},
                    "previous_paragraph": context["previous_paragraph"],
                    "next_paragraph": context["next_paragraph"],
                    "candidate_before_segment": context["candidate_before_segment"],
                    "candidate_after_segment": context["candidate_after_segment"],
                    "candidate_segment_ids": context["candidate_segment_ids"],
                    "allowed_actions": context["allowed_actions"],
                    "force_target_segment_id": context["force_target_segment_id"],
                }
            )

        prompt = self.user_prompt_template.format(
            repair_items_json=self._json_dumps(repair_items),
            current_segments_json=self._json_dumps(segments),
        )

        messages = [
            {"role": "system", "content": self.system_prompt},
            {"role": "user", "content": prompt},
        ]
        request_context = {
            "para_ids": para_ids,
            "repair_items": repair_items,
        }
        base_filename = f"segment_repair_round_{round_index:03d}_batch_{batch_index:03d}"
        attempt_logs: list[dict[str, Any]] = []
        last_error: Optional[ValueError] = None

        for attempt_index in range(1, self.max_llm_repair_attempts + 1):
            filename = (
                base_filename
                if attempt_index == 1
                else f"{base_filename}_retry_{attempt_index:02d}"
            )
            response_text = self.llm_caller(messages, filename, str(self.log_dir))
            attempt_log: dict[str, Any] = {
                "attempt": attempt_index,
                "filename": filename,
                "request_context": request_context,
                "response_text": response_text,
            }

            try:
                decision = self._parse_response_json(response_text)
                self._validate_batch_repair_decisions(decision, para_ids)
                self._apply_batch_repair_decisions(segments, para_ids, decision)
            except ValueError as exc:
                last_error = exc
                attempt_log["error"] = str(exc)
                self._write_json_log(
                    f"{filename}_validation_error.json",
                    attempt_log,
                )
                attempt_logs.append(attempt_log)
                continue

            attempt_log["decision"] = decision
            attempt_logs.append(attempt_log)
            return decision, {
                "batch_index": batch_index,
                "filename": filename,
                "attempts": attempt_logs,
                "request_context": request_context,
                "response_text": response_text,
                "decision": decision,
            }

        assert last_error is not None
        raise ValueError(
            "LLM 批量修复结果在多次重试后仍未通过校验。"
            f" attempts={self.max_llm_repair_attempts} last_error={last_error}"
        )

    def _chunk_para_ids(
        self,
        para_ids: list[str],
        batch_size: int,
    ) -> list[list[str]]:
        return [
            para_ids[index : index + batch_size]
            for index in range(0, len(para_ids), batch_size)
        ]

    def _apply_batch_repair_decisions(
        self,
        segments: list[dict[str, Any]],
        para_ids: list[str],
        decisions: list[dict[str, Any]],
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        updated_segments = deepcopy(segments)
        applied_items: list[dict[str, Any]] = []

        for expected_para_id, decision in zip(para_ids, decisions):
            decision_para_id = str(decision.get("para_id", "")).strip()
            para_id = decision_para_id or expected_para_id
            if para_id != expected_para_id:
                raise ValueError(
                    "LLM 返回的批量修复结果 para_id 顺序不匹配，"
                    f" expected={expected_para_id} actual={para_id}"
                )

            before_item_segments = deepcopy(updated_segments)
            updated_segments, apply_log = self._apply_repair_decision(
                updated_segments,
                para_id,
                decision,
            )
            applied_items.append(
                {
                    "para_id": para_id,
                    "decision": decision,
                    "apply_log": apply_log,
                    "before_segments": before_item_segments,
                    "after_segments": deepcopy(updated_segments),
                }
            )

        return updated_segments, {
            "action": "batch_reinsert",
            "para_ids": para_ids,
            "items": applied_items,
        }

    def _build_repair_context(
        self,
        segments: list[dict[str, Any]],
        para_id: str,
    ) -> dict[str, Any]:
        previous_para_id = self._find_existing_neighbor_para_id(segments, para_id, direction=-1)
        next_para_id = self._find_existing_neighbor_para_id(segments, para_id, direction=1)
        previous_location = self._find_para_location(segments, previous_para_id) if previous_para_id else None
        next_location = self._find_para_location(segments, next_para_id) if next_para_id else None

        candidate_before_segment = (
            deepcopy(segments[previous_location["segment_index"]])
            if previous_location is not None
            else None
        )
        candidate_after_segment = (
            deepcopy(segments[next_location["segment_index"]])
            if next_location is not None
            else None
        )

        candidate_segment_ids: list[str] = []
        for segment in (candidate_before_segment, candidate_after_segment):
            if segment and segment["segment_id"] not in candidate_segment_ids:
                candidate_segment_ids.append(segment["segment_id"])

        force_target_segment_id = None
        allowed_actions = ["append_to_existing", "create_new"]
        if (
            previous_location is not None
            and next_location is not None
            and previous_location["segment_index"] == next_location["segment_index"]
        ):
            force_target_segment_id = segments[previous_location["segment_index"]]["segment_id"]
            allowed_actions = ["append_to_existing"]
            candidate_segment_ids = [force_target_segment_id]

        return {
            "previous_paragraph": self._make_paragraph_payload(previous_para_id),
            "next_paragraph": self._make_paragraph_payload(next_para_id),
            "previous_location": previous_location,
            "next_location": next_location,
            "candidate_before_segment": candidate_before_segment,
            "candidate_after_segment": candidate_after_segment,
            "candidate_segment_ids": candidate_segment_ids,
            "allowed_actions": allowed_actions,
            "force_target_segment_id": force_target_segment_id,
        }

    def _apply_repair_decision(
        self,
        segments: list[dict[str, Any]],
        para_id: str,
        decision: dict[str, Any],
    ) -> tuple[list[dict[str, Any]], dict[str, Any]]:
        context = self._build_repair_context(segments, para_id)
        updated_segments = deepcopy(segments)

        forced_target_segment_id = context["force_target_segment_id"]
        action = str(decision.get("action", "")).strip()
        if forced_target_segment_id:
            action = "append_to_existing"
            target_segment_id = forced_target_segment_id
        else:
            if action not in context["allowed_actions"]:
                raise ValueError(
                    f"LLM 返回了不允许的 action: {action}, para_id={para_id}"
                )
            target_segment_id = str(decision.get("target_segment_id", "")).strip()

        if action == "append_to_existing":
            if not target_segment_id:
                target_segment_id = self._pick_default_target_segment_id(context)
            if target_segment_id not in context["candidate_segment_ids"]:
                raise ValueError(
                    f"LLM 选择的 target_segment_id 不在候选范围内: {target_segment_id}"
                )

            segment_index = self._find_segment_index(updated_segments, target_segment_id)
            if segment_index is None:
                raise ValueError(f"未找到目标 segment: {target_segment_id}")

            before_para_ids = list(updated_segments[segment_index]["para_ids"])
            insert_position = self._find_insert_position(before_para_ids, para_id)
            updated_segments[segment_index]["para_ids"].insert(insert_position, para_id)
            return updated_segments, {
                "action": "append_to_existing",
                "target_segment_id": target_segment_id,
                "insert_position": insert_position,
                "before_para_ids": before_para_ids,
                "after_para_ids": updated_segments[segment_index]["para_ids"],
            }

        summary = str(decision.get("summary", "")).strip()
        purpose = str(decision.get("purpose", "")).strip()
        if not summary or not purpose:
            raise ValueError(f"create_new 缺少 summary 或 purpose: para_id={para_id}")

        insert_index = self._find_new_segment_insert_index(context, len(updated_segments))
        updated_segments.insert(
            insert_index,
            {
                "segment_id": "pending_segment",
                "summary": summary,
                "purpose": purpose,
                "para_ids": [para_id],
            },
        )
        return updated_segments, {
            "action": "create_new",
            "insert_index": insert_index,
            "summary": summary,
            "purpose": purpose,
        }

    def _cleanup_and_renumber(self, segments: list[dict[str, Any]]) -> list[dict[str, Any]]:
        cleaned_segments: list[dict[str, Any]] = []
        for segment in segments:
            para_ids = [str(para_id).strip() for para_id in segment.get("para_ids", []) if str(para_id).strip()]
            if not para_ids:
                continue

            cleaned_segments.append(
                {
                    "segment_id": segment.get("segment_id", ""),
                    "summary": str(segment.get("summary", "")).strip(),
                    "purpose": str(segment.get("purpose", "")).strip(),
                    "para_ids": para_ids,
                }
            )

        for index, segment in enumerate(cleaned_segments, start=1):
            segment["segment_id"] = f"seg_{index}"
        return cleaned_segments

    def _write_segments(self, segments: list[dict[str, Any]]) -> None:
        self.output_path.parent.mkdir(parents=True, exist_ok=True)
        self.output_path.write_text(
            json.dumps(segments, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _write_json_log(self, name: str, payload: dict[str, Any]) -> None:
        self.log_dir.mkdir(parents=True, exist_ok=True)
        log_path = self.log_dir / name
        log_path.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )

    def _validate_batch_repair_decisions(
        self,
        decisions: Any,
        para_ids: list[str],
    ) -> None:
        if not isinstance(decisions, list):
            raise ValueError("LLM 返回的批量修复结果不是 JSON 数组")
        if len(decisions) != len(para_ids):
            raise ValueError(
                "LLM 返回的批量修复结果数量不匹配，"
                f" expected={len(para_ids)} actual={len(decisions)}"
            )

        for index, (expected_para_id, decision) in enumerate(zip(para_ids, decisions), start=1):
            if not isinstance(decision, dict):
                raise ValueError(
                    "LLM 返回的批量修复结果条目不是 JSON 对象，"
                    f" index={index} para_id={expected_para_id}"
                )

            decision_para_id = str(decision.get("para_id", "")).strip()
            if decision_para_id and decision_para_id != expected_para_id:
                raise ValueError(
                    "LLM 返回的批量修复结果 para_id 顺序不匹配，"
                    f" expected={expected_para_id} actual={decision_para_id}"
                )

    def _find_existing_neighbor_para_id(
        self,
        segments: list[dict[str, Any]],
        para_id: str,
        direction: int,
    ) -> Optional[str]:
        existing_para_ids = {
            current_para_id
            for segment in segments
            for current_para_id in segment["para_ids"]
            if current_para_id in self.para_index
        }
        start_index = self.para_index[para_id]
        cursor = start_index + direction
        while 0 <= cursor < len(self.all_para_ids):
            candidate = self.all_para_ids[cursor]
            if candidate in existing_para_ids:
                return candidate
            cursor += direction
        return None

    def _find_para_location(
        self,
        segments: list[dict[str, Any]],
        para_id: str,
    ) -> Optional[dict[str, Any]]:
        for segment_index, segment in enumerate(segments):
            for position, current_para_id in enumerate(segment["para_ids"]):
                if current_para_id == para_id:
                    return {
                        "segment_index": segment_index,
                        "segment_id": segment["segment_id"],
                        "position": position,
                    }
        return None

    def _find_segment_index(
        self,
        segments: list[dict[str, Any]],
        segment_id: str,
    ) -> Optional[int]:
        for index, segment in enumerate(segments):
            if segment["segment_id"] == segment_id:
                return index
        return None

    def _find_insert_position(self, para_ids: list[str], para_id: str) -> int:
        para_position = self.para_index[para_id]
        for index, current_para_id in enumerate(para_ids):
            current_position = self.para_index.get(current_para_id)
            if current_position is not None and current_position > para_position:
                return index
        return len(para_ids)

    def _find_new_segment_insert_index(
        self,
        context: dict[str, Any],
        segment_count: int,
    ) -> int:
        next_location = context["next_location"]
        previous_location = context["previous_location"]
        if next_location is not None:
            return next_location["segment_index"]
        if previous_location is not None:
            return previous_location["segment_index"] + 1
        return segment_count

    def _pick_default_target_segment_id(self, context: dict[str, Any]) -> str:
        candidate_segment_ids = context["candidate_segment_ids"]
        if candidate_segment_ids:
            return candidate_segment_ids[0]
        raise ValueError("没有可用的候选 segment 可供插回")

    def _make_paragraph_payload(self, para_id: Optional[str]) -> Any:
        if not para_id:
            return None
        paragraph = self.paragraph_map[para_id]
        return {"para_id": paragraph.para_id, "text": paragraph.text}

    def _parse_response_json(self, response_text: str) -> Any:
        text = response_text.strip()
        candidates = [text]

        fenced = re.findall(r"```(?:json)?\s*(.*?)```", text, flags=re.S)
        candidates.extend(item.strip() for item in fenced if item.strip())

        object_match = re.search(r"(\{\s*.*\s*\})", text, flags=re.S)
        if object_match:
            candidates.append(object_match.group(1).strip())

        array_match = re.search(r"(\[\s*.*\s*\])", text, flags=re.S)
        if array_match:
            candidates.append(array_match.group(1).strip())

        for candidate in candidates:
            try:
                return json.loads(candidate)
            except json.JSONDecodeError:
                continue

        raise ValueError("LLM 返回的修复结果不是有效 JSON")

    def _json_dumps(self, payload: Any) -> str:
        return json.dumps(payload, ensure_ascii=False, indent=2)


def main() -> None:
    validator = SegmentValidator(
        full_paragraphs_path="tests/阶段一输出/full_paragraphs.txt",
        input_segments_path="tests/阶段二输出/template_segments.json",
        output_path="tests/阶段三输出/template_segments.json",
        log_dir="tests/阶段三输出/logs",
    )
    result = validator.run()


if __name__ == "__main__":
    main()
