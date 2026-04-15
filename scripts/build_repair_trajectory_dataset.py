"""从金标准 PDF 工程构建多步修复轨迹训练集。

每条轨迹描述：对同一道题连续注入多种错误后，按什么顺序、用什么动作修复。
这是模仿学习第二阶段的数据源，和单步修复动作数据集互补。
"""
from __future__ import annotations

import argparse
import copy
import json
import random
from collections import Counter
from collections import defaultdict
from pathlib import Path
import sys
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.repair_action_features import build_repair_state_record
from domain.models import ALL_SUBJECT_KINDS, ExamProject, MaterialSet, OptionNode, QuestionNode, Section, SubjectKind
from exporters.manifest_json import load_project_manifest_project
from ingest.pdf.project_builder import build_exam_project_from_pdf
from scripts.build_repair_action_dataset import ACTION_FAMILY_MAP, generate_project_repair_rows

DEFAULT_CATALOG = ROOT / "data" / "gold_pdf_catalog.json"
DEFAULT_OUTPUT = ROOT / "data" / "datasets" / "repair_trajectories.jsonl"

# 可组合的错误注入器：每个返回 (corrupted_question, action_name, corruption_desc, target)
# 或 None（表示此错误不适用于该题）


def _inject_number_shift(question: QuestionNode, previous: QuestionNode | None) -> tuple | None:
    if previous is None:
        return None
    prev_num = previous.numeric_source_number
    cur_num = question.numeric_source_number
    if prev_num is None or cur_num is None or cur_num != prev_num + 1:
        return None
    corrupted = copy.deepcopy(question)
    corrupted.source_number = str(cur_num + 1)
    return (
        corrupted,
        "renumber_current_question",
        "question_number_shifted_forward",
        {"source_number": str(cur_num)},
    )


def _inject_spilled_option(question: QuestionNode, previous: QuestionNode | None) -> tuple | None:
    if previous is None or len(previous.options) != 4:
        return None
    spilled = previous.options[-1]
    if not (spilled.text or "").strip():
        return None
    stem = (question.stem or "").strip()
    if not stem:
        return None
    corrupted = copy.deepcopy(question)
    corrupted.stem = f"{spilled.letter}. {spilled.text} {question.source_number}. {stem}"
    return (
        corrupted,
        "move_spilled_option_back",
        "previous_option_spilled_into_current_stem",
        {"restored_option_letter": spilled.letter, "current_stem": question.stem},
    )


def _inject_embedded_next(question: QuestionNode, next_q: QuestionNode | None) -> tuple | None:
    if next_q is None:
        return None
    next_stem = (next_q.stem or "").strip()
    next_num = (next_q.source_number or "").strip()
    if not next_stem or not next_num:
        return None
    corrupted = copy.deepcopy(question)
    corrupted.stem = f"{(corrupted.stem or '').strip()} {next_num}. {next_stem}".strip()
    return (
        corrupted,
        "split_embedded_next_question",
        "next_question_embedded_in_current_stem",
        {"current_stem": question.stem, "next_stem": next_stem},
    )


def _inject_stem_assets_as_options(question: QuestionNode) -> tuple | None:
    if len(question.options) != 4 or question.stem_assets:
        return None
    image_options = [o for o in question.options if o.image_path]
    if len(image_options) != 4:
        return None
    corrupted = copy.deepcopy(question)
    corrupted.stem_assets = [
        {"kind": "image", "path": o.image_path} for o in corrupted.options if o.image_path
    ]
    for o in corrupted.options:
        o.image_path = None
    return (
        corrupted,
        "reassign_stem_image_to_options",
        "option_images_attached_to_stem",
        {"option_images": {o.letter: o.image_path for o in question.options}},
    )


INJECTORS = [
    _inject_number_shift,
    _inject_spilled_option,
    _inject_embedded_next,
]


def _build_step_record(
    *,
    step_index: int,
    action: str,
    corruption: str,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
    target: dict[str, Any],
) -> dict[str, Any]:
    return {
        "step": step_index,
        "action": action,
        "action_family": ACTION_FAMILY_MAP.get(action, "other_repair"),
        "corruption": corruption,
        "state_record": build_repair_state_record(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        ),
        "target": target,
    }


def _iter_question_rows(project: ExamProject):
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                for index, question in enumerate(material.questions):
                    prev_q = material.questions[index - 1] if index > 0 else None
                    next_q = material.questions[index + 1] if index + 1 < len(material.questions) else None
                    yield section, material, question, prev_q, next_q
        else:
            for index, question in enumerate(section.questions):
                prev_q = section.questions[index - 1] if index > 0 else None
                next_q = section.questions[index + 1] if index + 1 < len(section.questions) else None
                yield section, None, question, prev_q, next_q


def generate_trajectories_for_project(
    project: ExamProject,
    *,
    source_pdf: str,
    source_form: str,
    max_steps: int = 3,
    seed: int = 42,
) -> list[dict[str, Any]]:
    rng = random.Random(seed)
    trajectories: list[dict[str, Any]] = []
    grouped_rows: dict[tuple[str, str, str, str], list[dict[str, Any]]] = defaultdict(list)
    for row in generate_project_repair_rows(project, source_pdf=source_pdf, source_form=source_form):
        question_no = str(row.get("question_no") or "")
        if str(row.get("action") or "") == "renumber_current_question":
            question_no = str((row.get("target") or {}).get("source_number") or question_no)
        key = (
            str(row.get("subject") or "unknown"),
            question_no,
            str(row.get("previous_question_no") or ""),
            str(row.get("next_question_no") or ""),
        )
        grouped_rows[key].append(row)

    priority = {"boundary_repair": 0, "material_repair": 1, "asset_repair": 2, "other_repair": 3}
    for (subject, question_no, _previous_no, _next_no), applicable_rows in grouped_rows.items():
        if not applicable_rows:
            continue

        for row in applicable_rows:
            trajectories.append({
                "source_pdf": source_pdf,
                "source_form": source_form,
                "subject": subject,
                "question_no": question_no,
                "trajectory_length": 1,
                "steps": [{
                    "step": 0,
                    "action": str(row.get("action") or ""),
                    "action_family": str(row.get("action_family") or ACTION_FAMILY_MAP.get(str(row.get("action") or ""), "other_repair")),
                    "corruption": str(row.get("corruption") or ""),
                    "state_record": dict(row.get("state_record") or {}),
                    "target": dict(row.get("target") or {}),
                }],
            })

        if len(applicable_rows) >= 2:
            for combo_size in range(2, min(len(applicable_rows), max_steps) + 1):
                combo = rng.sample(applicable_rows, combo_size)
                combo.sort(
                    key=lambda row: priority.get(
                        str(row.get("action_family") or ACTION_FAMILY_MAP.get(str(row.get("action") or ""), "other_repair")),
                        3,
                    )
                )
                steps: list[dict[str, Any]] = []
                for step_idx, row in enumerate(combo):
                    steps.append({
                        "step": step_idx,
                        "action": str(row.get("action") or ""),
                        "action_family": str(row.get("action_family") or ACTION_FAMILY_MAP.get(str(row.get("action") or ""), "other_repair")),
                        "corruption": str(row.get("corruption") or ""),
                        "state_record": dict(row.get("state_record") or {}),
                        "target": dict(row.get("target") or {}),
                    })
                trajectories.append({
                    "source_pdf": source_pdf,
                    "source_form": source_form,
                    "subject": subject,
                    "question_no": question_no,
                    "trajectory_length": len(steps),
                    "steps": steps,
                })

    return trajectories


def _catalog_subject_hint(entry: dict[str, Any]) -> SubjectKind | None:
    if str(entry.get("form", "")).strip() != "single_subject_book":
        return None
    subject = str(entry.get("subject", "")).strip()
    if subject in ALL_SUBJECT_KINDS:
        return subject  # type: ignore[return-value]
    return None


def _iter_manifest_paths(inputs: list[str]) -> list[Path]:
    paths: list[Path] = []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            paths.extend(sorted(path.rglob("*.json")))
        elif any(token in item for token in ("*", "?")):
            paths.extend(sorted(Path().glob(item)))
        else:
            paths.append(path)
    deduped: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        resolved = str(path.resolve())
        if resolved in seen:
            continue
        seen.add(resolved)
        deduped.append(path)
    return deduped


def _iter_jsonl_paths(inputs: list[str]) -> list[Path]:
    paths: list[Path] = []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            paths.extend(sorted(path.rglob("*.jsonl")))
        elif any(token in item for token in ("*", "?")):
            paths.extend(sorted(Path().glob(item)))
        else:
            paths.append(path)
    deduped: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        resolved = str(path.resolve())
        if resolved in seen:
            continue
        seen.add(resolved)
        deduped.append(path)
    return deduped


def _entry_get(entry: Any, key: str, default: Any = None) -> Any:
    if isinstance(entry, dict):
        return entry.get(key, default)
    return getattr(entry, key, default)


def _snapshot_state_record(snapshot: dict[str, Any] | None) -> dict[str, Any]:
    record = dict((snapshot or {}).get("state_record") or {})
    record["meta"] = dict(record.get("meta") or {})
    return record


def _snapshot_meta(snapshot: dict[str, Any] | None) -> dict[str, Any]:
    return dict(_snapshot_state_record(snapshot).get("meta") or {})


def _float_meta(meta: dict[str, Any], key: str) -> float:
    try:
        return float(meta.get(key) or 0.0)
    except (TypeError, ValueError):
        return 0.0


def _infer_gui_repair_action(entry) -> str | None:
    explicit_action_map = {
        "renumber_question": "renumber_current_question",
    }
    raw_action = str(_entry_get(entry, "action", "") or "")
    if raw_action in explicit_action_map:
        return explicit_action_map[raw_action]

    before_state = dict(_entry_get(entry, "before_state", {}) or {})
    after_state = dict(_entry_get(entry, "after_state", {}) or {})
    if not before_state or not after_state:
        return None

    before_meta = _snapshot_meta(before_state)
    after_meta = _snapshot_meta(after_state)

    before_stem_assets = _float_meta(before_meta, "stem_asset_count")
    after_stem_assets = _float_meta(after_meta, "stem_asset_count")
    before_option_images = _float_meta(before_meta, "option_image_count")
    after_option_images = _float_meta(after_meta, "option_image_count")
    before_material_assets = _float_meta(before_meta, "material_asset_count")
    after_material_assets = _float_meta(after_meta, "material_asset_count")
    before_material_length = _float_meta(before_meta, "material_length")
    after_material_length = _float_meta(after_meta, "material_length")
    before_stem_length = _float_meta(before_meta, "stem_length")
    after_stem_length = _float_meta(after_meta, "stem_length")
    before_gap = _float_meta(before_meta, "number_gap_from_previous")
    after_gap = _float_meta(after_meta, "number_gap_from_previous")

    if after_option_images > before_option_images and after_stem_assets < before_stem_assets:
        return "reassign_stem_image_to_options"
    if after_material_assets > before_material_assets and after_stem_assets < before_stem_assets:
        return "move_data_assets_to_material"
    if (
        str(before_state.get("section_kind") or "") == "data"
        and after_material_length > before_material_length + 8
        and after_stem_length + 8 < before_stem_length
    ):
        return "move_data_intro_back_to_material"
    if (
        str(before_state.get("question_no") or "") != str(after_state.get("question_no") or "")
        or (abs(before_gap) >= 1.0 and abs(after_gap) < abs(before_gap))
    ):
        return "renumber_current_question"
    if (
        _float_meta(before_meta, "stem_starts_with_option_marker") > 0.0
        and _float_meta(after_meta, "stem_starts_with_option_marker") <= 0.0
    ):
        return "move_spilled_option_back"
    if (
        _float_meta(before_meta, "stem_has_embedded_question_no") > 0.0
        and _float_meta(after_meta, "stem_has_embedded_question_no") <= 0.0
    ):
        return "split_embedded_next_question"
    return None


def _build_gui_entry_trajectories(
    entries: list[tuple[int, Any, str]],
    *,
    source_pdf: str,
    session_id: str,
    question_id: str,
    source_form: str,
    trajectory_origin: str,
    manifest_path: str = "",
    dataset_path: str = "",
) -> list[dict[str, Any]]:
    entries.sort(key=lambda item: (str(_entry_get(item[1], "timestamp", "") or ""), item[0]))
    ordered_steps: list[dict[str, Any]] = []
    question_no = ""
    subject = "unknown"

    for _index, entry, action in entries:
        state_record = _snapshot_state_record(_entry_get(entry, "before_state", {}) or {})
        if not state_record.get("text"):
            continue
        question_no = str(_entry_get(entry, "question_no", question_no) or question_no or "")
        subject = str(_entry_get(entry, "section_kind", subject) or subject or "unknown")
        ordered_steps.append({
            "step": len(ordered_steps),
            "action": action,
            "action_family": ACTION_FAMILY_MAP.get(action, "other_repair"),
            "corruption": f"gui_log::{_entry_get(entry, 'action', '')}",
            "state_record": state_record,
            "target": {
                "question_id": _entry_get(entry, "question_id", "") or "",
                "question_no": _entry_get(entry, "question_no", "") or "",
                "source_action": _entry_get(entry, "action", "") or "",
                "metadata": dict(_entry_get(entry, "metadata", {}) or {}),
                "after_state": dict(_entry_get(entry, "after_state", {}) or {}),
            },
        })

    trajectories: list[dict[str, Any]] = []
    if not ordered_steps:
        return trajectories

    for start in range(len(ordered_steps)):
        suffix_steps: list[dict[str, Any]] = []
        for step_index, step in enumerate(ordered_steps[start:]):
            suffix_step = dict(step)
            suffix_step["step"] = step_index
            suffix_step["target"] = dict(step.get("target") or {})
            suffix_step["state_record"] = dict(step.get("state_record") or {})
            suffix_step["state_record"]["meta"] = dict(
                (suffix_step["state_record"].get("meta") or {})
            )
            suffix_steps.append(suffix_step)
        trajectories.append({
            "source_pdf": source_pdf,
            "source_form": source_form,
            "subject": subject,
            "question_no": question_no,
            "question_id": question_id,
            "trajectory_length": len(suffix_steps),
            "steps": suffix_steps,
            "trajectory_origin": trajectory_origin,
            "manifest_path": manifest_path,
            "dataset_path": dataset_path,
            "session_id": session_id,
        })
    return trajectories


def build_trajectories_from_gui_logs(manifest_inputs: list[str]) -> list[dict[str, Any]]:
    manifest_paths = _iter_manifest_paths(manifest_inputs)
    trajectories: list[dict[str, Any]] = []

    for manifest_path in manifest_paths:
        if not manifest_path.exists():
            continue
        try:
            project = load_project_manifest_project(str(manifest_path))
        except Exception:
            continue

        grouped_entries: dict[tuple[str, str], list[tuple[int, Any, str]]] = defaultdict(list)
        for index, entry in enumerate(getattr(project, "repair_log", []) or []):
            question_id = str(_entry_get(entry, "question_id", "") or "")
            if not (question_id and _entry_get(entry, "before_state", None) and _entry_get(entry, "after_state", None)):
                continue
            action = _infer_gui_repair_action(entry)
            if not action:
                continue
            grouped_entries[(getattr(project, "repair_session_id", "") or "", question_id)].append(
                (index, entry, action)
            )

        source_pdf = (
            str(getattr(getattr(project, "source", None), "pdf_path", "") or "").strip()
            or manifest_path.as_posix()
        )
        for (session_id, question_id), entries in grouped_entries.items():
            trajectories.extend(
                _build_gui_entry_trajectories(
                    entries,
                    source_pdf=source_pdf,
                    session_id=session_id,
                    question_id=question_id,
                    source_form="gui_repair_log",
                    trajectory_origin="gui_log_manifest",
                    manifest_path=str(manifest_path),
                )
            )

    return trajectories


def build_trajectories_from_gui_dataset_files(jsonl_inputs: list[str]) -> list[dict[str, Any]]:
    jsonl_paths = _iter_jsonl_paths(jsonl_inputs)
    trajectories: list[dict[str, Any]] = []

    for dataset_path in jsonl_paths:
        if not dataset_path.exists():
            continue
        grouped_entries: dict[tuple[str, str, str], list[tuple[int, Any, str]]] = defaultdict(list)
        try:
            lines = dataset_path.read_text(encoding="utf-8").splitlines()
        except Exception:
            continue
        for index, line in enumerate(lines):
            line = line.strip()
            if not line:
                continue
            try:
                entry = json.loads(line)
            except Exception:
                continue
            question_id = str(_entry_get(entry, "question_id", "") or "")
            session_id = str(_entry_get(entry, "session_id", "") or "")
            if not (question_id and _entry_get(entry, "before_state", None) and _entry_get(entry, "after_state", None)):
                continue
            action = _infer_gui_repair_action(entry)
            if not action:
                continue
            source_pdf = str(_entry_get(entry, "source_pdf", "") or "").strip()
            manifest_path = str(_entry_get(entry, "manifest_path", "") or "").strip()
            source_key = source_pdf or manifest_path or dataset_path.as_posix()
            grouped_entries[(session_id, question_id, source_key)].append((index, entry, action))

        for (session_id, question_id, source_pdf), entries in grouped_entries.items():
            manifest_path = str(_entry_get(entries[0][1], "manifest_path", "") or "")
            trajectories.extend(
                _build_gui_entry_trajectories(
                    entries,
                    source_pdf=source_pdf,
                    session_id=session_id,
                    question_id=question_id,
                    source_form="gui_repair_log_jsonl",
                    trajectory_origin="gui_log_jsonl",
                    manifest_path=manifest_path,
                    dataset_path=str(dataset_path),
                )
            )

    return trajectories


def build_trajectory_dataset(
    catalog_path: Path,
    output_path: Path,
    *,
    gui_manifest_inputs: list[str] | None = None,
    gui_jsonl_inputs: list[str] | None = None,
) -> dict[str, Any]:
    catalog = json.loads(catalog_path.read_text(encoding="utf-8")) if catalog_path.exists() else {"pdfs": []}
    trajectories: list[dict[str, Any]] = []
    action_counter: Counter[str] = Counter()
    length_counter: Counter[int] = Counter()
    origin_counter: Counter[str] = Counter()
    synthetic_count = 0

    for entry in catalog.get("pdfs", []) or []:
        rel_path = Path(str(entry.get("path", "")))
        pdf_path = ROOT / rel_path
        if not pdf_path.exists():
            continue
        project = build_exam_project_from_pdf(
            str(pdf_path),
            mode="all",
            document_subject_hint=_catalog_subject_hint(entry),
        )
        source_pdf = rel_path.as_posix()
        source_form = str(entry.get("form", "unknown")).strip() or "unknown"
        project_trajectories = generate_trajectories_for_project(
            project, source_pdf=source_pdf, source_form=source_form,
        )
        trajectories.extend(project_trajectories)
        synthetic_count += len(project_trajectories)
        for traj in project_trajectories:
            length_counter[traj["trajectory_length"]] += 1
            origin_counter["synthetic"] += 1
            for step in traj["steps"]:
                action_counter[step["action"]] += 1

    gui_manifest_trajectories = build_trajectories_from_gui_logs(gui_manifest_inputs or [])
    gui_jsonl_trajectories = build_trajectories_from_gui_dataset_files(gui_jsonl_inputs or [])
    gui_trajectories = gui_manifest_trajectories + gui_jsonl_trajectories
    trajectories.extend(gui_trajectories)
    for traj in gui_trajectories:
        length_counter[traj["trajectory_length"]] += 1
        origin_counter[str(traj.get("trajectory_origin") or "gui_log")] += 1
        for step in traj["steps"]:
            action_counter[step["action"]] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        for traj in trajectories:
            fh.write(json.dumps(traj, ensure_ascii=False) + "\n")

    summary = {
        "catalog": str(catalog_path),
        "output": str(output_path),
        "trajectory_count": len(trajectories),
        "synthetic_trajectory_count": synthetic_count,
        "gui_trajectory_count": len(gui_trajectories),
        "gui_manifest_trajectory_count": len(gui_manifest_trajectories),
        "gui_jsonl_trajectory_count": len(gui_jsonl_trajectories),
        "length_distribution": dict(sorted(length_counter.items())),
        "origin_distribution": dict(origin_counter),
        "action_distribution": dict(action_counter),
    }
    summary_path = output_path.with_suffix(".summary.json")
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="从金标准 PDF 构建多步修复轨迹训练集")
    parser.add_argument("--catalog", type=Path, default=DEFAULT_CATALOG)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--gui-manifests",
        nargs="*",
        default=[],
        help="可选：附加导入带 repair_log 的工程 JSON / 目录 / glob",
    )
    parser.add_argument(
        "--gui-jsonl",
        nargs="*",
        default=[],
        help="可选：附加导入 gui_repair_logs.jsonl 文件、目录或 glob",
    )
    args = parser.parse_args()

    summary = build_trajectory_dataset(
        args.catalog,
        args.output,
        gui_manifest_inputs=args.gui_manifests,
        gui_jsonl_inputs=args.gui_jsonl,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
