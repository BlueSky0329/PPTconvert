from __future__ import annotations

import argparse
import copy
import json
from collections import Counter
from pathlib import Path
import sys
from typing import Any

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.repair_action_features import build_repair_state_record
from domain.models import ALL_SUBJECT_KINDS, ExamProject, MaterialSet, OptionNode, QuestionNode, Section, SubjectKind
from ingest.pdf.project_builder import build_exam_project_from_pdf


DEFAULT_CATALOG = ROOT / "data" / "gold_pdf_catalog.json"
DEFAULT_OUTPUT = ROOT / "data" / "datasets" / "repair_actions.jsonl"

ACTION_FAMILY_MAP: dict[str, str] = {
    "renumber_current_question": "boundary_repair",
    "split_embedded_next_question": "boundary_repair",
    "move_spilled_option_back": "boundary_repair",
    "move_data_intro_back_to_material": "material_repair",
    "move_data_assets_to_material": "material_repair",
    "reassign_stem_image_to_options": "asset_repair",
}


def _material_text(material: MaterialSet | None) -> str:
    if material is None:
        return ""
    body = (material.body or "").strip()
    if body:
        return body
    return "\n".join(line.strip() for line in material.body_lines if (line or "").strip()).strip()


def _iter_question_rows(project: ExamProject):
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                for index, question in enumerate(material.questions):
                    prev_question = material.questions[index - 1] if index > 0 else None
                    next_question = material.questions[index + 1] if index + 1 < len(material.questions) else None
                    yield section, material, index, question, prev_question, next_question
        else:
            for index, question in enumerate(section.questions):
                prev_question = section.questions[index - 1] if index > 0 else None
                next_question = section.questions[index + 1] if index + 1 < len(section.questions) else None
                yield section, None, index, question, prev_question, next_question


def _make_row(
    *,
    source_pdf: str,
    source_form: str,
    action: str,
    corruption: str,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
    target: dict[str, Any],
) -> dict[str, Any]:
    action_family = ACTION_FAMILY_MAP.get(action, "other_repair")
    return {
        "source_pdf": source_pdf,
        "source_form": source_form,
        "subject": section.kind,
        "action": action,
        "action_family": action_family,
        "corruption": corruption,
        "question_no": question.source_number,
        "previous_question_no": previous_question.source_number if previous_question is not None else "",
        "next_question_no": next_question.source_number if next_question is not None else "",
        "state_record": build_repair_state_record(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        ),
        "target": target,
    }


def _row_for_renumber_current_question(
    *,
    source_pdf: str,
    source_form: str,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
) -> dict[str, Any] | None:
    if previous_question is None:
        return None
    previous_number = previous_question.numeric_source_number
    current_number = question.numeric_source_number
    if previous_number is None or current_number is None or current_number != previous_number + 1:
        return None
    corrupted = copy.deepcopy(question)
    corrupted.source_number = str(current_number + 1)
    return _make_row(
        source_pdf=source_pdf,
        source_form=source_form,
        action="renumber_current_question",
        corruption="question_number_shifted_forward",
        section=section,
        material=material,
        question=corrupted,
        previous_question=previous_question,
        next_question=next_question,
        target={"source_number": str(current_number)},
    )


def _row_for_move_spilled_option_back(
    *,
    source_pdf: str,
    source_form: str,
    section: Section,
    current_question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
) -> dict[str, Any] | None:
    if previous_question is None or len(previous_question.options) != 4:
        return None
    spilled_option = previous_question.options[-1]
    if not (spilled_option.text or "").strip():
        return None
    corrupted_prev = copy.deepcopy(previous_question)
    corrupted_prev.options = corrupted_prev.options[:3]
    corrupted_current = copy.deepcopy(current_question)
    current_stem = (corrupted_current.stem or "").strip()
    if not current_stem:
        return None
    corrupted_current.stem = (
        f"{spilled_option.letter}. {spilled_option.text} {current_question.source_number}. {current_stem}"
    )
    return _make_row(
        source_pdf=source_pdf,
        source_form=source_form,
        action="move_spilled_option_back",
        corruption="previous_option_spilled_into_current_stem",
        section=section,
        material=None,
        question=corrupted_current,
        previous_question=corrupted_prev,
        next_question=next_question,
        target={
            "previous_question_no": previous_question.source_number,
            "restored_option_letter": spilled_option.letter,
            "restored_option_text": spilled_option.text,
            "current_stem": current_question.stem,
        },
    )


def _row_for_split_embedded_next_question(
    *,
    source_pdf: str,
    source_form: str,
    section: Section,
    current_question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
) -> dict[str, Any] | None:
    if next_question is None:
        return None
    next_stem = (next_question.stem or "").strip()
    next_number = (next_question.source_number or "").strip()
    if not next_stem or not next_number:
        return None
    corrupted_current = copy.deepcopy(current_question)
    corrupted_current.stem = f"{(corrupted_current.stem or '').strip()} {next_number}. {next_stem}".strip()
    corrupted_next = copy.deepcopy(next_question)
    corrupted_next.stem = ""
    return _make_row(
        source_pdf=source_pdf,
        source_form=source_form,
        action="split_embedded_next_question",
        corruption="next_question_embedded_in_current_stem",
        section=section,
        material=None,
        question=corrupted_current,
        previous_question=previous_question,
        next_question=corrupted_next,
        target={
            "current_stem": current_question.stem,
            "next_question_no": next_question.source_number,
            "next_stem": next_question.stem,
        },
    )


def _row_for_move_data_intro_back_to_material(
    *,
    source_pdf: str,
    source_form: str,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
) -> dict[str, Any] | None:
    if material is None:
        return None
    material_body = _material_text(material)
    if not material_body:
        return None
    first_question = material.questions[0] if material.questions else None
    if first_question is not question:
        return None
    corrupted_material = copy.deepcopy(material)
    corrupted_material.body = ""
    corrupted_material.body_lines = []
    corrupted_question = copy.deepcopy(question)
    corrupted_question.stem = f"{material_body} {question.stem}".strip()
    return _make_row(
        source_pdf=source_pdf,
        source_form=source_form,
        action="move_data_intro_back_to_material",
        corruption="material_intro_embedded_in_first_question",
        section=section,
        material=corrupted_material,
        question=corrupted_question,
        previous_question=previous_question,
        next_question=next_question,
        target={
            "material_body": material_body,
            "question_stem": question.stem,
        },
    )


def _row_for_move_data_assets_to_material(
    *,
    source_pdf: str,
    source_form: str,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
) -> dict[str, Any] | None:
    if material is None or not material.body_assets:
        return None
    first_question = material.questions[0] if material.questions else None
    if first_question is not question:
        return None
    corrupted_material = copy.deepcopy(material)
    corrupted_question = copy.deepcopy(question)
    corrupted_question.stem_assets = [copy.deepcopy(asset) for asset in material.body_assets]
    corrupted_material.body_assets = []
    return _make_row(
        source_pdf=source_pdf,
        source_form=source_form,
        action="move_data_assets_to_material",
        corruption="material_assets_bound_to_first_question",
        section=section,
        material=corrupted_material,
        question=corrupted_question,
        previous_question=previous_question,
        next_question=next_question,
        target={
            "asset_paths": [asset.path for asset in material.body_assets],
        },
    )


def _row_for_reassign_stem_image_to_options(
    *,
    source_pdf: str,
    source_form: str,
    section: Section,
    question: QuestionNode,
    previous_question: QuestionNode | None,
    next_question: QuestionNode | None,
) -> dict[str, Any] | None:
    if len(question.options) != 4 or question.stem_assets:
        return None
    image_options = [option for option in question.options if option.image_path]
    if len(image_options) != 4:
        return None
    corrupted_question = copy.deepcopy(question)
    corrupted_question.stem_assets = [
        {
            "kind": "image",
            "path": option.image_path,
        }
        for option in corrupted_question.options
        if option.image_path
    ]
    for option in corrupted_question.options:
        option.image_path = None
    return _make_row(
        source_pdf=source_pdf,
        source_form=source_form,
        action="reassign_stem_image_to_options",
        corruption="option_images_attached_to_stem",
        section=section,
        material=None,
        question=corrupted_question,
        previous_question=previous_question,
        next_question=next_question,
        target={
            "option_images": {
                option.letter: option.image_path
                for option in question.options
            }
        },
    )


def generate_project_repair_rows(
    project: ExamProject,
    *,
    source_pdf: str,
    source_form: str,
) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    for section, material, _index, question, previous_question, next_question in _iter_question_rows(project):
        for builder in (
            lambda: _row_for_renumber_current_question(
                source_pdf=source_pdf,
                source_form=source_form,
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            ),
            lambda: _row_for_split_embedded_next_question(
                source_pdf=source_pdf,
                source_form=source_form,
                section=section,
                current_question=question,
                previous_question=previous_question,
                next_question=next_question,
            ),
            lambda: _row_for_move_spilled_option_back(
                source_pdf=source_pdf,
                source_form=source_form,
                section=section,
                current_question=question,
                previous_question=previous_question,
                next_question=next_question,
            ),
            lambda: _row_for_move_data_intro_back_to_material(
                source_pdf=source_pdf,
                source_form=source_form,
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            ),
            lambda: _row_for_move_data_assets_to_material(
                source_pdf=source_pdf,
                source_form=source_form,
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            ),
            lambda: _row_for_reassign_stem_image_to_options(
                source_pdf=source_pdf,
                source_form=source_form,
                section=section,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            ),
        ):
            row = builder()
            if row is not None:
                rows.append(row)
    return rows


def _catalog_subject_hint(entry: dict[str, Any]) -> SubjectKind | None:
    if str(entry.get("form", "")).strip() != "single_subject_book":
        return None
    subject = str(entry.get("subject", "")).strip()
    if subject in ALL_SUBJECT_KINDS:
        return subject  # type: ignore[return-value]
    return None


def build_repair_action_dataset(catalog_path: Path, output_path: Path) -> dict[str, Any]:
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    rows: list[dict[str, Any]] = []
    action_counter: Counter[str] = Counter()
    family_counter: Counter[str] = Counter()
    pdf_counter: Counter[str] = Counter()

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
        project_rows = generate_project_repair_rows(project, source_pdf=source_pdf, source_form=source_form)
        rows.extend(project_rows)
        pdf_counter.update([source_pdf] * len(project_rows))
        action_counter.update(row["action"] for row in project_rows)
        family_counter.update(row["action_family"] for row in project_rows)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")

    summary = {
        "catalog": str(catalog_path),
        "output": str(output_path),
        "row_count": len(rows),
        "pdf_count": len({row["source_pdf"] for row in rows}),
        "action_distribution": dict(action_counter),
        "action_family_distribution": dict(family_counter),
        "pdf_distribution": dict(pdf_counter),
    }
    summary_path = output_path.with_suffix(".summary.json")
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="从金标准 PDF 工程构建修复动作训练集")
    parser.add_argument("--catalog", type=Path, default=DEFAULT_CATALOG)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    summary = build_repair_action_dataset(args.catalog, args.output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
