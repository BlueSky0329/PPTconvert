from __future__ import annotations

import os
import shutil
from typing import Optional

from core.models import Question
from core.word_parser import WordParser
from domain.models import (
    ALL_SUBJECT_KINDS,
    AssetRef,
    ExamProject,
    MaterialSet,
    OptionNode,
    PaperSource,
    QuestionNode,
    Section,
    SubjectKind,
)


def _normalize_subject_kind(value: Optional[str]) -> SubjectKind:
    normalized = (value or "").strip().lower()
    if normalized in ALL_SUBJECT_KINDS:
        return normalized  # type: ignore[return-value]
    if normalized == "data":
        return "data"
    if normalized == "quant":
        return "quant"
    return "unknown"


def _default_section_title(kind: SubjectKind) -> str:
    mapping = {
        "politics": "政治理论",
        "common_sense": "常识判断",
        "verbal": "言语理解与表达",
        "quant": "数量关系",
        "reasoning": "判断推理",
        "data": "资料分析",
        "unknown": "题目列表",
    }
    return mapping.get(kind, "题目列表")


def _material_signature(question: Question) -> tuple[str, str, tuple[str, ...]]:
    header = (question.material_header or "").strip()
    body = (question.material_text or "").strip()
    images = tuple(path for path in getattr(question, "material_image_paths", []) or [] if path)
    return header, body, images


def _make_question_node(question: Question) -> QuestionNode:
    question_image_paths = list(getattr(question, "question_image_paths", []) or [])
    if not question_image_paths:
        material_image_paths = set(getattr(question, "material_image_paths", []) or [])
        question_image_paths = [
            path
            for path in (question.image_paths or [])
            if path and path not in material_image_paths
        ]

    return QuestionNode(
        source_number=str(question.source_question_number or question.number),
        stem=question.stem or "",
        options=[
            OptionNode(
                letter=option.letter,
                text=option.text or "",
                image_path=option.image_path,
            )
            for option in question.options
        ],
        stem_assets=[
            AssetRef(kind="image", path=path)
            for path in question_image_paths
            if path
        ],
        answer=question.answer,
        option_layout=question.option_layout,
    )


def build_exam_project_from_word_questions(
    questions: list[Question],
    *,
    title: str,
    docx_path: str | None = None,
    asset_dir: str | None = None,
) -> ExamProject:
    project = ExamProject(
        title=title,
        source=PaperSource(docx_path=docx_path, asset_dir=asset_dir),
    )

    current_section: Section | None = None
    current_material: MaterialSet | None = None

    for index, question in enumerate(questions, start=1):
        section_kind = _normalize_subject_kind(getattr(question, "section_kind", None))
        section_title = (getattr(question, "section_title", None) or "").strip() or _default_section_title(section_kind)
        needs_new_section = (
            current_section is None
            or current_section.kind != section_kind
            or current_section.title != section_title
            or (section_kind == "data" and current_section.kind != "data")
        )
        if needs_new_section:
            current_section = Section(kind=section_kind, title=section_title)
            project.sections.append(current_section)
            current_material = None

        question_node = _make_question_node(question)
        if current_section.kind != "data":
            current_section.questions.append(question_node)
            continue

        signature = _material_signature(question)
        if (
            current_material is None
            or _material_signature(question)
            != (
                current_material.header.strip(),
                current_material.body.strip(),
                tuple(asset.path for asset in current_material.body_assets if asset.path),
            )
        ):
            material_id = f"material-{len(current_section.material_sets) + 1}"
            header = (question.material_header or "").strip() or f"材料{len(current_section.material_sets) + 1}"
            body = (question.material_text or "").strip()
            current_material = MaterialSet(
                material_id=material_id,
                header=header,
                body=body,
                body_lines=[line for line in body.splitlines() if line.strip()],
                body_assets=[
                    AssetRef(kind="image", path=path)
                    for path in signature[2]
                    if path
                ],
            )
            current_section.material_sets.append(current_material)
        current_material.questions.append(question_node)

    seen_subjects: list[SubjectKind] = []
    for section in project.sections:
        if section.kind in ALL_SUBJECT_KINDS and section.kind not in seen_subjects:
            seen_subjects.append(section.kind)
    project.selected_subjects = seen_subjects
    return project


def build_exam_project_from_docx(
    docx_path: str,
    *,
    asset_dir: str | None = None,
    document_subject_hint: SubjectKind | None = None,
) -> tuple[ExamProject, list[Question], str]:
    chosen_asset_dir = asset_dir or os.path.splitext(os.path.abspath(docx_path))[0] + "_word_assets"
    if os.path.isdir(chosen_asset_dir):
        shutil.rmtree(chosen_asset_dir, ignore_errors=True)
    os.makedirs(chosen_asset_dir, exist_ok=True)

    parser = WordParser(temp_dir=chosen_asset_dir, document_subject_hint=document_subject_hint)
    try:
        questions = parser.parse(docx_path)
    finally:
        parser.cleanup()

    project = build_exam_project_from_word_questions(
        questions,
        title=os.path.splitext(os.path.basename(docx_path))[0],
        docx_path=docx_path,
        asset_dir=chosen_asset_dir,
    )
    return project, questions, chosen_asset_dir
