from __future__ import annotations

from dataclasses import asdict
import json
import os

from domain.models import (
    AssetRef,
    ExamProject,
    MaterialSet,
    OptionNode,
    PageRegion,
    PaperSource,
    QuestionNode,
    QuestionRange,
    RepairLogEntry,
    ReviewIssue,
    Section,
)
from core.ppt_layout import sanitize_ppt_layout


def export_project_manifest(project: ExamProject, out_path: str) -> str:
    output_dir = os.path.dirname(os.path.abspath(out_path))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as file_obj:
        json.dump(asdict(project), file_obj, ensure_ascii=False, indent=2)
    return out_path


def load_project_manifest(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as file_obj:
        return json.load(file_obj)


def _load_page_region(data: dict | None) -> PageRegion | None:
    if not data:
        return None
    return PageRegion(
        page_number=int(data.get("page_number") or 0),
        x0=float(data.get("x0") or 0.0),
        y0=float(data.get("y0") or 0.0),
        x1=float(data.get("x1") or 0.0),
        y1=float(data.get("y1") or 0.0),
    )


def _load_asset_ref(data: dict) -> AssetRef:
    return AssetRef(
        kind=str(data.get("kind") or ""),
        path=str(data.get("path") or ""),
        source_page=data.get("source_page"),
        page_region=_load_page_region(data.get("page_region")),
        label=str(data.get("label") or ""),
    )


def _load_option(data: dict) -> OptionNode:
    return OptionNode(
        letter=str(data.get("letter") or ""),
        text=str(data.get("text") or ""),
        image_path=data.get("image_path") or None,
        source_page=data.get("source_page"),
        page_region=_load_page_region(data.get("page_region")),
    )


def _load_question(data: dict) -> QuestionNode:
    raw_suggested_confidence = data.get("suggested_subject_confidence")
    raw_subtype_confidence = data.get("inferred_subtype_confidence")
    return QuestionNode(
        source_number=str(data.get("source_number") or ""),
        stem=str(data.get("stem") or ""),
        question_id=str(data.get("question_id") or ""),
        options=[_load_option(item) for item in data.get("options", []) or []],
        stem_assets=[_load_asset_ref(item) for item in data.get("stem_assets", []) or []],
        answer=data.get("answer") or None,
        page_numbers=[int(value) for value in data.get("page_numbers", []) or []],
        option_layout=data.get("option_layout") or None,
        ppt_layout=sanitize_ppt_layout(data.get("ppt_layout") or {}),
        review_confidence=float(data.get("review_confidence") or 1.0),
        review_issues=[
            ReviewIssue(
                code=str(item.get("code") or ""),
                title=str(item.get("title") or ""),
                detail=str(item.get("detail") or ""),
                severity=str(item.get("severity") or "warning"),
            )
            for item in data.get("review_issues", []) or []
        ],
        suggested_subject=data.get("suggested_subject") or None,
        suggested_subject_confidence=(
            float(raw_suggested_confidence)
            if raw_suggested_confidence not in (None, "")
            else None
        ),
        suggested_subject_reason=str(data.get("suggested_subject_reason") or ""),
        inferred_subtype=str(data.get("inferred_subtype") or ""),
        inferred_subtype_confidence=(
            float(raw_subtype_confidence)
            if raw_subtype_confidence not in (None, "")
            else None
        ),
        inferred_signals=[str(value) for value in data.get("inferred_signals", []) or []],
    )


def _load_repair_log_entry(data: dict) -> RepairLogEntry:
    return RepairLogEntry(
        timestamp=str(data.get("timestamp") or ""),
        source=str(data.get("source") or ""),
        action=str(data.get("action") or ""),
        question_id=str(data.get("question_id") or ""),
        question_no=str(data.get("question_no") or ""),
        section_kind=str(data.get("section_kind") or "unknown"),
        material_id=str(data.get("material_id") or ""),
        before_state=dict(data.get("before_state") or {}),
        after_state=dict(data.get("after_state") or {}),
        metadata=dict(data.get("metadata") or {}),
    )


def _load_material(data: dict) -> MaterialSet:
    body_regions: list[PageRegion] = []
    for item in data.get("body_regions", []) or []:
        region = _load_page_region(item)
        if region is not None:
            body_regions.append(region)
    return MaterialSet(
        material_id=str(data.get("material_id") or ""),
        header=str(data.get("header") or ""),
        body=str(data.get("body") or ""),
        body_lines=[str(value) for value in data.get("body_lines", []) or []],
        body_assets=[_load_asset_ref(item) for item in data.get("body_assets", []) or []],
        body_regions=body_regions,
        questions=[_load_question(item) for item in data.get("questions", []) or []],
    )


def _load_section(data: dict) -> Section:
    return Section(
        kind=str(data.get("kind") or "unknown"),
        title=str(data.get("title") or ""),
        questions=[_load_question(item) for item in data.get("questions", []) or []],
        material_sets=[_load_material(item) for item in data.get("material_sets", []) or []],
    )


def _load_paper_source(data: dict | None) -> PaperSource:
    payload = data or {}
    return PaperSource(
        pdf_path=payload.get("pdf_path") or None,
        docx_path=payload.get("docx_path") or None,
        asset_dir=payload.get("asset_dir") or None,
    )


def _load_question_range(data: dict) -> QuestionRange:
    return QuestionRange(
        start=int(data.get("start") or 0),
        end=int(data.get("end") or 0),
    )


def load_project_manifest_project(path: str) -> ExamProject:
    from core.project_quality import annotate_project_quality

    payload = load_project_manifest(path)
    project = ExamProject(
        title=str(payload.get("title") or os.path.splitext(os.path.basename(path))[0]),
        source=_load_paper_source(payload.get("source")),
        sections=[_load_section(item) for item in payload.get("sections", []) or []],
        selected_subjects=[str(value) for value in payload.get("selected_subjects", []) or []],
        selected_ranges=[_load_question_range(item) for item in payload.get("selected_ranges", []) or []],
        repair_session_id=str(payload.get("repair_session_id") or ""),
        repair_log=[_load_repair_log_entry(item) for item in payload.get("repair_log", []) or []],
    )
    has_review_metadata = any(
        bool(question.review_issues)
        or float(question.review_confidence or 1.0) != 1.0
        or bool(question.suggested_subject)
        or bool(question.inferred_subtype)
        for _section, _material, question in project.iter_questions()
    )
    if not has_review_metadata:
        annotate_project_quality(project)
    return project
