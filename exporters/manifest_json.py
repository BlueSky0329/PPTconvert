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
    Section,
)


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
    return QuestionNode(
        source_number=str(data.get("source_number") or ""),
        stem=str(data.get("stem") or ""),
        options=[_load_option(item) for item in data.get("options", []) or []],
        stem_assets=[_load_asset_ref(item) for item in data.get("stem_assets", []) or []],
        answer=data.get("answer") or None,
        page_numbers=[int(value) for value in data.get("page_numbers", []) or []],
        option_layout=data.get("option_layout") or None,
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
    payload = load_project_manifest(path)
    return ExamProject(
        title=str(payload.get("title") or os.path.splitext(os.path.basename(path))[0]),
        source=_load_paper_source(payload.get("source")),
        sections=[_load_section(item) for item in payload.get("sections", []) or []],
        selected_subjects=[str(value) for value in payload.get("selected_subjects", []) or []],
        selected_ranges=[_load_question_range(item) for item in payload.get("selected_ranges", []) or []],
    )
