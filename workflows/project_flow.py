from __future__ import annotations

from dataclasses import dataclass
import os

from domain.selectors import parse_question_ranges, parse_subject_spec, select_project
from exporters.docx_booklet import export_project_to_docx
from exporters.manifest_json import export_project_manifest
from exporters.pptx_slides import export_project_to_pptx
from ingest.docx.project_builder import build_exam_project_from_docx
from ingest.pdf.project_builder import build_exam_project_from_pdf


@dataclass
class ProjectOutputs:
    asset_dir: str
    docx_path: str | None = None
    pptx_path: str | None = None
    manifest_path: str | None = None


def _default_asset_dir(pdf_path: str) -> str:
    base_dir = os.path.dirname(os.path.abspath(pdf_path))
    pdf_name = os.path.splitext(os.path.basename(pdf_path))[0]
    return os.path.join(base_dir, f"{pdf_name}_assets")


def _default_docx_asset_dir(docx_path: str) -> str:
    base_dir = os.path.dirname(os.path.abspath(docx_path))
    docx_name = os.path.splitext(os.path.basename(docx_path))[0]
    return os.path.join(base_dir, f"{docx_name}_word_assets")


def build_pdf_project(
    pdf_path: str,
    *,
    mode: str = "all",
    question_range_spec: str = "",
    asset_dir: str | None = None,
    document_subject_hint=None,
):
    chosen_asset_dir = asset_dir or _default_asset_dir(pdf_path)
    selected_subjects = parse_subject_spec(mode)
    project = build_exam_project_from_pdf(
        pdf_path,
        mode="all",
        asset_dir=chosen_asset_dir,
        document_subject_hint=document_subject_hint,
    )

    question_ranges = parse_question_ranges(question_range_spec)
    project = select_project(
        project,
        subjects=selected_subjects,
        question_ranges=question_ranges,
    )
    return project, chosen_asset_dir


def build_word_project(
    docx_path: str,
    *,
    asset_dir: str | None = None,
    document_subject_hint=None,
):
    chosen_asset_dir = asset_dir or _default_docx_asset_dir(docx_path)
    project, questions, _asset_dir = build_exam_project_from_docx(
        docx_path,
        asset_dir=chosen_asset_dir,
        document_subject_hint=document_subject_hint,
    )
    return project, questions, chosen_asset_dir


def export_project_outputs(
    project,
    *,
    asset_dir: str,
    docx_output: str | None = None,
    ppt_output: str | None = None,
    manifest_output: str | None = None,
    template_path: str | None = None,
    ppt_config=None,
    font_name: str = "微软雅黑",
):
    outputs = ProjectOutputs(asset_dir=asset_dir)
    if manifest_output:
        outputs.manifest_path = export_project_manifest(project, manifest_output)
    if docx_output:
        outputs.docx_path = export_project_to_docx(project, docx_output, font_name=font_name)
    if ppt_output:
        outputs.pptx_path = export_project_to_pptx(
            project,
            ppt_output,
            template_path=template_path,
            config=ppt_config,
        )
    return outputs


def process_pdf_project(
    pdf_path: str,
    *,
    mode: str = "all",
    question_range_spec: str = "",
    asset_dir: str | None = None,
    docx_output: str | None = None,
    ppt_output: str | None = None,
    manifest_output: str | None = None,
    template_path: str | None = None,
    ppt_config=None,
    font_name: str = "微软雅黑",
):
    project, chosen_asset_dir = build_pdf_project(
        pdf_path,
        mode=mode,
        question_range_spec=question_range_spec,
        asset_dir=asset_dir,
    )
    outputs = export_project_outputs(
        project,
        asset_dir=chosen_asset_dir,
        docx_output=docx_output,
        ppt_output=ppt_output,
        manifest_output=manifest_output,
        template_path=template_path,
        ppt_config=ppt_config,
        font_name=font_name,
    )
    return project, outputs
