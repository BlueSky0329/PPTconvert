from __future__ import annotations

import json
import re
from dataclasses import dataclass, field
from pathlib import Path
from typing import Any

from core.ai_repair import AIRepairBatchSummary, AIRepairService, repair_project_questions
from core.project_quality import annotate_project_quality
from domain.models import ExamProject, QuestionNode, Section, SUBJECT_DISPLAY_NAMES
from domain.project_editor import apply_all_safe_subject_suggestions

_OCR_NOISE_CHARS = {"�", "■", "□", "▢", "¦"}
_SPACED_ASCII_RE = re.compile(r"(?:\b[A-Za-z0-9]\b[\s\u3000]*){6,}")
_BROKEN_PUNCT_RE = re.compile(r"[|¦_~]{3,}")
_QUESTION_OCR_ISSUE_CODES = {
    "missing_stem",
    "blank_option",
    "option_count",
    "number_gap",
    "number_order",
    "duplicate_number",
}
_OCR_PROBE_PAGE_LIMIT = 3
_OCR_PROBE_SAMPLE_CHARS = 80


@dataclass(frozen=True)
class OCRDiagnosticIssue:
    code: str
    title: str
    detail: str
    severity: str = "warning"
    affected_count: int = 0

    def to_dict(self) -> dict[str, Any]:
        return {
            "code": self.code,
            "title": self.title,
            "detail": self.detail,
            "severity": self.severity,
            "affected_count": self.affected_count,
        }


@dataclass(frozen=True)
class OCRQuestionSample:
    question_no: str
    section_kind: str
    preview: str
    issue_codes: tuple[str, ...] = ()

    def to_dict(self) -> dict[str, Any]:
        return {
            "question_no": self.question_no,
            "section_kind": self.section_kind,
            "preview": self.preview,
            "issue_codes": list(self.issue_codes),
        }


@dataclass(frozen=True)
class OCRRecoveryPreview:
    page_number: int
    sample_text: str
    line_count: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "page_number": self.page_number,
            "sample_text": self.sample_text,
            "line_count": self.line_count,
        }


@dataclass
class PDFOCRDiagnosticReport:
    pdf_path: str = ""
    question_count: int = 0
    flagged_questions: int = 0
    severe_questions: int = 0
    suspicious_text_questions: int = 0
    fragmented_questions: int = 0
    image_only_questions: int = 0
    image_dominant_pages: int = 0
    low_text_pages: int = 0
    page_count: int = 0
    fitz_available: bool = False
    likely_scanned_pdf: bool = False
    likely_ocr_noise: bool = False
    ocr_available: bool = False
    ocr_recoverable_samples: list[OCRRecoveryPreview] = field(default_factory=list)
    issues: list[OCRDiagnosticIssue] = field(default_factory=list)
    question_samples: list[OCRQuestionSample] = field(default_factory=list)
    suggestions: list[str] = field(default_factory=list)

    def to_dict(self) -> dict[str, Any]:
        return {
            "pdf_path": self.pdf_path,
            "question_count": self.question_count,
            "flagged_questions": self.flagged_questions,
            "severe_questions": self.severe_questions,
            "suspicious_text_questions": self.suspicious_text_questions,
            "fragmented_questions": self.fragmented_questions,
            "image_only_questions": self.image_only_questions,
            "image_dominant_pages": self.image_dominant_pages,
            "low_text_pages": self.low_text_pages,
            "page_count": self.page_count,
            "fitz_available": self.fitz_available,
            "likely_scanned_pdf": self.likely_scanned_pdf,
            "likely_ocr_noise": self.likely_ocr_noise,
            "ocr_available": self.ocr_available,
            "ocr_recoverable_samples": [
                preview.to_dict() for preview in self.ocr_recoverable_samples
            ],
            "issues": [issue.to_dict() for issue in self.issues],
            "question_samples": [sample.to_dict() for sample in self.question_samples],
            "suggestions": list(self.suggestions),
        }

    def summary_text(self) -> str:
        parts = [
            f"共 {self.question_count} 题",
            f"待确认 {self.flagged_questions} 题",
        ]
        if self.likely_scanned_pdf:
            parts.append("疑似扫描版")
        if self.likely_ocr_noise:
            parts.append("疑似 OCR 噪声偏重")
        if self.suspicious_text_questions:
            parts.append(f"碎字/乱码 {self.suspicious_text_questions} 题")
        if self.fragmented_questions:
            parts.append(f"结构碎裂 {self.fragmented_questions} 题")
        if self.ocr_recoverable_samples:
            parts.append(f"OCR 预览 {len(self.ocr_recoverable_samples)} 页")
        return "；".join(parts)


@dataclass
class OCRAutoRepairSummary:
    diagnostics_before: PDFOCRDiagnosticReport
    diagnostics_after: PDFOCRDiagnosticReport
    section_subject_changes: int = 0
    batch_summary: AIRepairBatchSummary = field(default_factory=AIRepairBatchSummary)

    def to_dict(self) -> dict[str, Any]:
        return {
            "section_subject_changes": self.section_subject_changes,
            "batch_summary": {
                "attempted_questions": self.batch_summary.attempted_questions,
                "changed_questions": self.batch_summary.changed_questions,
                "total_field_changes": self.batch_summary.total_field_changes,
                "subject_changes": self.batch_summary.subject_changes,
                "structural_repairs": self.batch_summary.structural_repairs,
                "skipped_questions": self.batch_summary.skipped_questions,
                "errors": list(self.batch_summary.errors),
            },
            "diagnostics_before": self.diagnostics_before.to_dict(),
            "diagnostics_after": self.diagnostics_after.to_dict(),
        }

    def summary_text(self) -> str:
        before_flagged = self.diagnostics_before.flagged_questions
        after_flagged = self.diagnostics_after.flagged_questions
        return (
            f"科目纠偏 {self.section_subject_changes} 段；"
            f"AI 处理 {self.batch_summary.attempted_questions} 题，写回 {self.batch_summary.changed_questions} 题；"
            f"待确认从 {before_flagged} 题降到 {after_flagged} 题"
        )


def _preview_text(question: QuestionNode) -> str:
    preview = " ".join(((question.stem or "").strip()).split())
    if len(preview) > 42:
        preview = preview[:39] + "..."
    return preview or "[空题干]"


def _looks_like_ocr_noise(text: str) -> bool:
    normalized = " ".join((text or "").split())
    if not normalized:
        return False
    if any(char in normalized for char in _OCR_NOISE_CHARS):
        return True
    if _SPACED_ASCII_RE.search(normalized):
        return True
    if _BROKEN_PUNCT_RE.search(normalized):
        return True

    tokens = re.findall(r"[A-Za-z0-9]+", normalized)
    if len(tokens) >= 8:
        singletons = sum(1 for token in tokens if len(token) == 1)
        if singletons / max(len(tokens), 1) >= 0.7:
            return True
    return False


def _collect_pdf_page_signals(pdf_path: str | None) -> dict[str, Any]:
    if not pdf_path:
        return {"fitz_available": False}
    try:
        import fitz  # type: ignore
    except Exception:
        return {"fitz_available": False}

    path = Path(pdf_path)
    if not path.exists():
        return {"fitz_available": True}

    page_count = 0
    low_text_pages = 0
    image_dominant_pages = 0
    image_dominant_page_numbers: list[int] = []
    with fitz.open(path) as doc:
        page_count = len(doc)
        for page_index, page in enumerate(doc):
            text_len = len((page.get_text("text") or "").strip())
            image_count = len(page.get_images(full=True))
            if text_len < 24:
                low_text_pages += 1
            if image_count > 0 and text_len < 40:
                image_dominant_pages += 1
                image_dominant_page_numbers.append(page_index + 1)
    return {
        "fitz_available": True,
        "page_count": page_count,
        "low_text_pages": low_text_pages,
        "image_dominant_pages": image_dominant_pages,
        "image_dominant_page_numbers": image_dominant_page_numbers,
    }


def _probe_ocr_recovery(
    pdf_path: str | None,
    page_numbers: list[int],
    *,
    limit: int = _OCR_PROBE_PAGE_LIMIT,
) -> list[OCRRecoveryPreview]:
    if not pdf_path or not page_numbers:
        return []
    from core.pdf_ocr_engine import is_ocr_available, ocr_pdf_page

    if not is_ocr_available():
        return []
    previews: list[OCRRecoveryPreview] = []
    for page_number in page_numbers[:limit]:
        lines = ocr_pdf_page(pdf_path, page_number)
        if not lines:
            continue
        joined = " ".join(line.text.strip() for line in lines if line.text.strip())
        sample = joined[:_OCR_PROBE_SAMPLE_CHARS]
        if len(joined) > _OCR_PROBE_SAMPLE_CHARS:
            sample = sample.rstrip() + "..."
        previews.append(
            OCRRecoveryPreview(
                page_number=page_number,
                sample_text=sample,
                line_count=len(lines),
            )
        )
    return previews


def diagnose_project_ocr_risks(
    project: ExamProject,
    *,
    pdf_path: str | None = None,
    include_ocr_probe: bool = False,
) -> PDFOCRDiagnosticReport:
    """对当前 ExamProject 做扫描/OCR 风险诊断。

    ``include_ocr_probe=True`` 时，若判定为扫描版且 OCR 引擎可用，会在前若干张
    图像主导页上跑一遍 OCR，用结果样本填充 ``ocr_recoverable_samples``。
    该选项会显著增加耗时（每页 ~1-3s CPU），默认关闭。
    """
    quality = annotate_project_quality(project)
    resolved_pdf_path = pdf_path or project.source.pdf_path
    page_signals = _collect_pdf_page_signals(resolved_pdf_path)
    report = PDFOCRDiagnosticReport(
        pdf_path=str(resolved_pdf_path or ""),
        question_count=quality.question_count,
        flagged_questions=quality.flagged_questions,
        severe_questions=quality.severe_questions,
        page_count=int(page_signals.get("page_count") or 0),
        low_text_pages=int(page_signals.get("low_text_pages") or 0),
        image_dominant_pages=int(page_signals.get("image_dominant_pages") or 0),
        fitz_available=bool(page_signals.get("fitz_available")),
    )

    suspicious_samples: list[OCRQuestionSample] = []
    fragmented_samples: list[OCRQuestionSample] = []
    image_only_samples: list[OCRQuestionSample] = []

    for section, _material, question in project.iter_questions():
        text_blob = " ".join(
            [
                question.stem or "",
                *(option.text or "" for option in question.options),
            ]
        )
        issue_codes = tuple(issue.code for issue in getattr(question, "review_issues", []) or [])
        has_ocr_noise = _looks_like_ocr_noise(text_blob)
        is_fragmented = any(code in _QUESTION_OCR_ISSUE_CODES for code in issue_codes)
        is_image_only = not (question.stem or "").strip() and bool(question.stem_assets)

        if has_ocr_noise:
            report.suspicious_text_questions += 1
            if len(suspicious_samples) < 5:
                suspicious_samples.append(
                    OCRQuestionSample(
                        question_no=question.source_number or "-",
                        section_kind=section.kind,
                        preview=_preview_text(question),
                        issue_codes=issue_codes,
                    )
                )
        if is_fragmented:
            report.fragmented_questions += 1
            if len(fragmented_samples) < 5:
                fragmented_samples.append(
                    OCRQuestionSample(
                        question_no=question.source_number or "-",
                        section_kind=section.kind,
                        preview=_preview_text(question),
                        issue_codes=issue_codes,
                    )
                )
        if is_image_only:
            report.image_only_questions += 1
            if len(image_only_samples) < 5:
                image_only_samples.append(
                    OCRQuestionSample(
                        question_no=question.source_number or "-",
                        section_kind=section.kind,
                        preview=_preview_text(question),
                        issue_codes=issue_codes,
                    )
                )

    question_count = max(report.question_count, 1)
    report.likely_scanned_pdf = (
        report.page_count > 0
        and report.image_dominant_pages >= max(1, round(report.page_count * 0.35))
    ) or (
        report.image_only_questions >= max(1, round(question_count * 0.08))
    )
    report.likely_ocr_noise = (
        report.suspicious_text_questions >= max(2, round(question_count * 0.12))
        or report.fragmented_questions >= max(3, round(question_count * 0.18))
        or report.flagged_questions >= max(4, round(question_count * 0.28))
    )

    if report.likely_scanned_pdf:
        report.issues.append(
            OCRDiagnosticIssue(
                code="scan_pdf",
                title="疑似扫描版 PDF",
                detail="整页图片或低文本页占比较高，当前导入结果可能严重依赖 OCR 质量。",
                severity="warning",
                affected_count=max(report.image_dominant_pages, report.image_only_questions),
            )
        )
        report.suggestions.append("建议先运行 OCR 自动修补，再重点复核待确认题和图片题。")
    if report.suspicious_text_questions:
        report.issues.append(
            OCRDiagnosticIssue(
                code="ocr_noise_text",
                title="存在碎字或乱码迹象",
                detail="部分题干/选项出现单字符碎裂、乱码符号或异常断词，像是 OCR 文本噪声。",
                severity="warning",
                affected_count=report.suspicious_text_questions,
            )
        )
        report.suggestions.append("优先复核碎字/乱码样本，必要时回到原 PDF 页面对照。")
    if report.fragmented_questions:
        report.issues.append(
            OCRDiagnosticIssue(
                code="ocr_fragmented_structure",
                title="题目结构可能被 OCR 切碎",
                detail="存在空题干、空选项、题号跳跃或选项数量异常，说明 OCR 可能破坏了题目边界。",
                severity="error" if report.fragmented_questions >= 3 else "warning",
                affected_count=report.fragmented_questions,
            )
        )
        report.suggestions.append("可以直接运行 OCR 自动修补，让边界修复和材料回收先处理一轮。")
    if report.image_only_questions:
        report.issues.append(
            OCRDiagnosticIssue(
                code="ocr_image_only_question",
                title="存在几乎只有图片的题目",
                detail="部分题目没有正文文本，只保留了题干图片，通常需要结合原页截图复核。",
                severity="warning",
                affected_count=report.image_only_questions,
            )
        )
        report.suggestions.append("图片题较多时，建议在 GUI 里逐题查看题干图和选项图。")

    try:
        from core.pdf_ocr_engine import is_ocr_available

        report.ocr_available = is_ocr_available()
    except Exception:
        report.ocr_available = False

    if report.likely_scanned_pdf and not report.ocr_available:
        report.suggestions.append(
            "扫描版 PDF 建议安装 OCR 引擎（pip install -r requirements-ocr.txt），"
            "以便后续自动回收文字。"
        )

    if include_ocr_probe and report.likely_scanned_pdf and report.ocr_available:
        raw_page_numbers = page_signals.get("image_dominant_page_numbers") or []
        probe_page_numbers = [int(pn) for pn in raw_page_numbers if int(pn) > 0]
        report.ocr_recoverable_samples = _probe_ocr_recovery(
            resolved_pdf_path,
            probe_page_numbers,
        )
        if report.ocr_recoverable_samples:
            report.suggestions.append(
                f"已通过 OCR 预览回收 {len(report.ocr_recoverable_samples)} 页文字，"
                "后续可作为扫描版重建的依据。"
            )

    if not report.suggestions:
        report.suggestions.append("当前没有明显的扫描/OCR 高风险信号，仍建议抽查待确认题。")

    report.question_samples = suspicious_samples + fragmented_samples + image_only_samples
    return report


def write_ocr_diagnostic_report(report: PDFOCRDiagnosticReport | OCRAutoRepairSummary, output_path: str | Path) -> Path:
    path = Path(output_path)
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = report.to_dict()
    path.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    return path


def auto_repair_ocr_project(
    project: ExamProject,
    *,
    pdf_path: str | None = None,
    service: AIRepairService | None = None,
    only_flagged: bool = True,
    limit: int = 18,
    include_ocr_probe: bool = True,
) -> OCRAutoRepairSummary:
    diagnostics_before = diagnose_project_ocr_risks(
        project,
        pdf_path=pdf_path,
        include_ocr_probe=include_ocr_probe,
    )
    section_subject_changes = apply_all_safe_subject_suggestions(project)
    annotate_project_quality(project)
    repair_service = service or AIRepairService(mode="policy")
    batch_summary = repair_project_questions(
        project,
        service=repair_service,
        only_flagged=only_flagged,
        limit=limit,
    )
    diagnostics_after = diagnose_project_ocr_risks(
        project,
        pdf_path=pdf_path,
        include_ocr_probe=include_ocr_probe,
    )
    return OCRAutoRepairSummary(
        diagnostics_before=diagnostics_before,
        diagnostics_after=diagnostics_after,
        section_subject_changes=section_subject_changes,
        batch_summary=batch_summary,
    )
