from __future__ import annotations

import json
import os

from core.project_quality import iter_flagged_question_rows, question_max_severity, question_review_summary
from domain.models import ExamProject, SUBJECT_DISPLAY_NAMES


def build_quality_report_payload(project: ExamProject) -> dict:
    items: list[dict] = []
    for section, material, question in iter_flagged_question_rows(project):
        items.append(
            {
                "source_number": question.source_number,
                "section_kind": section.kind,
                "section_title": section.title,
                "section_label": SUBJECT_DISPLAY_NAMES.get(section.kind, section.kind),
                "material_id": getattr(material, "material_id", None),
                "material_header": getattr(material, "header", None),
                "confidence": round(float(getattr(question, "review_confidence", 1.0) or 1.0), 4),
                "severity": question_max_severity(question),
                "summary": question_review_summary(question),
                "suggested_subject": getattr(question, "suggested_subject", None),
                "suggested_subject_label": (
                    SUBJECT_DISPLAY_NAMES.get(question.suggested_subject, question.suggested_subject)
                    if getattr(question, "suggested_subject", None)
                    else None
                ),
                "suggested_subject_confidence": getattr(question, "suggested_subject_confidence", None),
                "suggested_subject_reason": getattr(question, "suggested_subject_reason", ""),
                "issues": [
                    {
                        "code": issue.code,
                        "title": issue.title,
                        "detail": issue.detail,
                        "severity": issue.severity,
                    }
                    for issue in getattr(question, "review_issues", []) or []
                ],
            }
        )
    return {
        "title": project.title,
        "question_count": project.question_count,
        "flagged_question_count": len(items),
        "items": items,
    }


def export_quality_report(project: ExamProject, out_path: str) -> str:
    output_dir = os.path.dirname(os.path.abspath(out_path))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    payload = build_quality_report_payload(project)
    with open(out_path, "w", encoding="utf-8") as file_obj:
        json.dump(payload, file_obj, ensure_ascii=False, indent=2)
    return out_path
