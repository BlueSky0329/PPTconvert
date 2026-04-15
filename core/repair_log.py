from __future__ import annotations

from datetime import datetime, timezone
from uuid import uuid4

from core.repair_action_features import build_repair_state_record
from domain.models import ExamProject, MaterialSet, QuestionNode, RepairLogEntry, Section


def _utc_now() -> str:
    return datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z")


def ensure_project_tracking_ids(project: ExamProject) -> None:
    if not project.repair_session_id:
        project.repair_session_id = uuid4().hex
    for _section, _material, question in project.iter_questions():
        ensure_question_id(question)


def ensure_question_id(question: QuestionNode) -> str:
    if not question.question_id:
        question.question_id = uuid4().hex
    return question.question_id


def capture_question_state(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
) -> dict:
    ensure_question_id(question)
    state_record = build_repair_state_record(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    )
    return {
        "question_id": question.question_id,
        "question_no": question.source_number or "",
        "section_kind": section.kind,
        "material_id": material.material_id if material is not None else "",
        "review_confidence": float(getattr(question, "review_confidence", 1.0) or 1.0),
        "review_issue_codes": [issue.code for issue in (getattr(question, "review_issues", None) or [])],
        "state_record": state_record,
    }


def append_question_repair_log(
    project: ExamProject,
    *,
    source: str,
    action: str,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
    before_state: dict | None = None,
    metadata: dict | None = None,
) -> RepairLogEntry | None:
    ensure_project_tracking_ids(project)
    before = dict(before_state or capture_question_state(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    ))
    after = capture_question_state(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    )
    if before == after:
        return None
    entry = RepairLogEntry(
        timestamp=_utc_now(),
        source=source,
        action=action,
        question_id=after.get("question_id", ""),
        question_no=after.get("question_no", ""),
        section_kind=after.get("section_kind", "unknown"),
        material_id=after.get("material_id", ""),
        before_state=before,
        after_state=after,
        metadata=dict(metadata or {}),
    )
    project.repair_log.append(entry)
    return entry


def append_project_repair_log(
    project: ExamProject,
    *,
    source: str,
    action: str,
    metadata: dict | None = None,
) -> RepairLogEntry:
    ensure_project_tracking_ids(project)
    entry = RepairLogEntry(
        timestamp=_utc_now(),
        source=source,
        action=action,
        metadata=dict(metadata or {}),
    )
    project.repair_log.append(entry)
    return entry
