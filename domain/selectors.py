from __future__ import annotations

import copy
from typing import Iterable, Optional

from domain.models import ALL_SUBJECT_KINDS, ExamProject, MaterialSet, QuestionNode, QuestionRange, Section, SubjectKind

_SUBJECT_ALIAS_TO_KIND: dict[str, SubjectKind] = {
    "politics": "politics",
    "政治理论": "politics",
    "common_sense": "common_sense",
    "common-sense": "common_sense",
    "常识判断": "common_sense",
    "verbal": "verbal",
    "言语理解与表达": "verbal",
    "言语理解和表达": "verbal",
    "quant": "quant",
    "数量关系": "quant",
    "reasoning": "reasoning",
    "判断推理": "reasoning",
    "data": "data",
    "资料分析": "data",
}


def parse_question_ranges(spec: str) -> list[QuestionRange]:
    ranges: list[QuestionRange] = []
    for chunk in (spec or "").split(","):
        item = chunk.strip()
        if not item:
            continue
        if "-" in item:
            start_text, end_text = item.split("-", 1)
            if start_text.strip().isdigit() and end_text.strip().isdigit():
                start = int(start_text.strip())
                end = int(end_text.strip())
                if start <= end:
                    ranges.append(QuestionRange(start=start, end=end))
            continue
        if item.isdigit():
            value = int(item)
            ranges.append(QuestionRange(start=value, end=value))
    return ranges


def _matches_ranges(question: QuestionNode, ranges: Iterable[QuestionRange]) -> bool:
    number = question.numeric_source_number
    if number is None:
        return False
    return any(question_range.contains(number) for question_range in ranges)


def parse_subject_spec(spec: str | Iterable[str] | None) -> list[SubjectKind]:
    if spec is None:
        return list(ALL_SUBJECT_KINDS)

    if isinstance(spec, str):
        raw_parts = [
            part.strip()
            for chunk in spec.replace("，", ",").replace("、", ",").split(",")
            for part in [chunk]
            if part.strip()
        ]
    else:
        raw_parts = [str(part).strip() for part in spec if str(part).strip()]

    if not raw_parts:
        return list(ALL_SUBJECT_KINDS)

    selected: list[SubjectKind] = []
    for raw in raw_parts:
        token = raw.lower()
        if token in ("all", "*"):
            return list(ALL_SUBJECT_KINDS)
        if token == "both":
            for kind in ("quant", "data"):
                if kind not in selected:
                    selected.append(kind)
            continue
        kind = _SUBJECT_ALIAS_TO_KIND.get(token) or _SUBJECT_ALIAS_TO_KIND.get(raw)
        if kind and kind not in selected:
            selected.append(kind)

    return selected or list(ALL_SUBJECT_KINDS)


def select_project(
    project: ExamProject,
    subjects: Optional[Iterable[SubjectKind]] = None,
    question_ranges: Optional[Iterable[QuestionRange]] = None,
) -> ExamProject:
    selected_subjects = [subject for subject in (subjects or []) if subject]
    selected_ranges = list(question_ranges or [])
    clone = copy.deepcopy(project)
    clone.selected_subjects = selected_subjects
    clone.selected_ranges = selected_ranges

    filtered_sections: list[Section] = []
    for section in clone.sections:
        if selected_subjects and section.kind not in selected_subjects and section.kind != "unknown":
            continue
        if section.kind == "data":
            material_sets: list[MaterialSet] = []
            for material in section.material_sets:
                if selected_ranges:
                    material.questions = [
                        question
                        for question in material.questions
                        if _matches_ranges(question, selected_ranges)
                    ]
                if material.questions:
                    material_sets.append(material)
            section.material_sets = material_sets
            if section.material_sets:
                filtered_sections.append(section)
        else:
            if selected_ranges:
                section.questions = [
                    question
                    for question in section.questions
                    if _matches_ranges(question, selected_ranges)
                ]
            if section.questions:
                filtered_sections.append(section)

    clone.sections = filtered_sections
    return clone
