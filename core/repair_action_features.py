from __future__ import annotations

from typing import Any, Iterable

from domain.models import MaterialSet, QuestionNode, Section


def _clean_text(parts: Iterable[str]) -> str:
    return "\n".join(part.strip() for part in parts if (part or "").strip()).strip()


def _material_text(material: MaterialSet | None) -> str:
    if material is None:
        return ""
    body = (material.body or "").strip()
    if body:
        return body
    return _clean_text(material.body_lines)


def _question_blob(question: QuestionNode | None) -> str:
    if question is None:
        return ""
    option_lines = [
        f"{option.letter}. {option.text}".strip()
        for option in question.options
        if (option.text or "").strip() or option.image_path
    ]
    stem = (question.stem or "").strip() or "[空题干]"
    return _clean_text(
        [
            f"题号：{question.source_number or '-'}",
            f"题干：{stem}",
            *option_lines,
            f"题干图片数：{len(question.stem_assets)}",
        ]
    )


def build_repair_state_record(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
) -> dict[str, Any]:
    stem_text = (question.stem or "").strip()
    option_texts = [(option.text or "").strip() for option in question.options]
    material_text = _material_text(material)
    previous_option_count = float(len(previous_question.options)) if previous_question is not None else 0.0
    previous_missing_option_count = (
        float(max(0, 4 - len(previous_question.options))) if previous_question is not None else 0.0
    )
    next_stem = (next_question.stem or "").strip() if next_question is not None else ""
    previous_number = previous_question.numeric_source_number if previous_question is not None else None
    current_number = question.numeric_source_number
    gap_from_previous = 0.0
    if previous_number is not None and current_number is not None:
        gap_from_previous = float(current_number - previous_number)
    text = _clean_text(
        [
            f"[SECTION] {section.kind} / {section.title}",
            f"[MATERIAL] {material_text}" if material is not None else "",
            f"[PREV] {_question_blob(previous_question)}" if previous_question is not None else "",
            f"[CURRENT] {_question_blob(question)}",
            f"[NEXT] {_question_blob(next_question)}" if next_question is not None else "",
        ]
    )
    option_image_count = sum(1 for option in question.options if option.image_path)
    blank_option_count = sum(1 for text in option_texts if not text)
    return {
        "text": text,
        "meta": {
            "section_kind": section.kind,
            "has_material": 1.0 if material is not None else 0.0,
            "material_length": float(len(material_text)),
            "material_asset_count": float(len(material.body_assets)) if material is not None else 0.0,
            "question_number": float(question.numeric_source_number or 0),
            "option_count": float(len(question.options)),
            "stem_asset_count": float(len(question.stem_assets)),
            "option_image_count": float(option_image_count),
            "blank_option_count": float(blank_option_count),
            "all_options_are_images": 1.0 if question.options and option_image_count == len(question.options) else 0.0,
            "stem_length": float(len(stem_text)),
            "stem_has_embedded_question_no": 1.0 if _question_number_like(stem_text) else 0.0,
            "stem_starts_with_option_marker": 1.0 if _leading_option_like(stem_text) else 0.0,
            "stem_has_data_prompt": 1.0 if _data_prompt_like(stem_text) else 0.0,
            "prev_present": 1.0 if previous_question is not None else 0.0,
            "previous_option_count": previous_option_count,
            "previous_missing_option_count": previous_missing_option_count,
            "next_present": 1.0 if next_question is not None else 0.0,
            "next_stem_empty": 1.0 if next_question is not None and not next_stem else 0.0,
            "number_gap_from_previous": gap_from_previous,
        },
    }


def _question_number_like(text: str) -> bool:
    for index, char in enumerate(text or ""):
        if char.isdigit():
            tail = (text or "")[index : index + 6]
            if any(marker in tail for marker in (".", "．", "、", ":", "：")):
                return True
    return False


def _leading_option_like(text: str) -> bool:
    normalized = (text or "").lstrip()
    if len(normalized) < 2:
        return False
    return normalized[0] in {"A", "B", "C", "D"} and normalized[1] in {".", "．", "、", ":", "："}


def _data_prompt_like(text: str) -> bool:
    return any(
        marker in (text or "")
        for marker in (
            "根据上述资料",
            "根据以下资料",
            "根据所给资料",
            "根据材料",
            "下列说法正确的是",
            "下列说法错误的是",
        )
    )
