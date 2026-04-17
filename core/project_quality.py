from __future__ import annotations

from collections import Counter
from dataclasses import dataclass
from pathlib import Path
import re

try:
    import fitz
except Exception:  # pragma: no cover - optional dependency
    fitz = None

from core.subject_inference import infer_pdf_filename_profile, infer_subject_diagnostics
from domain.models import ExamProject, QuestionNode, ReviewIssue, SUBJECT_DISPLAY_NAMES, SubjectKind

_SEVERITY_WEIGHTS = {
    "info": 0.08,
    "warning": 0.18,
    "error": 0.34,
}

_FLAGGED_CONFIDENCE_THRESHOLD = 0.82
_DATA_STEM_ASK_MARKERS = (
    "根据上述资料",
    "根据以下资料",
    "根据所给资料",
    "根据材料",
    "下列说法正确的是",
    "下列说法错误的是",
    "以下哪项",
)
_VISUAL_CHOICE_MARKERS = (
    "坐标图",
    "图形",
    "示意图",
    "图中",
    "如下图",
    "下列图",
    "哪个图",
)
_INLINE_OPTION_MARKER = re.compile(r"(?<![A-Za-z0-9])([A-D])[\.．、:：]")


@dataclass
class ProjectQualitySummary:
    question_count: int = 0
    flagged_questions: int = 0
    severe_questions: int = 0
    source_defect_questions: int = 0
    total_issue_count: int = 0


def severity_rank(value: str) -> int:
    return {"error": 3, "warning": 2, "info": 1}.get((value or "").lower(), 0)


def question_max_severity(question: QuestionNode) -> str:
    issues = getattr(question, "review_issues", []) or []
    if not issues:
        return "none"
    return max((issue.severity for issue in issues), key=severity_rank)


def _add_issue(
    issues: list[ReviewIssue],
    seen_codes: set[str],
    code: str,
    title: str,
    detail: str = "",
    severity: str = "warning",
) -> None:
    if code in seen_codes:
        return
    seen_codes.add(code)
    issues.append(
        ReviewIssue(
            code=code,
            title=title,
            detail=detail,
            severity=severity,  # type: ignore[arg-type]
        )
    )


def _normalized_option_text(value: str) -> str:
    return "".join((value or "").split()).strip(".,，。:：;；")


def _question_expects_visual_choice(question: QuestionNode) -> bool:
    stem = (question.stem or "").strip()
    if not stem:
        return False
    return any(marker in stem for marker in _VISUAL_CHOICE_MARKERS)


def _question_number_token(number: int) -> re.Pattern[str]:
    return re.compile(rf"(?<!\d){number}\s*[、.．](?!\d)")


def _question_page_numbers(question: QuestionNode) -> list[int]:
    return [
        page_number
        for page_number in (getattr(question, "page_numbers", []) or [])
        if isinstance(page_number, int) and page_number > 0
    ]


def _page_text_map(pdf_path: str | Path | None, page_numbers: set[int]) -> dict[int, str]:
    if fitz is None or not pdf_path or not page_numbers:
        return {}
    path = Path(pdf_path)
    if not path.exists():
        return {}
    try:
        document = fitz.open(path)
    except Exception:
        return {}
    try:
        result: dict[int, str] = {}
        for page_number in sorted(page_numbers):
            if 1 <= page_number <= len(document):
                result[page_number] = document[page_number - 1].get_text("text")
        return result
    finally:
        document.close()


def _gap_numbers_come_from_source(
    pdf_path: str | Path | None,
    previous_question: QuestionNode,
    current_question: QuestionNode,
) -> bool:
    previous_numeric = previous_question.numeric_source_number
    current_numeric = current_question.numeric_source_number
    if previous_numeric is None or current_numeric is None or current_numeric <= previous_numeric + 1:
        return False

    missing_numbers = list(range(previous_numeric + 1, current_numeric))
    candidate_pages = set(_question_page_numbers(previous_question)) | set(_question_page_numbers(current_question))
    if not candidate_pages:
        return False

    page_texts = _page_text_map(pdf_path, candidate_pages)
    combined_text = "\n".join(page_texts.get(page_number, "") for page_number in sorted(candidate_pages))
    if not combined_text:
        return False

    if not _question_number_token(previous_numeric).search(combined_text):
        return False
    if not _question_number_token(current_numeric).search(combined_text):
        return False
    return all(not _question_number_token(number).search(combined_text) for number in missing_numbers)


def _looks_like_chapter_reset(section_kind: SubjectKind, previous_numeric: int, question: QuestionNode) -> bool:
    if section_kind not in {"politics", "common_sense"}:
        return False
    numeric = question.numeric_source_number
    if numeric is None or numeric >= previous_numeric or previous_numeric < 2:
        return False
    stem = (question.stem or "").strip()
    if len(stem) < 12:
        return False
    looks_like_fresh_question = stem.startswith("(") or stem.startswith("（") or "·" in stem[:24]
    if numeric == 1:
        return True
    if previous_numeric - numeric >= 5 and numeric <= 10:
        return True
    return looks_like_fresh_question


def _looks_like_empty_source_placeholder(
    section_kind: SubjectKind,
    sequential_questions: list[QuestionNode],
    index: int,
) -> bool:
    if section_kind == "data":
        return False
    question = sequential_questions[index]
    if (question.stem or "").strip():
        return False
    if question.stem_assets or question.options:
        return False
    current = question.numeric_source_number
    if current is None:
        return False
    previous = sequential_questions[index - 1] if index > 0 else None
    nxt = sequential_questions[index + 1] if index + 1 < len(sequential_questions) else None
    if previous is None or nxt is None:
        return False
    previous_numeric = previous.numeric_source_number
    next_numeric = nxt.numeric_source_number
    if previous_numeric != current - 1 or next_numeric != current + 1:
        return False
    return any(
        (
            (candidate.stem or "").strip()
            or candidate.stem_assets
            or candidate.options
        )
        for candidate in (previous, nxt)
    )


def _looks_like_partial_inline_option_loss(
    section_kind: SubjectKind,
    sequential_questions: list[QuestionNode],
    index: int,
) -> bool:
    if section_kind == "data":
        return False
    question = sequential_questions[index]
    stem = (question.stem or "").strip()
    if len(stem) < 40 or question.stem_assets or question.options:
        return False

    current = question.numeric_source_number
    nxt = sequential_questions[index + 1] if index + 1 < len(sequential_questions) else None
    if current is None or nxt is None or nxt.numeric_source_number != current + 1:
        return False
    if not ((nxt.stem or "").strip() or nxt.stem_assets or nxt.options):
        return False

    marker_matches = list(_INLINE_OPTION_MARKER.finditer(stem))
    inline_letters: list[str] = []
    for match in marker_matches:
        letter = match.group(1)
        if not inline_letters or inline_letters[-1] != letter:
            inline_letters.append(letter)

    if len(inline_letters) < 2 or len(inline_letters) >= 4:
        return False
    expected_prefix = list("ABCD")[: len(inline_letters)]
    if inline_letters != expected_prefix:
        return False

    if marker_matches:
        first_marker_pos = marker_matches[0].start()
        prompt_window = stem[max(0, first_marker_pos - 12):first_marker_pos]
        has_prompt_boundary = any(token in prompt_window for token in ("()", "）", ")"))
        if not has_prompt_boundary and first_marker_pos < max(16, len(stem) // 3):
            return False

    return True


def _pages_have_visual_candidates(pdf_path: str | None, page_numbers: list[int]) -> bool | None:
    if fitz is None or not pdf_path or not page_numbers:
        return None
    path = Path(pdf_path)
    if not path.exists():
        return None
    try:
        document = fitz.open(path)
    except Exception:
        return None
    try:
        for page_number in page_numbers:
            if page_number < 1 or page_number > len(document):
                continue
            page = document[page_number - 1]
            for image in page.get_images(full=True):
                for rect in page.get_image_rects(image[0], transform=False):
                    if rect.y1 <= 80 or rect.get_area() < 1200:
                        continue
                    return True
            for drawing in page.get_drawings():
                rect = drawing.get("rect")
                if rect is None or rect.y1 <= 80:
                    continue
                if rect.get_area() < 600:
                    continue
                fill = drawing.get("fill")
                color = drawing.get("color")
                width = drawing.get("width")
                if fill == (1.0, 1.0, 1.0) and color in (None, (1.0, 1.0, 1.0)) and not width:
                    continue
                return True
    finally:
        document.close()
    return False


def _looks_like_embedded_data_intro(material: MaterialSet | None, question: QuestionNode) -> bool:
    if material is None:
        return False
    if (material.body or "").strip():
        return False
    stem = (question.stem or "").strip()
    if len(stem) < 36:
        return False
    ask_index = max((stem.find(marker) for marker in _DATA_STEM_ASK_MARKERS), default=-1)
    if ask_index <= 12:
        return False
    intro = stem[:ask_index]
    return any(token in intro for token in ("同比", "环比", "增长率", "表", "图", "资料", "材料", "比重", "投资", "人数", "企业", "亿元"))


def _looks_like_data_stem_assets_belong_to_material(material: MaterialSet | None, question: QuestionNode) -> bool:
    if material is None:
        return False
    if not question.stem_assets:
        return False
    if material.questions and material.questions[0] is not question:
        return False

    stem = (question.stem or "").strip()
    if not stem or len(stem) > 38:
        return False
    if not any(marker in stem for marker in _DATA_STEM_ASK_MARKERS):
        return False
    if any(option.image_path for option in question.options):
        return False

    image_pages = {asset.source_page for asset in question.stem_assets if asset.source_page is not None}
    if len(question.stem_assets) == 1 and (material.body or "").strip() and not image_pages:
        return False
    return True


def _guess_subject_reason(
    current_kind: SubjectKind,
    inferred_kind: SubjectKind,
    confidence: float,
    subtype: str | None = None,
    matched_signals: tuple[str, ...] = (),
) -> str:
    inferred_label = SUBJECT_DISPLAY_NAMES.get(inferred_kind, inferred_kind)
    current_label = SUBJECT_DISPLAY_NAMES.get(current_kind, current_kind)
    inferred_with_subtype = f"{inferred_label} / {subtype}" if subtype else inferred_label
    confidence_text = f"本地置信度 {int(round(max(0.0, min(confidence, 1.0)) * 100))}%。"
    signal_text = ""
    if matched_signals:
        signal_text = "命中线索：" + "、".join(matched_signals[:3]) + "。"
    if current_kind == "unknown":
        return f"{signal_text}{confidence_text}当前未能稳定归类，这道题更像 {inferred_with_subtype}，建议优先人工确认。"
    return f"{signal_text}{confidence_text}按题干和选项特征，这道题更像 {inferred_with_subtype}，与当前科目 {current_label} 不一致。"


def _question_review_confidence(issues: list[ReviewIssue], suggested_subject: SubjectKind | None) -> float:
    penalty = sum(_SEVERITY_WEIGHTS.get(issue.severity, 0.12) for issue in issues)
    if suggested_subject is not None:
        penalty += 0.08
    return max(0.05, min(1.0, 1.0 - penalty))


def _subject_suggestion_threshold(current_kind: SubjectKind, inferred_kind: SubjectKind) -> float:
    if current_kind == "unknown":
        return 0.62
    if {current_kind, inferred_kind} == {"politics", "common_sense"}:
        return 0.82
    if {current_kind, inferred_kind} == {"common_sense", "reasoning"}:
        return 0.76
    return 0.7


def _subject_mismatch_enabled(project: ExamProject, section_kind: SubjectKind) -> bool:
    if section_kind in {"unknown", "data"}:
        return True
    pdf_path = getattr(project.source, "pdf_path", None)
    if not pdf_path:
        return True
    profile = infer_pdf_filename_profile(pdf_path)
    if (
        profile.form == "single_subject_book"
        and profile.subject_hint == section_kind
        and profile.confidence >= 0.82
    ):
        return False
    return True


def is_flagged_question(question: QuestionNode) -> bool:
    if question.review_issues:
        return True
    return float(getattr(question, "review_confidence", 1.0) or 1.0) < _FLAGGED_CONFIDENCE_THRESHOLD


def question_review_summary(question: QuestionNode) -> str:
    issues = getattr(question, "review_issues", []) or []
    if not issues:
        return "结构稳定"
    return "；".join(issue.title for issue in issues[:3])


def iter_flagged_question_rows(project: ExamProject):
    for section, material, question in project.iter_questions():
        if is_flagged_question(question):
            yield section, material, question


def annotate_project_quality(project: ExamProject) -> ProjectQualitySummary:
    summary = ProjectQualitySummary()
    question_rows = list(project.iter_questions())
    summary.question_count = len(question_rows)
    scheduled_issues: dict[int, list[tuple[str, str, str, str]]] = {}
    pdf_path = getattr(project.source, "pdf_path", None)

    duplicate_numbers = Counter()
    all_number_questions: dict[str, list[QuestionNode]] = {}
    for _section, _material, question in question_rows:
        number = (question.source_number or "").strip()
        if not number:
            continue
        duplicate_numbers[number] += 1
        all_number_questions.setdefault(number, []).append(question)

    for section in project.sections:
        sequential_questions: list[QuestionNode] = []
        if section.kind == "data":
            for material in section.material_sets:
                sequential_questions.extend(material.questions)
        else:
            sequential_questions.extend(section.questions)

        empty_source_placeholders = {
            id(question)
            for index, question in enumerate(sequential_questions)
            if _looks_like_empty_source_placeholder(section.kind, sequential_questions, index)
        }
        partial_inline_source_losses = {
            id(question)
            for index, question in enumerate(sequential_questions)
            if _looks_like_partial_inline_option_loss(section.kind, sequential_questions, index)
        }

        previous_numeric: int | None = None
        previous_question: QuestionNode | None = None
        for question in sequential_questions:
            numeric = question.numeric_source_number
            if numeric is None:
                continue
            if previous_numeric is not None:
                if numeric == previous_numeric:
                    scheduled_issues.setdefault(id(question), []).append(
                        (
                            "duplicate_number",
                            "题号重复",
                            f"这道题与前一道题都标成了 {numeric}。",
                            "error",
                        )
                    )
                    if previous_question is not None:
                        scheduled_issues.setdefault(id(previous_question), []).append(
                            (
                                "duplicate_number",
                                "题号重复",
                                f"这道题与后一道题都标成了 {numeric}。",
                                "error",
                            )
                        )
                elif numeric < previous_numeric:
                    if _looks_like_chapter_reset(section.kind, previous_numeric, question):
                        previous_numeric = numeric
                        previous_question = question
                        continue
                    scheduled_issues.setdefault(id(question), []).append(
                        (
                            "number_order",
                            "题序倒退",
                            f"当前题号 {numeric} 小于上一题 {previous_numeric}，通常说明切题顺序有问题。",
                            "error",
                        )
                    )
                elif numeric > previous_numeric + 1:
                    if previous_question is not None and _gap_numbers_come_from_source(pdf_path, previous_question, question):
                        previous_numeric = numeric
                        previous_question = question
                        continue
                    scheduled_issues.setdefault(id(question), []).append(
                        (
                            "number_gap",
                            "题号跳跃",
                            f"当前题号从 {previous_numeric} 跳到了 {numeric}，中间可能有漏题。",
                            "warning",
                        )
                    )
            previous_numeric = numeric
            previous_question = question

        for question in sequential_questions:
            if id(question) in empty_source_placeholders:
                scheduled_issues.setdefault(id(question), []).append(
                    (
                        "source_text_missing",
                        "源 PDF 题目文本缺失",
                        "当前题只保留了题号，占位内容为空；结合前后连续题号判断，更像源文件当前页原题内容缺失。",
                        "error",
                    )
                )
            if id(question) in partial_inline_source_losses:
                scheduled_issues.setdefault(id(question), []).append(
                    (
                        "source_text_missing",
                        "源 PDF 题目文本缺失",
                        "当前题干尾部只残留了前半组选项，随后已直接切到下一题；更像源 PDF 当前页后半题干或选项缺失。",
                        "error",
                    )
                )

    for section, material, question in question_rows:
        issues: list[ReviewIssue] = []
        seen_codes: set[str] = set()
        question.review_issues = []
        question.suggested_subject = None
        question.suggested_subject_confidence = None
        question.suggested_subject_reason = ""
        question.inferred_subtype = ""
        question.inferred_subtype_confidence = None
        question.inferred_signals = []
        for code, title, detail, severity in scheduled_issues.get(id(question), []):
            _add_issue(issues, seen_codes, code, title, detail, severity)

        stem = (question.stem or "").strip()
        image_option_count = sum(1 for option in question.options if option.image_path)
        option_texts = [option.text or "" for option in question.options]
        normalized_option_texts = [_normalized_option_text(text) for text in option_texts if _normalized_option_text(text)]

        if "source_text_missing" in seen_codes:
            pass
        elif not stem and not question.stem_assets:
            _add_issue(
                issues,
                seen_codes,
                "missing_stem",
                "题干为空",
                "当前题没有可用题干文本，也没有题干图片。",
                severity="error",
            )
        elif stem in {"缺失", "题目缺失"}:
            _add_issue(
                issues,
                seen_codes,
                "source_text_missing",
                "源 PDF 题目文本缺失",
                "源文件当前页只保留了“缺失”占位文本，疑似题目原文在 PDF 中已缺失或 OCR 完全损坏。",
                severity="error",
                )

        if (
            section.kind != "data"
            and not question.stem_assets
            and not any(option.image_path for option in question.options)
            and len(question.options) <= 1
            and _question_expects_visual_choice(question)
        ):
            visual_candidates = _pages_have_visual_candidates(
                getattr(project.source, "pdf_path", None),
                list(getattr(question, "page_numbers", []) or []),
            )
            if visual_candidates is False:
                _add_issue(
                    issues,
                    seen_codes,
                    "source_visual_missing",
                    "源 PDF 疑似缺少图形选项",
                    "当前题干明确要求根据图形或坐标图判断，但对应 PDF 页没有可用图形对象，疑似源文件本身缺图。",
                    severity="error",
                )

        source_missing_option_texts = [
            option.letter
            for option in question.options
            if _normalized_option_text(option.text) in {"缺失", "题目缺失"}
        ]
        if source_missing_option_texts:
            _add_issue(
                issues,
                seen_codes,
                "source_text_missing",
                "源 PDF 题目文本缺失",
                "、".join(source_missing_option_texts) + " 选项在源文件中只保留了“缺失”占位文本，更像原题文字缺失而非切题错误。",
                severity="error",
            )

        if "source_text_missing" not in seen_codes and "source_visual_missing" not in seen_codes and section.kind != "unknown" and len(question.options) not in {4}:
            _add_issue(
                issues,
                seen_codes,
                "option_count",
                "选项数量异常",
                f"当前识别到 {len(question.options)} 个选项，通常应为 4 个。",
                severity="error" if len(question.options) <= 1 else "warning",
            )

        blank_letters = [
            option.letter
            for option in question.options
            if not (option.text or "").strip() and not option.image_path
        ]
        if blank_letters:
            filled_option_count = sum(
                1
                for option in question.options
                if (option.text or "").strip() or option.image_path
            )
            if (
                "source_text_missing" not in seen_codes
                and len(blank_letters) == 1
                and filled_option_count >= 3
                and (image_option_count >= 2 or (not stem and bool(question.stem_assets)))
            ):
                _add_issue(
                    issues,
                    seen_codes,
                    "source_text_missing",
                    "源 PDF 题目文本缺失",
                    "当前仅有一个选项完全空白，其余选项已恢复为图像或文本，更像源文件当前页该选项原文缺失。",
                    severity="error",
                )
            if "source_text_missing" in seen_codes:
                blank_letters = []
        if blank_letters:
            filled_option_count = sum(
                1
                for option in question.options
                if (option.text or "").strip() or option.image_path
            )
            _add_issue(
                issues,
                seen_codes,
                "blank_option",
                "存在空白选项",
                "、".join(blank_letters) + " 选项没有文字，也没有图片。",
                severity="warning" if len(blank_letters) == 1 and filled_option_count >= 3 else "error",
            )

        option_counter = Counter(normalized_option_texts)
        repeated_options = [value for value, count in option_counter.items() if count > 1]
        if repeated_options and "source_text_missing" not in seen_codes:
            _add_issue(
                issues,
                seen_codes,
                "duplicate_option_text",
                "选项文字重复",
                "发现重复选项内容，建议人工检查是否切分串题。",
            )

        if section.kind == "unknown":
            _add_issue(
                issues,
                seen_codes,
                "unknown_subject",
                "科目待确认",
                "当前篇题没有稳定归类，建议优先人工确认科目。",
            )

        if material is not None and not ((material.body or "").strip() or material.body_assets):
            _add_issue(
                issues,
                seen_codes,
                "material_empty",
                "材料正文可能缺失",
                "当前材料没有正文文本，也没有材料图片，建议回到 PDF 原文复核。",
            )
        if _looks_like_embedded_data_intro(material, question):
            _add_issue(
                issues,
                seen_codes,
                "material_intro_embedded_in_stem",
                "材料说明可能挂进了首题",
                "当前材料正文为空，但首题题干前半段像材料说明，建议把它移回材料区。",
            )
        if _looks_like_data_stem_assets_belong_to_material(material, question):
            _add_issue(
                issues,
                seen_codes,
                "material_asset_binding",
                "材料图表可能挂进了首题",
                "当前像是资料分析的材料图表，但图片挂在首题题干上，建议移回材料区。",
            )

        diagnostics = infer_subject_diagnostics(
            stem=stem,
            options=option_texts,
            material_text=(material.body or "") if material is not None else "",
            image_count=len(question.stem_assets) + image_option_count,
            material_header=(material.header or "") if material is not None else "",
            allow_data=True,
        )
        inferred_kind = diagnostics.kind
        inferred_strength = diagnostics.margin
        inferred_confidence = diagnostics.confidence
        question.inferred_subtype = diagnostics.subtype or ""
        question.inferred_subtype_confidence = inferred_confidence if diagnostics.subtype else None
        question.inferred_signals = list(diagnostics.matched_signals or [])

        if (
            question.inferred_subtype == "图形推理"
            and len(question.stem_assets) == 4
            and len(question.options) >= 4
            and not any(option.image_path for option in question.options[:4])
        ):
            short_option_count = sum(
                1
                for option in question.options[:4]
                if len((option.text or "").strip().strip("A.B.C.D．、:： ")) <= 2
            )
            if short_option_count >= 4:
                _add_issue(
                    issues,
                    seen_codes,
                    "graphic_asset_binding",
                    "图形题图片可能挂错位置",
                    "当前像是图形推理，但 4 张图片还挂在题干上，A-D 选项本身没有图片。",
                )

        threshold = _subject_suggestion_threshold(section.kind, inferred_kind)
        if (
            section.kind == "unknown"
            and inferred_kind != "unknown"
            and inferred_strength >= 0.75
            and inferred_confidence >= threshold
        ):
            question.suggested_subject = inferred_kind
            question.suggested_subject_confidence = inferred_confidence
            question.suggested_subject_reason = _guess_subject_reason(
                section.kind,
                inferred_kind,
                inferred_confidence,
                diagnostics.subtype,
                diagnostics.matched_signals,
            )
            _add_issue(
                issues,
                seen_codes,
                "subject_suggestion",
                f"建议改为 {SUBJECT_DISPLAY_NAMES.get(inferred_kind, inferred_kind)}",
                question.suggested_subject_reason,
            )
        elif (
            section.kind not in {"unknown", "data"}
            and inferred_kind not in {"unknown", section.kind}
            and inferred_strength >= 1.05
            and inferred_confidence >= threshold
        ):
            question.suggested_subject = inferred_kind
            question.suggested_subject_confidence = inferred_confidence
            question.suggested_subject_reason = _guess_subject_reason(
                section.kind,
                inferred_kind,
                inferred_confidence,
                diagnostics.subtype,
                diagnostics.matched_signals,
            )
            _add_issue(
                issues,
                seen_codes,
                "subject_mismatch",
                "题型可能分错科目",
                question.suggested_subject_reason,
            )

        question.review_issues = issues
        question.review_confidence = _question_review_confidence(issues, question.suggested_subject)

        if is_flagged_question(question):
            summary.flagged_questions += 1
        if any(issue.severity == "error" for issue in issues):
            summary.severe_questions += 1
        if any(issue.code.startswith("source_") for issue in issues):
            summary.source_defect_questions += 1
        summary.total_issue_count += len(issues)

    return summary
