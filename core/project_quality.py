from __future__ import annotations

from collections import Counter
from dataclasses import dataclass

from core.subject_inference import infer_subject_diagnostics
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


@dataclass
class ProjectQualitySummary:
    question_count: int = 0
    flagged_questions: int = 0
    severe_questions: int = 0
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
                    scheduled_issues.setdefault(id(question), []).append(
                        (
                            "number_order",
                            "题序倒退",
                            f"当前题号 {numeric} 小于上一题 {previous_numeric}，通常说明切题顺序有问题。",
                            "error",
                        )
                    )
                elif numeric > previous_numeric + 1:
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

        if not stem and not question.stem_assets:
            _add_issue(
                issues,
                seen_codes,
                "missing_stem",
                "题干为空",
                "当前题没有可用题干文本，也没有题干图片。",
                severity="error",
            )

        if section.kind != "unknown" and len(question.options) not in {4}:
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
            _add_issue(
                issues,
                seen_codes,
                "blank_option",
                "存在空白选项",
                "、".join(blank_letters) + " 选项没有文字，也没有图片。",
                severity="error",
            )

        normalized_options = [
            _normalized_option_text(option.text)
            for option in question.options
            if _normalized_option_text(option.text)
        ]
        option_counter = Counter(normalized_options)
        repeated_options = [value for value, count in option_counter.items() if count > 1]
        if repeated_options:
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
        summary.total_issue_count += len(issues)

    return summary
