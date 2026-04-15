from __future__ import annotations

from dataclasses import dataclass, field
import os
import re
from typing import Literal, Optional

from core.project_quality import annotate_project_quality, is_flagged_question, question_review_summary
from core.repair_action_model import RepairActionPrediction, predict_repair_action_distribution
from core.repair_trajectory_model import TrajectoryPlan, build_rule_based_trajectory_plan, predict_trajectory_plan
from core.subject_inference import infer_subject_diagnostics
from domain.models import ExamProject, MaterialSet, OptionNode, QuestionNode, Section, SUBJECT_DISPLAY_NAMES
from domain.project_editor import (
    insert_question_before,
    insert_option_after,
    move_stem_assets_to_material,
    prepend_material_body_text,
    reassign_stem_assets_to_options,
    renumber_question,
    reclassify_objective_section,
    set_question_option_layout,
    update_option_text,
    update_question_stem,
)

_DEFAULT_BATCH_LIMIT = 12
_SUPPORTED_LAYOUTS = {"one_row", "grid", "list"}
_FILL_BLANK_PREFIXES = (",", "，", "。", "；", ";", ":", "：")

_INLINE_OPTION_PATTERN = re.compile(
    r"(?:(?<=^)|(?<=[\s\u3000]))([A-D])(?:[\.．、:：]|(?=\s))\s*"
)
_LEADING_OPTION_PATTERN = re.compile(r"^\s*([A-D])(?:[\.．、:：])\s*(.+)$")
_LEADING_QUESTION_PATTERNS = (
    re.compile(r"^\s*(?P<number>\d{1,3})\s*[\.．、)\uFF09:：]\s*(?P<stem>.+)$"),
    re.compile(r"^\s*第\s*(?P<number>\d{1,3})\s*题(?:\s*[\.．、)\uFF09:：]\s*)?(?P<stem>.+)$"),
    re.compile(r"^\s*(?P<number>\d{2,3})\s+(?P<stem>.+)$"),
)
_PROMPT_START_MARKERS = (
    "请根据",
    "根据以下",
    "根据上述",
    "根据下列",
    "请阅读以下",
    "阅读以下",
    "请开始答题",
)
_PROMPT_CONTEXT_MARKERS = ("材料", "资料", "图表", "文字", "信息", "内容")
_PROMPT_ACTION_MARKERS = ("回答", "作答", "解答")
_DEFINITION_ASK_MARKERS = ("下列", "以下", "其中", "请问", "哪一项")
_SENTENCE_ORDER_TAIL_MARKERS = (
    "将以上",
    "将上述",
    "把以上",
    "把上述",
    "重新排列",
    "重新排序",
    "语序最",
    "排序最",
    "衔接最",
)
_ENUMERATION_MARKER_RE = re.compile(r"([①②③④⑤⑥⑦⑧⑨⑩])")
_PAREN_ENUMERATION_MARKER_RE = re.compile(r"([（(]\s*\d+\s*[）)])")
_READING_ASK_MARKERS = (
    "这段文字意在说明",
    "这段文字主要说明",
    "这段文字主要讲述",
    "这段文字主要介绍",
    "作者意在",
    "文段意在说明",
    "文段主要说明",
    "最适合做这段文字标题的是",
    "最适合做上述文字标题的是",
    "对这段文字概括最准确的是",
)
_DATA_ASK_MARKERS = (
    "根据上述资料",
    "根据以下资料",
    "根据所给资料",
    "根据材料",
    "下列说法正确的是",
    "下列说法错误的是",
    "能够推出的是",
    "不能推出的是",
    "以下哪项",
)
_REPAIR_ACTION_DISPLAY_NAMES = {
    "move_data_assets_to_material": "图表回挂材料区",
    "move_data_intro_back_to_material": "材料说明回挂材料区",
    "move_spilled_option_back": "串出的选项归位",
    "reassign_stem_image_to_options": "图片选项归位",
    "renumber_current_question": "题号回正",
    "split_embedded_next_question": "拆出嵌入的下一题",
}
_REPAIR_ACTION_FAMILY_DISPLAY_NAMES = {
    "asset_repair": "图片归位",
    "boundary_repair": "边界修复",
    "material_repair": "材料修复",
}
_TRAJECTORY_ACTION_DISPLAY = _REPAIR_ACTION_DISPLAY_NAMES


@dataclass
class AIOptionPatch:
    letter: str
    text: Optional[str] = None


@dataclass
class AIQuestionPatch:
    should_apply: bool = False
    confidence: float = 0.0
    summary: str = ""
    source_number: Optional[str] = None
    stem: Optional[str] = None
    option_layout: Optional[Literal["one_row", "grid", "list"]] = None
    subject_kind: Optional[Literal["politics", "common_sense", "verbal", "quant", "reasoning", "data", "unknown"]] = None
    reassign_stem_assets_to_options: bool = False
    move_stem_assets_to_material: bool = False
    prepend_material_body: Optional[str] = None
    options: list[AIOptionPatch] = field(default_factory=list)


@dataclass
class AIRepairResult:
    question_number: str
    patch: AIQuestionPatch
    applied_changes: int = 0
    applied_subject_change: bool = False


@dataclass
class AIRepairBatchSummary:
    attempted_questions: int = 0
    changed_questions: int = 0
    total_field_changes: int = 0
    subject_changes: int = 0
    structural_repairs: int = 0
    skipped_questions: int = 0
    errors: list[str] = field(default_factory=list)


@dataclass(frozen=True)
class RepairStrategySnapshot:
    mode: str
    action_prediction: RepairActionPrediction | None = None
    family_prediction: RepairActionPrediction | None = None
    trajectory_plan: TrajectoryPlan | None = None
    priority_score: float = 0.0


def _patch_action_candidates(patch: AIQuestionPatch) -> tuple[set[str], set[str]]:
    actions: set[str] = set()
    families: set[str] = set()
    if (patch.source_number or "").strip():
        actions.add("renumber_current_question")
        families.add("boundary_repair")
    if patch.prepend_material_body:
        actions.add("move_data_intro_back_to_material")
        families.add("material_repair")
    if patch.move_stem_assets_to_material:
        actions.add("move_data_assets_to_material")
        families.add("material_repair")
    if patch.reassign_stem_assets_to_options:
        actions.add("reassign_stem_image_to_options")
        families.add("asset_repair")
    return actions, families


def _format_policy_prediction(prediction: RepairActionPrediction | None) -> str:
    if prediction is None or not prediction.best_label:
        return ""
    if prediction.label_field == "action_family":
        label = _REPAIR_ACTION_FAMILY_DISPLAY_NAMES.get(prediction.best_label, prediction.best_label)
    else:
        label = _REPAIR_ACTION_DISPLAY_NAMES.get(prediction.best_label, prediction.best_label)
    confidence = int(round(max(0.0, min(prediction.best_confidence, 1.0)) * 100))
    return f"{label} {confidence}%"


def _format_trajectory_plan(plan: TrajectoryPlan | None) -> str:
    if plan is None or plan.is_empty:
        return ""
    parts: list[str] = []
    for step in plan.steps[:3]:
        label = _TRAJECTORY_ACTION_DISPLAY.get(step.action, step.action)
        conf = int(round(max(0.0, min(step.confidence, 1.0)) * 100))
        parts.append(f"{label}({conf}%)")
    return "→".join(parts)


def inspect_repair_strategy(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
    mode: str = "balanced",
) -> RepairStrategySnapshot:
    service = AIRepairService(mode=mode)
    trajectory_plan = service._trajectory_plan(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    )
    if service.mode != "policy":
        priority_score = 0.0
        if trajectory_plan is not None and not trajectory_plan.is_empty:
            priority_score = trajectory_plan.primary_confidence
        return RepairStrategySnapshot(
            mode=service.mode,
            trajectory_plan=trajectory_plan,
            priority_score=priority_score,
        )

    action_prediction, family_prediction = service._policy_predictions(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    )
    priority_score = service._policy_priority_score(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    )
    return RepairStrategySnapshot(
        mode=service.mode,
        action_prediction=action_prediction,
        family_prediction=family_prediction,
        trajectory_plan=trajectory_plan,
        priority_score=priority_score,
    )


def build_repair_strategy_summary(snapshot: RepairStrategySnapshot | None) -> str:
    if snapshot is None:
        return "修复策略会在这里显示。"
    mode_label = "策略增强" if snapshot.mode == "policy" else "规则优先"
    lines = [f"当前模式：{mode_label}"]
    if snapshot.action_prediction is not None and snapshot.action_prediction.best_label:
        lines.append(f"动作模型：{_format_policy_prediction(snapshot.action_prediction)}")
    if snapshot.family_prediction is not None and snapshot.family_prediction.best_label:
        lines.append(f"修复类型：{_format_policy_prediction(snapshot.family_prediction)}")
    if snapshot.trajectory_plan is not None and not snapshot.trajectory_plan.is_empty:
        prefix = "轨迹策略" if snapshot.mode == "policy" else "规则轨迹"
        lines.append(f"{prefix}：{_format_trajectory_plan(snapshot.trajectory_plan)}")
    else:
        prefix = "轨迹策略" if snapshot.mode == "policy" else "规则轨迹"
        lines.append(f"{prefix}：当前没有明显的多步修复线索")
    if snapshot.priority_score > 0:
        lines.append(f"优先级分：{snapshot.priority_score:.2f}")
    return "\n".join(lines)


def _patch_with_policy_hint(
    patch: AIQuestionPatch,
    *,
    action_prediction: RepairActionPrediction | None,
    family_prediction: RepairActionPrediction | None,
    trajectory_plan: TrajectoryPlan | None = None,
) -> AIQuestionPatch:
    hints: list[str] = []
    action_candidates, family_candidates = _patch_action_candidates(patch)
    if action_prediction and action_prediction.best_label in action_candidates:
        patch.confidence = max(patch.confidence, min(0.98, action_prediction.best_confidence))
        hints.append(f"策略模型支持 {_format_policy_prediction(action_prediction)}")
    elif family_prediction and family_prediction.best_label in family_candidates:
        patch.confidence = max(patch.confidence, min(0.96, family_prediction.best_confidence))
        hints.append(f"策略模型支持 {_format_policy_prediction(family_prediction)}")
    elif family_prediction and family_prediction.best_confidence >= 0.7:
        hints.append(f"策略模型更像 {_format_policy_prediction(family_prediction)}")

    # 轨迹策略模型加成
    if trajectory_plan is not None and not trajectory_plan.is_empty:
        primary = trajectory_plan.primary_action
        if primary and primary in action_candidates:
            patch.confidence = max(patch.confidence, min(0.97, trajectory_plan.primary_confidence))
            traj_desc = _format_trajectory_plan(trajectory_plan)
            hints.append(f"轨迹策略 {traj_desc}")
        elif primary:
            traj_desc = _format_trajectory_plan(trajectory_plan)
            hints.append(f"轨迹建议 {traj_desc}")

    if hints:
        base_summary = (patch.summary or "").strip()
        if base_summary:
            patch.summary = base_summary + "；" + "；".join(hints[:3])
        else:
            patch.summary = "；".join(hints[:3])
    return patch


def _find_option(question: QuestionNode, letter: str):
    normalized = (letter or "").strip().upper()
    for option in question.options:
        if (option.letter or "").strip().upper() == normalized:
            return option
    return None


def _ensure_option_exists(question: QuestionNode, letter: str):
    normalized = (letter or "").strip().upper()
    if len(normalized) != 1 or not ("A" <= normalized <= "Z"):
        return None
    option = _find_option(question, normalized)
    if option is not None:
        return option

    target_index = ord(normalized) - ord("A")
    while len(question.options) <= target_index:
        anchor_letter = question.options[-1].letter if question.options else None
        if not insert_option_after(question, anchor_letter):
            break
    return _find_option(question, normalized)


def _normalize_question_number(value: str) -> str:
    text = (value or "").strip()
    if not text:
        return ""
    match = re.search(r"(\d{1,4})", text)
    if match:
        return match.group(1)
    return text


def _extract_leading_question_prefix(value: str) -> tuple[str, str] | None:
    text = _normalize_text(value, keep_line_breaks=True)
    if not text:
        return None
    for pattern in _LEADING_QUESTION_PATTERNS:
        match = pattern.match(text)
        if not match:
            continue
        number = _normalize_question_number(match.group("number"))
        stem = _normalize_text(match.group("stem"), keep_line_breaks=True)
        if number and stem:
            return number, stem
    return None


def _normalize_text(value: str, *, keep_line_breaks: bool = False) -> str:
    text = (value or "").replace("\xa0", " ").replace("\u3000", " ").strip()
    if not text:
        return ""
    text = re.sub(r"(?<=\d)\s*/\s*(?=\d)", "/", text)
    text = re.sub(r"(?<=\d)\s+(?=\d/\d)", "", text)
    text = re.sub(r"(?<=\d/\d)\s+(?=\d)", "", text)
    lines = [re.sub(r"\s+", " ", line).strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return ""
    if keep_line_breaks:
        preserved: list[str] = []
        for line in lines:
            if re.match(r"^(?:[（(]?\d+[)）]|[①②③④⑤⑥⑦⑧⑨⑩])", line):
                preserved.append(line)
            else:
                if preserved and not re.match(r"^(?:[（(]?\d+[)）]|[①②③④⑤⑥⑦⑧⑨⑩])", preserved[-1]):
                    preserved[-1] = f"{preserved[-1]} {line}".strip()
                else:
                    preserved.append(line)
        text = "\n".join(preserved)
    else:
        text = " ".join(lines)
    text = re.sub(r"\s+([，。；：！？,.!?])", r"\1", text)
    text = re.sub(r"([（(【《“‘])\s+", r"\1", text)
    text = re.sub(r"\s+([）)】》”’])", r"\1", text)
    if text.startswith(_FILL_BLANK_PREFIXES):
        text = "____" + text
    return text.strip()


def _normalize_definition_judgment_stem(value: str) -> str:
    text = _normalize_text(value, keep_line_breaks=True)
    if not text or "是指" not in text:
        return text
    for marker in _DEFINITION_ASK_MARKERS:
        index = text.rfind(marker)
        if index <= 8:
            continue
        tail = text[index:].strip()
        if len(tail) < 4:
            continue
        if "\n" in text[max(0, index - 2) : index + 2]:
            return text
        head = text[:index].rstrip(" ，,;；")
        if len(head) < 8:
            continue
        return f"{head}\n{tail}".strip()
    return text


def _normalize_sentence_expression_stem(value: str) -> str:
    text = _normalize_text(value, keep_line_breaks=True)
    if not text:
        return text
    enum_count = len(_ENUMERATION_MARKER_RE.findall(text)) + len(_PAREN_ENUMERATION_MARKER_RE.findall(text))
    if enum_count < 3 and not any(marker in text for marker in _SENTENCE_ORDER_TAIL_MARKERS):
        return text
    text = re.sub(r"\s*([①②③④⑤⑥⑦⑧⑨⑩])\s*", r"\n\1", text).lstrip()
    text = re.sub(r"\s*([（(]\s*\d+\s*[）)])\s*", r"\n\1", text).lstrip()
    for marker in _SENTENCE_ORDER_TAIL_MARKERS:
        index = text.find(marker)
        if index > 0 and "\n" not in text[max(0, index - 1) : index + 1]:
            text = f"{text[:index].rstrip()}\n{text[index:].lstrip()}".strip()
            break
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    return "\n".join(lines)


def _normalize_stem_for_subtype(value: str, subtype: str) -> str:
    text = _normalize_text(value, keep_line_breaks=True)
    normalized_subtype = (subtype or "").strip()
    if normalized_subtype == "定义判断":
        return _normalize_definition_judgment_stem(text)
    if normalized_subtype == "语句表达":
        return _normalize_sentence_expression_stem(text)
    if normalized_subtype == "片段阅读":
        for marker in _READING_ASK_MARKERS:
            index = text.rfind(marker)
            if index > 18 and "\n" not in text[max(0, index - 2) : index + 2]:
                return f"{text[:index].rstrip()}\n{text[index:].lstrip()}".strip()
    return text


def _normalize_options_for_subtype(question: QuestionNode, subtype: str) -> list[AIOptionPatch]:
    normalized_subtype = (subtype or "").strip()
    patches: list[AIOptionPatch] = []
    if normalized_subtype == "逻辑填空":
        for option in question.options:
            text = _strip_option_marker(option.text or "")
            normalized = _normalize_text(text)
            if normalized and normalized != (option.text or "").strip():
                patches.append(AIOptionPatch(letter=option.letter, text=normalized))
    return patches


def _should_reassign_graphic_stem_assets(question: QuestionNode, subtype: str) -> bool:
    if (subtype or "").strip() != "图形推理":
        return False
    if len(question.stem_assets) != 4 or len(question.options) < 4:
        return False
    if any(option.image_path for option in question.options[:4]):
        return False
    short_or_blank_options = 0
    for option in question.options[:4]:
        text = _strip_option_marker(option.text or "")
        if len(text) <= 2:
            short_or_blank_options += 1
    if short_or_blank_options < 4:
        return False
    stem_text = _normalize_text(question.stem or "", keep_line_breaks=True)
    if len(stem_text) > 24:
        return False
    return True


def _should_move_data_stem_assets_to_material(material: MaterialSet | None, question: QuestionNode, subtype: str) -> bool:
    if material is None or not question.stem_assets:
        return False
    if material.questions and material.questions[0] is not question:
        return False

    normalized_subtype = (subtype or "").strip()
    if normalized_subtype and normalized_subtype not in {"表格型资料分析", "图形型资料分析", "综合型资料分析", "文字型资料分析"}:
        return False

    stem_text = _normalize_text(question.stem or "", keep_line_breaks=True)
    if not stem_text or len(stem_text) > 38:
        return False
    if not any(marker in stem_text for marker in _DATA_ASK_MARKERS):
        return False
    if any(option.image_path for option in question.options):
        return False

    if (material.body or "").strip():
        signal_count = sum(
            1 for token in ("同比", "环比", "增长率", "表", "图", "资料", "材料", "比重", "投资", "人数", "企业", "亿元")
            if token in (material.body or "")
        )
        if signal_count <= 0 and len(question.stem_assets) <= 1:
            return False
    return True


def _split_data_material_intro_from_stem(material: MaterialSet | None, question: QuestionNode, subtype: str) -> tuple[str, str] | None:
    if material is None:
        return None
    stem_text = _normalize_text(question.stem or "", keep_line_breaks=True)
    if len(stem_text) < 36:
        return None
    if (material.body or "").strip() and len((material.body or "").strip()) >= 12:
        return None
    normalized_subtype = (subtype or "").strip()
    if normalized_subtype and normalized_subtype not in {"表格型资料分析", "图形型资料分析", "综合型资料分析", "文字型资料分析"}:
        return None
    ask_index = min((index for marker in _DATA_ASK_MARKERS if (index := stem_text.find(marker)) > 12), default=-1)
    if ask_index <= 12:
        return None
    intro = stem_text[:ask_index].strip(" ：:；;，,")
    remain = stem_text[ask_index:].strip()
    if len(intro) < 14 or len(remain) < 8:
        return None
    if not any(token in intro for token in ("同比", "环比", "增长率", "表", "图", "资料", "材料", "比重", "投资", "人数", "企业", "亿元")):
        return None
    return intro, remain


def _merge_page_numbers(*groups: list[int] | tuple[int, ...] | None) -> list[int]:
    values: list[int] = []
    seen: set[int] = set()
    for group in groups:
        for value in group or []:
            try:
                numeric = int(value)
            except Exception:
                continue
            if numeric not in seen:
                seen.add(numeric)
                values.append(numeric)
    return sorted(values)


def _strip_option_marker(value: str) -> str:
    text = _normalize_text(value)
    return re.sub(r"^[A-D](?:[\.．、:：]|(?=\s))\s*", "", text).strip()


def _extract_inline_options(value: str) -> dict[str, str]:
    text = _normalize_text(value)
    matches = list(_INLINE_OPTION_PATTERN.finditer(text))
    if len(matches) < 2:
        return {}
    letters = [match.group(1) for match in matches]
    if letters != sorted(letters) or len(set(letters)) != len(letters):
        return {}

    result: dict[str, str] = {}
    for index, match in enumerate(matches):
        start = match.end()
        end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        option_text = text[start:end].strip()
        if option_text:
            result[match.group(1)] = option_text
    return result if len(result) >= 2 else {}


def _extract_question_parts_from_text(value: str) -> tuple[str, dict[str, str]] | None:
    text = _normalize_text(value, keep_line_breaks=True)
    if not text:
        return None
    matches = list(_INLINE_OPTION_PATTERN.finditer(text))
    if len(matches) < 4:
        return None
    letters = [match.group(1) for match in matches[:4]]
    if letters[:4] != ["A", "B", "C", "D"]:
        return None
    stem = _normalize_text(text[: matches[0].start()], keep_line_breaks=True)
    if len(stem) < 4:
        return None
    parsed = _extract_inline_options(text)
    if len(parsed) < 4:
        return None
    return stem, {letter: parsed[letter] for letter in ("A", "B", "C", "D") if letter in parsed}


def _combined_option_parse(question: QuestionNode) -> dict[str, str]:
    joined = " ".join((option.text or "").strip() for option in question.options if (option.text or "").strip())
    parsed = _extract_inline_options(joined)
    if parsed:
        return parsed
    stem_parsed = _extract_inline_options(question.stem or "")
    return stem_parsed


def _suggest_option_layout(
    question: QuestionNode,
    option_texts: list[str] | None = None,
    *,
    subtype: str = "",
) -> str | None:
    normalized_subtype = (subtype or "").strip()
    if normalized_subtype == "图形推理":
        return "grid"
    if any(option.image_path for option in question.options):
        return "grid"
    if option_texts is None:
        texts = [_strip_option_marker(option.text or "") for option in question.options if _strip_option_marker(option.text or "")]
    else:
        texts = [_normalize_text(text) for text in option_texts if _normalize_text(text)]
    if not texts:
        return None
    if normalized_subtype == "逻辑填空":
        max_len = max(len(text) for text in texts)
        if max_len <= 8:
            return "one_row"
        if max_len <= 16:
            return "grid"
        return "list"
    if normalized_subtype == "类比推理" and max(len(text) for text in texts) <= 18:
        return "grid"
    max_len = max(len(text) for text in texts)
    avg_len = sum(len(text) for text in texts) / max(len(texts), 1)
    numeric_like = all(bool(re.fullmatch(r"[\d\.\-/:%×xX+\s]+", text)) for text in texts)
    if numeric_like and max_len <= 12:
        return "one_row"
    if max_len <= 6 and avg_len <= 4:
        return "one_row"
    if max_len <= 18 and avg_len <= 11:
        return "grid"
    return "list"


def _subject_patch(section: Section, material: MaterialSet | None, question: QuestionNode):
    diagnostics = infer_subject_diagnostics(
        stem=question.stem or "",
        options=[option.text or "" for option in question.options],
        material_text=(material.body or "") if material is not None else "",
        image_count=len(question.stem_assets) + sum(1 for option in question.options if option.image_path),
        material_header=(material.header or "") if material is not None else "",
        allow_data=True,
    )
    inferred_kind = diagnostics.kind
    inferred_confidence = diagnostics.confidence
    if inferred_kind == "unknown":
        return None, 0.0
    if section.kind == "unknown" and inferred_confidence >= 0.72:
        return inferred_kind, inferred_confidence
    if section.kind != "data" and inferred_kind != section.kind and inferred_confidence >= 0.78:
        if len(section.questions) <= 1:
            return inferred_kind, inferred_confidence
    return None, inferred_confidence


def _expected_boundary_letter(question: QuestionNode) -> str | None:
    blank_letters = [
        option.letter
        for option in question.options
        if not (option.text or "").strip() and not option.image_path
    ]
    if blank_letters:
        return blank_letters[0]
    if len(question.options) >= 4:
        return None
    next_index = len(question.options)
    if 0 <= next_index < 4:
        return chr(ord("A") + next_index)
    return None


def _expected_boundary_letters(question: QuestionNode) -> list[str]:
    letters: list[str] = []
    for code in range(ord("A"), ord("D") + 1):
        letter = chr(code)
        option = _find_option(question, letter)
        is_missing = option is None or (not (option.text or "").strip() and not option.image_path)
        if is_missing:
            letters.append(letter)
        elif letters:
            break
    return letters


def _split_leading_spilled_option(
    stem: str,
    *,
    expected_letter: str,
    current_number: str,
) -> tuple[str, str] | None:
    text = _normalize_text(stem)
    if not text:
        return None
    match = _LEADING_OPTION_PATTERN.match(text)
    if not match or match.group(1) != expected_letter:
        return None
    rest = match.group(2).strip()
    normalized_number = _normalize_question_number(current_number)
    if not normalized_number:
        return None
    number_pattern = re.compile(
        rf"(?<!\d){re.escape(normalized_number)}(?:[\.．、]?\s*)(?=[^\d]|$)"
    )
    for number_match in number_pattern.finditer(rest):
        if number_match.start() <= 0 or number_match.start() > 60:
            continue
        option_text = rest[: number_match.start()].strip(" ;；，,")
        new_stem = rest[number_match.end() :].strip()
        if not option_text or len(new_stem) < 2:
            continue
        return option_text, _normalize_text(new_stem, keep_line_breaks=True)
    return None


def _split_leading_spilled_options(
    stem: str,
    *,
    expected_letters: list[str],
    current_number: str,
) -> tuple[list[tuple[str, str]], str] | None:
    text = _normalize_text(stem)
    if not text or not expected_letters:
        return None
    first_match = _LEADING_OPTION_PATTERN.match(text)
    if not first_match or first_match.group(1) != expected_letters[0]:
        return None
    normalized_number = _normalize_question_number(current_number)
    if not normalized_number:
        return None
    number_pattern = re.compile(
        rf"(?<!\d){re.escape(normalized_number)}(?:[\.．、]?\s*)(?=[^\d]|$)"
    )
    for number_match in number_pattern.finditer(text):
        if number_match.start() <= 0 or number_match.start() > 120:
            continue
        prefix = text[: number_match.start()].strip()
        new_stem = text[number_match.end() :].strip()
        if len(new_stem) < 2:
            continue
        parsed = _extract_inline_options(prefix)
        if not parsed:
            single = _LEADING_OPTION_PATTERN.match(prefix)
            if single:
                parsed = {single.group(1): single.group(2).strip()}
        if not parsed:
            continue
        parsed_letters = list(parsed.keys())
        if parsed_letters != expected_letters[: len(parsed_letters)]:
            continue
        option_pairs = [(letter, _normalize_text(parsed[letter])) for letter in parsed_letters if parsed[letter].strip()]
        if option_pairs:
            return option_pairs, _normalize_text(new_stem, keep_line_breaks=True)
    return None


def _split_embedded_next_question_transition(
    stem: str,
    *,
    next_number: str,
    min_head_length: int = 6,
) -> tuple[str, str] | None:
    text = _normalize_text(stem, keep_line_breaks=True)
    normalized_next_number = _normalize_question_number(next_number)
    if not text or not normalized_next_number:
        return None

    patterns = (
        re.compile(
            rf"^(?P<head>.*?\S)\s+(?P<number>{re.escape(normalized_next_number)})\s*[\.．、)\uFF09:：]\s*(?P<tail>.+)$"
        ),
        re.compile(
            rf"^(?P<head>.*?\S)\s+第\s*(?P<number>{re.escape(normalized_next_number)})\s*题(?:\s*[\.．、)\uFF09:：]\s*)?(?P<tail>.+)$"
        ),
    )
    for pattern in patterns:
        match = pattern.match(text)
        if not match:
            continue
        head = _normalize_text(match.group("head"), keep_line_breaks=True)
        tail = _normalize_text(match.group("tail"), keep_line_breaks=True)
        if len(head) >= min_head_length and len(tail) >= 4:
            return head, tail
    return None


def _looks_like_prompt_text(value: str) -> bool:
    text = _normalize_text(value, keep_line_breaks=True)
    if not text or len(text) > 48:
        return False
    if text.startswith("请开始答题"):
        return True
    if not text.startswith(_PROMPT_START_MARKERS):
        return False
    if not any(marker in text for marker in _PROMPT_CONTEXT_MARKERS):
        return False
    return any(marker in text for marker in _PROMPT_ACTION_MARKERS)


def _extract_prompt_suffix(value: str) -> tuple[str, str] | None:
    text = _normalize_text(value, keep_line_breaks=True)
    if not text:
        return None
    candidate_indexes = sorted(
        {
            text.rfind(marker)
            for marker in _PROMPT_START_MARKERS
            if text.rfind(marker) > 0
        }
    )
    for start in candidate_indexes:
        head = _normalize_text(text[:start], keep_line_breaks=True)
        tail = _normalize_text(text[start:], keep_line_breaks=True)
        if len(head) < 4 or not _looks_like_prompt_text(tail):
            continue
        return head, tail
    return None


def _extract_missing_question_before_current(
    previous_question: QuestionNode | None,
    current_question: QuestionNode,
) -> tuple[QuestionNode, str] | None:
    previous_numeric = getattr(previous_question, "numeric_source_number", None)
    current_numeric = getattr(current_question, "numeric_source_number", None)
    if previous_numeric is None or current_numeric is None:
        return None
    if current_numeric <= previous_numeric + 1:
        return None
    missing_number = str(previous_numeric + 1)

    leading = _extract_leading_question_prefix(current_question.stem or "")
    if not leading:
        return None
    leading_number, stripped = leading
    if leading_number != missing_number:
        return None

    patterns = (
        re.compile(
            rf"^(?P<body>.*?\S)\s+(?P<number>{re.escape(str(current_numeric))})\s*[\.．、)\uFF09:：]\s*(?P<tail>.+)$"
        ),
        re.compile(
            rf"^(?P<body>.*?\S)\s+第\s*(?P<number>{re.escape(str(current_numeric))})\s*题(?:\s*[\.．、)\uFF09:：]\s*)?(?P<tail>.+)$"
        ),
    )
    for pattern in patterns:
        match = pattern.match(stripped)
        if not match:
            continue
        body = _normalize_text(match.group("body"), keep_line_breaks=True)
        current_stem = _normalize_text(match.group("tail"), keep_line_breaks=True)
        if not current_stem:
            continue
        question_parts = _extract_question_parts_from_text(body)
        if not question_parts:
            continue
        new_stem, option_map = question_parts
        options = [
            OptionNode(letter=letter, text=_normalize_text(option_map.get(letter, "")))
            for letter in ("A", "B", "C", "D")
        ]
        new_question = QuestionNode(
            source_number=missing_number,
            stem=new_stem,
            options=options,
            page_numbers=list(current_question.page_numbers),
        )
        return new_question, current_stem
    return None


def _extract_embedded_question_from_source(
    source_text: str,
    *,
    missing_number: str,
    current_number: str | None,
    page_numbers: list[int],
    source_page_numbers: list[int] | None = None,
) -> tuple[str, QuestionNode, str | None] | None:
    text = _normalize_text(source_text, keep_line_breaks=True)
    if not text:
        return None

    prefix = ""
    body = ""
    direct = _extract_leading_question_prefix(text)
    if direct and direct[0] == missing_number:
        body = direct[1]
    else:
        patterns = (
            re.compile(
                rf"^(?P<prefix>.*?\S)\s*(?P<number>{re.escape(missing_number)})\s*[\.．、)\uFF09:：]\s*(?P<body>.+)$"
            ),
            re.compile(
                rf"^(?P<prefix>.*?\S)\s*第\s*(?P<number>{re.escape(missing_number)})\s*题(?:\s*[\.．、)\uFF09:：]\s*)?(?P<body>.+)$"
            ),
        )
        for pattern in patterns:
            match = pattern.match(text)
            if not match:
                continue
            prefix = _normalize_text(match.group("prefix"), keep_line_breaks=True)
            body = _normalize_text(match.group("body"), keep_line_breaks=True)
            break
    if not body:
        return None

    current_stem: str | None = None
    question_body = body
    normalized_current = _normalize_question_number(current_number or "")
    if normalized_current:
        patterns = (
            re.compile(
                rf"^(?P<body>.*?\S)\s+(?P<number>{re.escape(normalized_current)})\s*[\.．、)\uFF09:：]\s*(?P<tail>.+)$"
            ),
            re.compile(
                rf"^(?P<body>.*?\S)\s+第\s*(?P<number>{re.escape(normalized_current)})\s*题(?:\s*[\.．、)\uFF09:：]\s*)?(?P<tail>.+)$"
            ),
        )
        for pattern in patterns:
            match = pattern.match(body)
            if not match:
                continue
            question_body = _normalize_text(match.group("body"), keep_line_breaks=True)
            current_stem = _normalize_text(match.group("tail"), keep_line_breaks=True)
            break

    question_parts = _extract_question_parts_from_text(question_body)
    if not question_parts:
        return None
    stem, option_map = question_parts
    new_question = QuestionNode(
        source_number=missing_number,
        stem=stem,
        options=[
            OptionNode(letter=letter, text=_normalize_text(option_map.get(letter, "")))
            for letter in ("A", "B", "C", "D")
        ],
        page_numbers=_merge_page_numbers(source_page_numbers, page_numbers),
    )
    return prefix, new_question, current_stem


def repair_spilled_boundary_between_questions(
    previous_question: QuestionNode | None,
    current_question: QuestionNode,
) -> tuple[int, str]:
    if previous_question is None:
        return 0, ""
    expected_letters = _expected_boundary_letters(previous_question)
    if not expected_letters:
        return 0, ""
    split = _split_leading_spilled_options(
        current_question.stem or "",
        expected_letters=expected_letters,
        current_number=current_question.source_number or "",
    )
    if not split:
        return 0, ""

    option_pairs, new_stem = split
    previous_count = len(previous_question.options)
    changes = 0
    repaired_letters: list[str] = []
    for letter, option_text in option_pairs:
        option = _ensure_option_exists(previous_question, letter)
        if option is None:
            continue
        repaired_letters.append(letter)
        if (option.text or "").strip() != option_text:
            update_option_text(previous_question, letter, option_text)
            changes += 1
    option_growth = len(previous_question.options) - previous_count
    if option_growth > 0:
        changes += option_growth
    if (current_question.stem or "").strip() != new_stem:
        update_question_stem(current_question, new_stem)
        changes += 1
    if changes <= 0:
        return 0, ""
    if len(repaired_letters) == 1:
        label = repaired_letters[0]
    else:
        label = "/".join(repaired_letters)
    return changes, f"把上一题遗落的 {label} 选项从当前题题干中拆回。"


def repair_number_gap_from_stem(
    previous_question: QuestionNode | None,
    current_question: QuestionNode,
) -> tuple[int, str]:
    current_source = _normalize_question_number(current_question.source_number)
    leading = _extract_leading_question_prefix(current_question.stem or "")
    if not leading:
        return 0, ""
    leading_number, stripped_stem = leading
    previous_numeric = getattr(previous_question, "numeric_source_number", None)
    current_numeric = getattr(current_question, "numeric_source_number", None)

    changes = 0
    reasons: list[str] = []
    if current_source and leading_number == current_source:
        if stripped_stem != (current_question.stem or "").strip():
            update_question_stem(current_question, stripped_stem)
            changes += 1
            reasons.append("去掉了题干里重复的题号前缀")
        return changes, "；".join(reasons)

    if previous_numeric is None or leading_number == current_source:
        return 0, ""

    if int(leading_number) == previous_numeric + 1:
        renumber_question(current_question, leading_number)
        changes += 1
        reasons.append("按题干前缀回正了跳号题号")
        if stripped_stem != (current_question.stem or "").strip():
            update_question_stem(current_question, stripped_stem)
            changes += 1
            reasons.append("去掉了题干里的题号前缀")
        return changes, "；".join(reasons)

    if current_numeric is not None and int(leading_number) == current_numeric - 1:
        renumber_question(current_question, leading_number)
        changes += 1
        reasons.append("按题干前缀回正了题号")
        if stripped_stem != (current_question.stem or "").strip():
            update_question_stem(current_question, stripped_stem)
            changes += 1
            reasons.append("去掉了题干里的题号前缀")
        return changes, "；".join(reasons)
    return 0, ""


def repair_embedded_next_question_between_questions(
    current_question: QuestionNode,
    next_question: QuestionNode | None,
) -> tuple[int, str]:
    if next_question is None:
        return 0, ""
    split = _split_embedded_next_question_transition(
        current_question.stem or "",
        next_number=next_question.source_number or "",
        min_head_length=6,
    )
    if not split:
        return 0, ""

    current_stem, next_stem_candidate = split
    next_stem = _normalize_text(next_question.stem or "", keep_line_breaks=True)
    changes = 0
    reasons: list[str] = []

    if current_stem != (current_question.stem or "").strip():
        update_question_stem(current_question, current_stem)
        changes += 1
        reasons.append("从当前题干里切走了下一题题号")

    should_update_next = False
    if not next_stem:
        should_update_next = True
    elif next_stem_candidate.startswith(next_stem) or next_stem.startswith(next_stem_candidate):
        should_update_next = len(next_stem_candidate) > len(next_stem)
    elif next_stem[:8] and next_stem_candidate[:8] and next_stem[:8] == next_stem_candidate[:8]:
        should_update_next = len(next_stem_candidate) > len(next_stem)

    if should_update_next and next_stem_candidate != (next_question.stem or "").strip():
        update_question_stem(next_question, next_stem_candidate)
        changes += 1
        reasons.append("把切出的题干补给了下一题")

    if changes <= 0:
        return 0, ""
    return changes, "；".join(reasons)


def repair_embedded_next_question_from_options(
    current_question: QuestionNode,
    next_question: QuestionNode | None,
) -> tuple[int, str]:
    if next_question is None:
        return 0, ""
    next_number = next_question.source_number or ""
    next_stem = _normalize_text(next_question.stem or "", keep_line_breaks=True)
    for option in reversed(list(current_question.options)):
        option_text = _normalize_text(option.text or "", keep_line_breaks=True)
        if not option_text:
            continue
        split = _split_embedded_next_question_transition(
            option_text,
            next_number=next_number,
            min_head_length=1,
        )
        if not split:
            continue
        current_option_text, next_stem_candidate = split
        changes = 0
        reasons: list[str] = []
        if current_option_text != (option.text or "").strip():
            update_option_text(current_question, option.letter, current_option_text)
            changes += 1
            reasons.append(f"从 {option.letter} 选项里切走了下一题题号")

        should_update_next = False
        if not next_stem:
            should_update_next = True
        elif next_stem_candidate.startswith(next_stem) or next_stem.startswith(next_stem_candidate):
            should_update_next = len(next_stem_candidate) > len(next_stem)
        elif next_stem[:8] and next_stem_candidate[:8] and next_stem[:8] == next_stem_candidate[:8]:
            should_update_next = len(next_stem_candidate) > len(next_stem)

        if should_update_next and next_stem_candidate != (next_question.stem or "").strip():
            update_question_stem(next_question, next_stem_candidate)
            changes += 1
            reasons.append("把切出的题干补给了下一题")
        if changes > 0:
            return changes, "；".join(reasons)
    return 0, ""


def repair_missing_question_before_current(
    project: ExamProject,
    previous_question: QuestionNode | None,
    current_question: QuestionNode,
) -> tuple[int, str]:
    extracted = _extract_missing_question_before_current(previous_question, current_question)
    if not extracted:
        return 0, ""
    new_question, current_stem = extracted
    inserted = insert_question_before(project, current_question, new_question)
    if not inserted:
        return 0, ""
    changes = 1
    if current_stem != (current_question.stem or "").strip():
        update_question_stem(current_question, current_stem)
        changes += 1
    return changes, f"从当前题题干前部恢复出漏掉的第 {new_question.source_number} 题。"


def repair_missing_question_from_previous_option(
    project: ExamProject,
    previous_question: QuestionNode | None,
    current_question: QuestionNode,
) -> tuple[int, str]:
    previous_numeric = getattr(previous_question, "numeric_source_number", None)
    current_numeric = getattr(current_question, "numeric_source_number", None)
    if previous_question is None or previous_numeric is None or current_numeric is None:
        return 0, ""
    if current_numeric <= previous_numeric + 1:
        return 0, ""

    missing_number = str(previous_numeric + 1)
    current_number = current_question.source_number or ""
    for option in reversed(list(previous_question.options)):
        extracted = _extract_embedded_question_from_source(
            option.text or "",
            missing_number=missing_number,
            current_number=current_number,
            page_numbers=current_question.page_numbers,
            source_page_numbers=previous_question.page_numbers,
        )
        if not extracted:
            continue
        prefix_text, new_question, current_stem = extracted
        if not prefix_text:
            continue
        if not insert_question_before(project, current_question, new_question):
            continue

        changes = 1
        if prefix_text != (option.text or "").strip():
            update_option_text(previous_question, option.letter, prefix_text)
            changes += 1
        if current_stem and current_stem != (current_question.stem or "").strip():
            update_question_stem(current_question, current_stem)
            changes += 1
        return changes, f"从上一题 {option.letter} 选项尾部恢复出漏掉的第 {missing_number} 题。"
    return 0, ""


def repair_missing_question_from_previous_stem(
    project: ExamProject,
    previous_question: QuestionNode | None,
    current_question: QuestionNode,
) -> tuple[int, str]:
    previous_numeric = getattr(previous_question, "numeric_source_number", None)
    current_numeric = getattr(current_question, "numeric_source_number", None)
    if previous_question is None or previous_numeric is None or current_numeric is None:
        return 0, ""
    if current_numeric <= previous_numeric + 1:
        return 0, ""

    missing_number = str(previous_numeric + 1)
    extracted = _extract_embedded_question_from_source(
        previous_question.stem or "",
        missing_number=missing_number,
        current_number=current_question.source_number or "",
        page_numbers=current_question.page_numbers,
        source_page_numbers=previous_question.page_numbers,
    )
    if not extracted:
        return 0, ""
    prefix_text, new_question, current_stem = extracted
    if not prefix_text:
        return 0, ""
    if not insert_question_before(project, current_question, new_question):
        return 0, ""

    changes = 1
    if prefix_text != (previous_question.stem or "").strip():
        update_question_stem(previous_question, prefix_text)
        changes += 1
    if current_stem and current_stem != (current_question.stem or "").strip():
        update_question_stem(current_question, current_stem)
        changes += 1
    return changes, f"从上一题题干尾部恢复出漏掉的第 {missing_number} 题。"


def repair_missing_question_after_current_from_stem(
    project: ExamProject,
    current_question: QuestionNode,
    next_question: QuestionNode | None,
) -> tuple[int, str]:
    current_numeric = getattr(current_question, "numeric_source_number", None)
    next_numeric = getattr(next_question, "numeric_source_number", None)
    if next_question is None or current_numeric is None or next_numeric is None:
        return 0, ""
    if next_numeric <= current_numeric + 1:
        return 0, ""

    missing_number = str(current_numeric + 1)
    extracted = _extract_embedded_question_from_source(
        current_question.stem or "",
        missing_number=missing_number,
        current_number=next_question.source_number or "",
        page_numbers=next_question.page_numbers,
        source_page_numbers=current_question.page_numbers,
    )
    if not extracted:
        return 0, ""
    prefix_text, new_question, next_stem = extracted
    if not prefix_text:
        return 0, ""
    if not insert_question_before(project, next_question, new_question):
        return 0, ""

    changes = 1
    if prefix_text != (current_question.stem or "").strip():
        update_question_stem(current_question, prefix_text)
        changes += 1
    if next_stem and next_stem != (next_question.stem or "").strip():
        update_question_stem(next_question, next_stem)
        changes += 1
    return changes, f"从当前题题干尾部恢复出漏掉的第 {missing_number} 题。"


def repair_missing_question_after_current_from_option(
    project: ExamProject,
    current_question: QuestionNode,
    next_question: QuestionNode | None,
) -> tuple[int, str]:
    current_numeric = getattr(current_question, "numeric_source_number", None)
    next_numeric = getattr(next_question, "numeric_source_number", None)
    if next_question is None or current_numeric is None or next_numeric is None:
        return 0, ""
    if next_numeric <= current_numeric + 1:
        return 0, ""

    missing_number = str(current_numeric + 1)
    for option in reversed(list(current_question.options)):
        extracted = _extract_embedded_question_from_source(
            option.text or "",
            missing_number=missing_number,
            current_number=next_question.source_number or "",
            page_numbers=next_question.page_numbers,
            source_page_numbers=current_question.page_numbers,
        )
        if not extracted:
            continue
        prefix_text, new_question, next_stem = extracted
        if not prefix_text:
            continue
        if not insert_question_before(project, next_question, new_question):
            continue

        changes = 1
        if prefix_text != (option.text or "").strip():
            update_option_text(current_question, option.letter, prefix_text)
            changes += 1
        if next_stem and next_stem != (next_question.stem or "").strip():
            update_question_stem(next_question, next_stem)
            changes += 1
        return changes, f"从当前题 {option.letter} 选项尾部恢复出漏掉的第 {missing_number} 题。"
    return 0, ""


def repair_prompt_contamination_between_questions(
    current_question: QuestionNode,
    next_question: QuestionNode | None,
) -> tuple[int, str]:
    if next_question is None:
        return 0, ""

    stem_split = _extract_prompt_suffix(current_question.stem or "")
    if stem_split:
        cleaned_stem, prompt_text = stem_split
        changes = 0
        if cleaned_stem != (current_question.stem or "").strip():
            update_question_stem(current_question, cleaned_stem)
            changes += 1
        next_stem = _normalize_text(next_question.stem or "", keep_line_breaks=True)
        if not next_stem and prompt_text:
            update_question_stem(next_question, prompt_text)
            changes += 1
            return changes, "把误挂在当前题题干尾部的提示语移到了下一题。"
        if changes > 0:
            return changes, "清理了误挂在当前题题干尾部的提示语。"

    for option in reversed(list(current_question.options)):
        option_split = _extract_prompt_suffix(option.text or "")
        if not option_split:
            continue
        cleaned_text, prompt_text = option_split
        changes = 0
        if cleaned_text != (option.text or "").strip():
            update_option_text(current_question, option.letter, cleaned_text)
            changes += 1
        next_stem = _normalize_text(next_question.stem or "", keep_line_breaks=True)
        if not next_stem and prompt_text:
            update_question_stem(next_question, prompt_text)
            changes += 1
            return changes, f"把误挂在当前题 {option.letter} 选项尾部的提示语移到了下一题。"
        if changes > 0:
            return changes, f"清理了误挂在当前题 {option.letter} 选项尾部的提示语。"
    return 0, ""


def repair_question_boundary(
    project: ExamProject,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
) -> tuple[int, str]:
    total_changes = 0
    reasons: list[str] = []

    def _apply_step(step):
        nonlocal total_changes
        previous_question, next_question = _neighbor_questions(section, material, question)
        changes, reason = step(previous_question, next_question)
        if changes > 0:
            total_changes += changes
            if reason:
                reasons.append(reason)

    for step in (
        lambda previous_question, _next_question: repair_missing_question_before_current(project, previous_question, question),
        lambda previous_question, _next_question: repair_missing_question_from_previous_stem(project, previous_question, question),
        lambda previous_question, _next_question: repair_missing_question_from_previous_option(project, previous_question, question),
        lambda previous_question, _next_question: repair_number_gap_from_stem(previous_question, question),
        lambda previous_question, _next_question: repair_spilled_boundary_between_questions(previous_question, question),
        lambda _previous_question, next_question: repair_missing_question_after_current_from_stem(project, question, next_question),
        lambda _previous_question, next_question: repair_missing_question_after_current_from_option(project, question, next_question),
        lambda _previous_question, next_question: repair_prompt_contamination_between_questions(question, next_question),
        lambda _previous_question, next_question: repair_embedded_next_question_between_questions(question, next_question),
        lambda _previous_question, next_question: repair_embedded_next_question_from_options(question, next_question),
    ):
        _apply_step(step)
    return total_changes, "；".join(reasons)


def _iter_question_groups(project: ExamProject):
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                yield section, material, material.questions
        else:
            yield section, None, section.questions


def repair_project_boundaries(project: ExamProject) -> tuple[set[int], int]:
    changed_questions: set[int] = set()
    field_changes = 0
    for _section, _material, question_list in _iter_question_groups(project):
        index = 1
        while index < len(question_list):
            previous_question = question_list[index - 1]
            current_question = question_list[index]
            changes = 0
            for structural_changes, _reason in (
                repair_missing_question_before_current(project, previous_question, current_question),
                repair_missing_question_from_previous_stem(project, previous_question, current_question),
                repair_missing_question_from_previous_option(project, previous_question, current_question),
            ):
                if structural_changes > 0:
                    changes = structural_changes
                    break
            if changes > 0:
                field_changes += changes
                changed_questions.add(id(previous_question))
                changed_questions.add(id(current_question))
                if index < len(question_list):
                    changed_questions.add(id(question_list[index]))
                index = min(index + 1, len(question_list) - 1)
                continue
            pair_changes = 0
            for one_changes, _reason in (
                repair_number_gap_from_stem(previous_question, current_question),
                repair_spilled_boundary_between_questions(previous_question, current_question),
            ):
                pair_changes += one_changes
            if pair_changes > 0:
                field_changes += pair_changes
                changed_questions.add(id(previous_question))
                changed_questions.add(id(current_question))
            index += 1
        index = 0
        while index < len(question_list) - 1:
            current_question = question_list[index]
            next_question = question_list[index + 1]
            changes = 0
            for one_changes, _reason in (
                repair_missing_question_after_current_from_stem(project, current_question, next_question),
                repair_missing_question_after_current_from_option(project, current_question, next_question),
                repair_prompt_contamination_between_questions(current_question, next_question),
                repair_embedded_next_question_between_questions(current_question, next_question),
                repair_embedded_next_question_from_options(current_question, next_question),
            ):
                changes += one_changes
            if changes > 0:
                field_changes += changes
                changed_questions.add(id(current_question))
                changed_questions.add(id(next_question))
                if index + 1 < len(question_list):
                    changed_questions.add(id(question_list[index + 1]))
            index += 1
    if field_changes > 0:
        annotate_project_quality(project)
    return changed_questions, field_changes


def _suggest_local_patch(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
) -> AIQuestionPatch:
    del next_question
    patch = AIQuestionPatch(confidence=0.0)
    reasons: list[str] = []
    diagnostics = infer_subject_diagnostics(
        stem=question.stem or "",
        options=[option.text or "" for option in question.options],
        material_text=(material.body or "") if material is not None else "",
        image_count=len(question.stem_assets) + sum(1 for option in question.options if option.image_path),
        material_header=(material.header or "") if material is not None else "",
        allow_data=True,
    )
    effective_subtype = (question.inferred_subtype or diagnostics.subtype or "").strip()

    normalized_number = _normalize_question_number(question.source_number)
    if normalized_number and normalized_number != (question.source_number or "").strip():
        patch.source_number = normalized_number
        reasons.append("规范了题号")
        patch.confidence = max(patch.confidence, 0.92)

    leading_question = _extract_leading_question_prefix(question.stem or "")
    if leading_question:
        leading_number, stripped_stem = leading_question
        current_number = patch.source_number or normalized_number
        previous_numeric = getattr(previous_question, "numeric_source_number", None)
        current_numeric = getattr(question, "numeric_source_number", None)
        if current_number and leading_number == current_number:
            patch.stem = stripped_stem
            reasons.append("去掉了题干里重复的题号前缀")
            patch.confidence = max(patch.confidence, 0.86)
        elif previous_numeric is not None and int(leading_number) == previous_numeric + 1:
            patch.source_number = leading_number
            patch.stem = stripped_stem
            reasons.append("按题干前缀回正了跳号题号")
            patch.confidence = max(patch.confidence, 0.94)
        elif current_numeric is not None and int(leading_number) == current_numeric - 1:
            patch.source_number = leading_number
            patch.stem = stripped_stem
            reasons.append("按题干前缀回正了题号")
            patch.confidence = max(patch.confidence, 0.88)

    normalized_stem = _normalize_stem_for_subtype(question.stem or "", effective_subtype)
    if normalized_stem and normalized_stem != (question.stem or "").strip() and patch.stem is None:
        patch.stem = normalized_stem
        if effective_subtype:
            reasons.append(f"按{effective_subtype}结构整理了题干")
        else:
            reasons.append("清理了题干断行和空格")
        patch.confidence = max(patch.confidence, 0.74)

    combined_options = _combined_option_parse(question)
    if len(combined_options) >= 4:
        patch.options = [AIOptionPatch(letter=letter, text=text) for letter, text in sorted(combined_options.items())]
        reasons.append("从混排文本里重新拆出了选项")
        patch.confidence = max(patch.confidence, 0.9)
    else:
        option_patches: list[AIOptionPatch] = []
        for option in question.options:
            cleaned = _strip_option_marker(option.text or "")
            if cleaned and cleaned != (option.text or "").strip():
                option_patches.append(AIOptionPatch(letter=option.letter, text=cleaned))
        option_patches.extend(_normalize_options_for_subtype(question, effective_subtype))
        deduped: dict[str, AIOptionPatch] = {}
        for one_patch in option_patches:
            if one_patch.letter:
                deduped[one_patch.letter] = one_patch
        if option_patches:
            patch.options = [deduped[key] for key in sorted(deduped.keys())]
            if effective_subtype == "逻辑填空":
                reasons.append("按逻辑填空格式清理了选项")
            else:
                reasons.append("清理了选项前缀和空格")
            patch.confidence = max(patch.confidence, 0.7)

    layout_texts = list(combined_options.values()) if combined_options else None
    suggested_layout = _suggest_option_layout(question, layout_texts, subtype=effective_subtype)
    if suggested_layout and suggested_layout in _SUPPORTED_LAYOUTS and suggested_layout != ((question.option_layout or "").strip().lower()):
        patch.option_layout = suggested_layout  # type: ignore[assignment]
        if effective_subtype:
            reasons.append(f"按{effective_subtype}重新估算了选项布局")
        else:
            reasons.append("重新估算了更合适的选项布局")
        patch.confidence = max(patch.confidence, 0.7)

    if _should_reassign_graphic_stem_assets(question, effective_subtype):
        patch.reassign_stem_assets_to_options = True
        reasons.append("把图形题题干图片按顺序重挂到了选项")
        patch.confidence = max(patch.confidence, 0.88)

    if _should_move_data_stem_assets_to_material(material, question, effective_subtype):
        patch.move_stem_assets_to_material = True
        reasons.append("把首题里误挂的图表移回了材料区")
        patch.confidence = max(patch.confidence, 0.87)

    material_intro_split = _split_data_material_intro_from_stem(material, question, effective_subtype)
    if material_intro_split:
        intro_text, remain_stem = material_intro_split
        patch.prepend_material_body = intro_text
        patch.stem = remain_stem
        reasons.append("把首题里误挂的材料说明移回了材料区")
        patch.confidence = max(patch.confidence, 0.86)

    suggested_subject, inferred_confidence = _subject_patch(section, material, question)
    if suggested_subject:
        patch.subject_kind = suggested_subject  # type: ignore[assignment]
        reasons.append(f"题目更像 {SUBJECT_DISPLAY_NAMES.get(suggested_subject, suggested_subject)}")
        patch.confidence = max(patch.confidence, inferred_confidence)

    issue_codes = {issue.code for issue in getattr(question, "review_issues", []) or []}
    if "blank_option" in issue_codes and not patch.options:
        patch.confidence = max(patch.confidence, 0.45)
    if "option_count" in issue_codes and len(combined_options) < 4:
        patch.confidence = max(patch.confidence, 0.4)

    patch.should_apply = bool(
        patch.source_number
        or patch.stem is not None
        or patch.option_layout
        or patch.options
        or patch.subject_kind
        or patch.reassign_stem_assets_to_options
        or patch.move_stem_assets_to_material
        or patch.prepend_material_body
    )
    if patch.should_apply:
        patch.summary = "；".join(reasons[:3]) or "本地 AI 已生成结构化修复建议。"
    else:
        patch.summary = "本地 AI 判断当前题暂时没有足够安全的自动修复。"
    return patch


class AIRepairService:
    """本地低阶 AI 修复器：规则 + 打分 + 结构化补丁。"""

    def __init__(self, *, mode: str = "balanced") -> None:
        configured_mode = (mode or os.environ.get("PPTCONVERT_AI_REPAIR_MODE", "balanced")).strip().lower()
        self.mode = configured_mode or "balanced"

    def _policy_predictions(
        self,
        *,
        section: Section,
        material: MaterialSet | None,
        question: QuestionNode,
        previous_question: QuestionNode | None = None,
        next_question: QuestionNode | None = None,
    ) -> tuple[RepairActionPrediction | None, RepairActionPrediction | None]:
        family_prediction = predict_repair_action_distribution(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
            label_field="action_family",
        )
        action_prediction = predict_repair_action_distribution(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
            label_field="action",
        )
        return action_prediction, family_prediction

    def _trajectory_plan(
        self,
        *,
        section: Section,
        material: MaterialSet | None,
        question: QuestionNode,
        previous_question: QuestionNode | None = None,
        next_question: QuestionNode | None = None,
    ) -> TrajectoryPlan | None:
        plan = predict_trajectory_plan(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        )
        if plan is not None:
            return plan
        return build_rule_based_trajectory_plan(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        )

    def _policy_priority_score(
        self,
        *,
        section: Section,
        material: MaterialSet | None,
        question: QuestionNode,
        previous_question: QuestionNode | None = None,
        next_question: QuestionNode | None = None,
    ) -> float:
        issues = getattr(question, "review_issues", []) or []
        score = max(0.0, 1.0 - float(getattr(question, "review_confidence", 1.0) or 1.0))
        score += 0.28 * sum(1 for issue in issues if issue.severity == "error")
        score += 0.1 * sum(1 for issue in issues if issue.severity == "warning")
        action_prediction, family_prediction = self._policy_predictions(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        )
        if family_prediction is not None and family_prediction.best_label:
            family_boost = 0.95 if family_prediction.best_label in {"material_repair", "asset_repair"} else 0.7
            score += family_prediction.best_confidence * family_boost
        if action_prediction is not None and action_prediction.best_label in _REPAIR_ACTION_DISPLAY_NAMES:
            score += action_prediction.best_confidence * 0.45
        # 轨迹策略得分加成
        traj_plan = self._trajectory_plan(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        )
        if traj_plan is not None and not traj_plan.is_empty:
            score += traj_plan.primary_confidence * 0.35
            score += 0.1 * min(len(traj_plan.steps) - 1, 2)
        return score

    def prioritize_candidate_rows(
        self,
        rows: list[tuple[Section, MaterialSet | None, QuestionNode]],
    ) -> list[tuple[Section, MaterialSet | None, QuestionNode]]:
        if self.mode != "policy" or len(rows) < 2:
            return rows
        scored_rows: list[tuple[float, tuple[Section, MaterialSet | None, QuestionNode]]] = []
        for row in rows:
            section, material, question = row
            previous_question, next_question = _neighbor_questions(section, material, question)
            score = self._policy_priority_score(
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            )
            scored_rows.append((score, row))
        scored_rows.sort(key=lambda item: item[0], reverse=True)
        return [row for _score, row in scored_rows]

    def repair_question(
        self,
        *,
        section: Section,
        material: MaterialSet | None,
        question: QuestionNode,
        previous_question: QuestionNode | None = None,
        next_question: QuestionNode | None = None,
    ) -> AIRepairResult:
        patch = _suggest_local_patch(
            section=section,
            material=material,
            question=question,
            previous_question=previous_question,
            next_question=next_question,
        )
        if self.mode == "policy":
            action_prediction, family_prediction = self._policy_predictions(
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            )
            traj_plan = self._trajectory_plan(
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            )
            patch = _patch_with_policy_hint(
                patch,
                action_prediction=action_prediction,
                family_prediction=family_prediction,
                trajectory_plan=traj_plan,
            )
        return AIRepairResult(question_number=question.source_number or "-", patch=patch)


def apply_ai_question_patch(
    question: QuestionNode,
    patch: AIQuestionPatch,
    *,
    section: Section | None = None,
    material: MaterialSet | None = None,
    project: ExamProject | None = None,
) -> tuple[int, bool]:
    if not patch.should_apply:
        return 0, False

    changes = 0
    subject_changed = False

    source_number = (patch.source_number or "").strip()
    if source_number and source_number != (question.source_number or "").strip():
        renumber_question(question, source_number)
        changes += 1

    if patch.stem is not None:
        stem = (patch.stem or "").strip()
        if stem and stem != (question.stem or "").strip():
            update_question_stem(question, stem)
            changes += 1

    layout = (patch.option_layout or "").strip().lower()
    if layout in _SUPPORTED_LAYOUTS and layout != ((question.option_layout or "").strip().lower()):
        set_question_option_layout(question, layout)
        changes += 1

    if patch.reassign_stem_assets_to_options:
        changes += reassign_stem_assets_to_options(question, count=4)

    if patch.move_stem_assets_to_material and material is not None:
        changes += move_stem_assets_to_material(material, question)

    if patch.prepend_material_body and material is not None:
        if prepend_material_body_text(material, patch.prepend_material_body):
            changes += 1

    for option_patch in patch.options:
        letter = (option_patch.letter or "").strip().upper()
        if not letter:
            continue
        option = _ensure_option_exists(question, letter)
        if option is None or option_patch.text is None:
            continue
        new_text = (option_patch.text or "").strip()
        if new_text != (option.text or "").strip():
            update_option_text(question, letter, new_text)
            changes += 1

    if section is not None and patch.subject_kind and patch.subject_kind not in {section.kind, "unknown", "data"}:
        objective_count = len(section.questions) if section.kind != "data" else 0
        if section.kind == "unknown" or objective_count <= 1:
            if reclassify_objective_section(section, patch.subject_kind, project=project):
                subject_changed = True
                changes += 1

    return changes, subject_changed


def _section_question_rows(section: Section, material: MaterialSet | None) -> list[QuestionNode]:
    if section.kind == "data":
        if material is not None:
            return list(material.questions)
        rows: list[QuestionNode] = []
        for one_material in section.material_sets:
            rows.extend(one_material.questions)
        return rows
    return list(section.questions)


def _neighbor_questions(section: Section, material: MaterialSet | None, question: QuestionNode) -> tuple[QuestionNode | None, QuestionNode | None]:
    rows = _section_question_rows(section, material)
    for index, item in enumerate(rows):
        if item is question:
            previous_question = rows[index - 1] if index > 0 else None
            next_question = rows[index + 1] if index + 1 < len(rows) else None
            return previous_question, next_question
    return None, None


def repair_project_questions(
    project: ExamProject,
    *,
    service: AIRepairService,
    only_flagged: bool = True,
    limit: int | None = None,
) -> AIRepairBatchSummary:
    summary = AIRepairBatchSummary()
    changed_question_ids: set[int] = set()
    target_limit = _DEFAULT_BATCH_LIMIT if limit is None else max(0, int(limit))
    if target_limit == 0:
        return summary

    boundary_changed_question_ids, boundary_field_changes = repair_project_boundaries(project)
    if boundary_field_changes > 0:
        summary.structural_repairs = len(boundary_changed_question_ids)
        summary.total_field_changes += boundary_field_changes
        changed_question_ids.update(boundary_changed_question_ids)

    candidate_rows: list[tuple[Section, MaterialSet | None, QuestionNode]] = []
    for section, material, question in project.iter_questions():
        if only_flagged and not is_flagged_question(question):
            continue
        candidate_rows.append((section, material, question))
    candidate_rows = service.prioritize_candidate_rows(candidate_rows)[:target_limit]

    for section, material, question in candidate_rows:
        summary.attempted_questions += 1
        previous_question, next_question = _neighbor_questions(section, material, question)
        try:
            result = service.repair_question(
                section=section,
                material=material,
                question=question,
                previous_question=previous_question,
                next_question=next_question,
            )
            changes, subject_changed = apply_ai_question_patch(
                question,
                result.patch,
                section=section,
                material=material,
                project=project,
            )
        except Exception as exc:
            summary.errors.append(f"{question.source_number or '-'}：{exc}")
            continue

        if not result.patch.should_apply:
            summary.skipped_questions += 1
            continue
        if changes > 0:
            summary.total_field_changes += changes
            changed_question_ids.add(id(question))
        else:
            summary.skipped_questions += 1
        if subject_changed:
            summary.subject_changes += 1

    annotate_project_quality(project)
    summary.changed_questions = len(changed_question_ids)
    return summary


def build_ai_repair_summary_text(question: QuestionNode) -> str:
    issues = question_review_summary(question)
    score = int(round((getattr(question, "review_confidence", 1.0) or 1.0) * 100))
    label = SUBJECT_DISPLAY_NAMES.get(getattr(question, "suggested_subject", None) or "unknown", "未知科目")
    subtype = (getattr(question, "inferred_subtype", "") or "").strip()
    if getattr(question, "suggested_subject", None):
        if subtype:
            return f"当前置信度 {score}% · 主要问题：{issues} · 更像 {label} / {subtype}"
        return f"当前置信度 {score}% · 主要问题：{issues} · 更像 {label}"
    if subtype:
        return f"当前置信度 {score}% · 主要问题：{issues} · 子题型 {subtype}"
    return f"当前置信度 {score}% · 主要问题：{issues}"
