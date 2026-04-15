from __future__ import annotations

import re
import unicodedata
from collections.abc import Iterable
from dataclasses import dataclass
import os
from typing import Literal

from core.learned_subject_model import predict_subject_distribution
from core.local_ai_knowledge import (
    confidence_settings,
    subject_keywords,
    subject_negative_keywords,
    subject_structural_markers,
    subject_subtypes,
)
from domain.models import SubjectKind

OBJECTIVE_SUBJECT_KINDS: tuple[SubjectKind, ...] = (
    "politics",
    "common_sense",
    "verbal",
    "quant",
    "reasoning",
)

_DISPLAY_NAMES: dict[SubjectKind, str] = {
    "politics": "政治理论",
    "common_sense": "常识判断",
    "verbal": "言语理解与表达",
    "quant": "数量关系",
    "reasoning": "判断推理",
    "data": "资料分析",
    "unknown": "待确认科目",
}

_KEYWORDS: dict[SubjectKind, tuple[str, ...]] = {
    "politics": (
        "习近平",
        "新时代",
        "中国特色社会主义",
        "中国共产党",
        "党的",
        "党内",
        "社会主义",
        "马克思主义",
        "二十大",
        "民族复兴",
        "党章",
        "全面从严治党",
    ),
    "common_sense": (
        "下列说法",
        "我国",
        "宪法",
        "行政处罚",
        "民法典",
        "刑法",
        "地理",
        "物理",
        "化学",
        "生物",
        "历史",
        "文化常识",
        "天文",
        "节气",
        "法律",
        "科学",
    ),
    "verbal": (
        "填入划横线部分",
        "填入画横线部分",
        "最恰当",
        "最贴切",
        "词语",
        "成语",
        "语句",
        "语序",
        "排序",
        "段文字",
        "这段文字",
        "作者意在",
        "主要想表达",
        "文中",
        "理解正确",
        "阅读",
    ),
    "quant": (
        "利润",
        "成本",
        "折扣",
        "打折",
        "浓度",
        "速度",
        "路程",
        "工程",
        "甲乙",
        "平均数",
        "概率",
        "排列组合",
        "至少",
        "最多",
        "相遇",
        "追及",
        "增长了",
        "几何",
        "方程",
        "余数",
        "倍数",
    ),
    "reasoning": (
        "图形推理",
        "定义判断",
        "类比推理",
        "逻辑判断",
        "如果",
        "那么",
        "由此可以推出",
        "能够推出",
        "不能推出",
        "最能支持",
        "最能削弱",
        "加强",
        "削弱",
        "符合定义",
        "不符合定义",
        "属于",
    ),
    "data": (
        "根据下列资料",
        "根据以下资料",
        "根据所给资料",
        "资料显示",
        "同比",
        "环比",
        "百分点",
        "增长率",
        "占比",
        "比重",
        "图表",
        "下表",
        "上表",
        "图中",
        "表中",
        "材料",
    ),
}

_NUMERICISH_OPTION = re.compile(r"^[\d\s\.\-+/%％:：,，、()（）千万亿百十个元人次天小时公里吨亩顷米件台套家年以上以下左右约大于小于不超过不少于]+$")
_DIGIT_RE = re.compile(r"\d")
_PERCENT_RE = re.compile(r"%|％|百分点|同比|环比|增长率|比重")
_ORDERED_SENTENCE_MARKER_RE = re.compile(
    r"[①②③④⑤⑥⑦⑧⑨⑩]|(?:[（(]\s*\d+\s*[）)])|(?:^|[\s，,。；;])[1-9](?=[^\d])"
)
_ANALOGY_PATTERN = re.compile(r"\S+\s*之于\s*\S+\s*(?:相当于|对应)\s*\S+")
_LAW_TITLE_RE = re.compile(r"《[^》]{1,30}(?:法|条例|规定|办法)》")
_COMBO_OPTION_RE = re.compile(r"^\d+项$")
_GROUPING_OPTION_RE = re.compile(r"^(?:[1-6①②③④⑤⑥]+[，,][1-6①②③④⑤⑥]+)$")
_ORDERING_OPTION_RE = re.compile(r"^(?:[1-6]{4,6}|[①②③④⑤⑥]{4,6})$")
_SHORT_COLON_ANALOGY_RE = re.compile(r"^[^:：]{1,8}(?:[:：][^:：]{1,8}){1,2}$")
_CONSTRAINT_MARKER_RE = re.compile(r"(?:\(\d+\)|（\d+）|[①②③④⑤⑥⑦⑧⑨⑩])")
_CASE_OPTION_PREFIX_RE = re.compile(
    r"^(?:[甲乙丙丁戊己庚辛壬癸](?:某|公司|机关|单位|市民|村民|老师|同学|部门)?|小[\u4e00-\u9fff]{1,2}|老[\u4e00-\u9fff]{1,2}|[王李张赵刘陈杨黄周吴徐孙胡朱高林何郭马罗梁宋郑谢韩唐冯于董萧程曹袁邓许傅沈曾彭吕苏卢蒋蔡贾丁魏薛叶阎余潘杜戴夏钟汪田任姜范方石姚谭廖邹熊金陆郝孔白崔康毛邱秦江史顾侯邵孟龙万段雷钱汤尹黎易常武乔贺赖龚文]某)"
)
_BLANK_SLOT_RE = re.compile(r"(?:\(\s*\)|（\s*）|____+|——+|—\s*—)")
_TABLE_WORDS: tuple[str, ...] = ("表中", "下表", "如下表", "表1", "表2", "表格", "统计表")
_CHART_WORDS: tuple[str, ...] = ("图中", "下图", "如图", "柱状图", "折线图", "饼图", "图1", "图2", "图表")
_STRONG_LOGIC_MARKERS: tuple[str, ...] = (
    "最能削弱",
    "最能支持",
    "能够推出",
    "不能推出",
    "由此可以推出",
    "以下哪项如果为真",
    "以下哪项最能",
    "不能成立",
    "必然为真",
    "一定为真",
    "最能解释",
    "最不能解释",
    "要使上述推理成立",
    "要使结论成立",
    "最可能是",
    "得出结论的条件",
)
_WEAK_LOGIC_MARKERS: tuple[str, ...] = (
    "如果",
    "若",
    "那么",
    "则",
    "只有",
    "除非",
)
_SENTENCE_EXPRESSION_MARKERS: tuple[str, ...] = (
    "重新排列",
    "重新排序",
    "语序",
    "排序最",
    "衔接最",
    "填入文中",
    "依次填入",
)
_READING_MARKERS: tuple[str, ...] = (
    "这段文字",
    "文段",
    "作者意在",
    "主要想表达",
    "意在说明",
    "标题的是",
    "概括最准确",
    "这段文字告诉我们",
)
_READING_INFERENCE_MARKERS: tuple[str, ...] = (
    "据此",
    "与原文相符",
    "与文意相符",
    "可以得知",
    "能够看出",
)
_FILL_BLANK_MARKERS: tuple[str, ...] = (
    "填入画横线部分",
    "填入划横线部分",
    "依次填入",
    "最恰当",
    "最贴切",
    "最合适",
)
_GENERIC_SECTION_PREFIX = re.compile(
    r"^(?:第[一二三四五六七八九十百\d〇零]+部分|[一二三四五六七八九十百\d〇零]+[、．。.．]?|[（(]\s*[一二三四五六七八九十百\d〇零]+\s*[）)])\s*"
)
_DEFINITION_MARKERS: tuple[str, ...] = (
    "根据上述定义",
    "根据以下定义",
    "根据定义",
    "上述定义",
    "以下定义",
    "符合定义",
    "不符合定义",
    "属于下列",
    "不属于下列",
    "属于上述",
    "不属于上述",
    "定义判断",
)
_COMMON_SENSE_PATTERNS: tuple[str, ...] = (
    "下列关于",
    "关于下列",
    "下列哪一说法",
    "下列说法正确",
    "下列说法错误",
    "下列做法正确",
    "下列做法错误",
    "下列情形",
    "以下有关",
    "下列有关",
    "相关规定",
    "最可行",
    "不恰当",
    "不合适",
    "不适合",
    "不适合用作",
    "对应正确",
    "对应不正确",
)
_SCENARIO_OPTION_MARKERS: tuple[str, ...] = (
    "某",
    "张三",
    "李四",
    "王某",
    "甲公司",
    "乙公司",
    "行政机关",
    "当事人",
    "行为",
    "措施",
)
_FACTUAL_OPTION_MARKERS: tuple[str, ...] = (
    "我国",
    "可以",
    "应当",
    "不得",
    "属于",
    "不属于",
    "错误",
    "正确",
    "符合",
)
_SENTENCE_END_MARKERS: tuple[str, ...] = ("。", "；", "？", "!", "，")
_POLITICS_ANCHORS: tuple[str, ...] = (
    "习近平",
    "总书记",
    "中国式现代化",
    "中国特色社会主义",
    "中央经济工作会议",
    "中央一号文件",
    "政府工作报告",
    "中央政治局",
    "中共中央",
    "国务院",
    "五年规划",
    "教育强国",
    "科技强国",
    "农业强国",
    "高质量发展",
    "全面依法治国",
    "马克思主义",
)
_LEGAL_SCENARIO_MARKERS: tuple[str, ...] = (
    "根据我国",
    "相关规定",
    "实施条例",
    "监察法",
    "行政诉讼法",
    "民法典",
    "刑法",
    "行政处罚",
    "行政复议",
)
_ASSIGNMENT_REASONING_MARKERS: tuple[str, ...] = (
    "分别",
    "均不相同",
    "各不相同",
    "已知",
    "每人",
    "每个",
    "甲、乙、丙",
    "甲乙丙",
)
_SET_RELATION_MARKERS: tuple[str, ...] = (
    "用一个圆来表示",
    "集合",
    "关系符合下图",
    "之间的关系",
)
_RELATIONAL_REASONING_MARKERS: tuple[str, ...] = (
    "各不相同",
    "均不相同",
    "分别是",
    "都不是",
    "不同的",
    "颜色",
    "形状",
    "位置",
    "座位",
)
_RELATIONAL_ENTITY_MARKERS: tuple[str, ...] = (
    "小明",
    "小军",
    "小花",
    "小雅",
    "甲",
    "乙",
    "丙",
    "丁",
    "老师",
    "同学",
    "学生",
    "工人",
    "司机",
    "乘客",
    "顾客",
    "员工",
    "拿着",
    "气球",
)
_SUMMARY_OPTION_MARKERS: tuple[str, ...] = (
    "说明",
    "表明",
    "强调",
    "反映",
    "揭示",
    "启示",
    "趋势",
    "原因",
    "困境",
    "挑战",
    "意义",
    "作用",
    "变化",
    "本质",
    "核心",
)
_CULTURAL_KNOWLEDGE_MARKERS: tuple[str, ...] = (
    "诗词",
    "作者",
    "朝代",
    "典故",
    "地貌",
    "节气",
    "成语",
)
_QUANT_PROMPT_MARKERS: tuple[str, ...] = (
    "多少",
    "几人",
    "几天",
    "几次",
    "几处",
    "几种",
    "约为",
    "至少",
    "至多",
    "平均",
    "售价",
    "进价",
    "利润",
    "配速",
    "半径",
    "直径",
    "土方量",
    "排队",
    "存活",
    "速度",
    "效率",
)
_READING_SUMMARY_MARKERS: tuple[str, ...] = (
    "重在强调",
    "旨在说明",
    "意在强调",
    "主要强调",
    "主要说明",
    "主要告诉我们",
)
_REASONING_SUPPORT_MARKERS: tuple[str, ...] = (
    "要使上述推理成立",
    "最可能是",
    "前提",
    "条件",
    "能够支持",
    "最能支持",
    "最能削弱",
)
_QUANT_GRAPH_MARKERS: tuple[str, ...] = (
    "图线分析",
    "配速",
    "跑动距离",
    "市场规模",
    "占比情况",
    "市场占比",
    "区间",
    "半径",
    "直径",
    "土方量",
    "观测视角",
)
_CYCLE_QUANT_MARKERS: tuple[str, ...] = (
    "换座位",
    "向前换",
    "向左换",
    "回到原位",
    "回到第一排",
    "首次回到",
)
_SCIENCE_REASONING_MARKERS: tuple[str, ...] = (
    "示意图如下",
    "根据图中的信息",
    "反应过程",
    "图线",
)

_FILENAME_SUBJECT_LABELS: dict[SubjectKind, tuple[str, ...]] = {
    "politics": ("政治理论", "政治"),
    "common_sense": ("常识判断", "常识"),
    "verbal": ("言语理解与表达", "言语理解和表达", "言语理解"),
    "quant": ("数量关系",),
    "reasoning": ("判断推理",),
    "data": ("资料分析",),
}

_SINGLE_SUBJECT_FILENAME_MARKERS: tuple[str, ...] = (
    "题库",
    "专项",
    "专练",
    "专训",
    "单项",
    "单科",
    "模块",
    "分类",
    "分项",
    "刷题",
    "高频",
    "易错",
    "强化",
    "训练",
)

_SET_PAPER_FILENAME_MARKERS: tuple[str, ...] = (
    "模拟卷",
    "套卷",
    "试卷",
    "真题",
    "联考",
    "国考",
    "省考",
    "考前卷",
    "押题卷",
    "冲刺卷",
    "全真模拟",
)

_FILENAME_QUESTION_COUNT_RE = re.compile(r"\d{2,5}\s*题")


@dataclass(frozen=True)
class SubjectInferenceDiagnostics:
    kind: SubjectKind
    margin: float
    confidence: float
    best_score: float
    second_score: float
    subtype: str | None = None
    matched_signals: tuple[str, ...] = ()


@dataclass(frozen=True)
class PdfFilenameProfile:
    form: Literal["single_subject_book", "set_paper", "unknown"]
    subject_hint: SubjectKind | None = None
    confidence: float = 0.0
    matched_signals: tuple[str, ...] = ()


def default_subject_title(kind: SubjectKind) -> str:
    return _DISPLAY_NAMES.get(kind, "待确认科目")


def infer_pdf_filename_profile(filename: str) -> PdfFilenameProfile:
    stem = os.path.splitext(os.path.basename(filename or ""))[0]
    normalized = re.sub(r"\s+", "", _nfkc(stem))
    if not normalized:
        return PdfFilenameProfile(form="unknown")

    matched_subjects: list[SubjectKind] = []
    matched_signals: list[str] = []
    for kind, labels in _FILENAME_SUBJECT_LABELS.items():
        matched = next((label for label in labels if label in normalized), None)
        if matched:
            matched_subjects.append(kind)
            matched_signals.append(f"文件名命中{matched}")

    matched_subjects = list(dict.fromkeys(matched_subjects))
    has_question_count = bool(_FILENAME_QUESTION_COUNT_RE.search(normalized))
    has_single_subject_marker = _has_any(normalized, _SINGLE_SUBJECT_FILENAME_MARKERS)
    has_set_paper_marker = _has_any(normalized, _SET_PAPER_FILENAME_MARKERS)

    if len(matched_subjects) == 1:
        matched_signals.append("单科文件名")
        if has_question_count:
            matched_signals.append("题量标记")
        elif has_single_subject_marker:
            matched_signals.append("专项/题库标记")
        confidence = 0.96 if has_question_count else 0.9 if has_single_subject_marker else 0.82
        return PdfFilenameProfile(
            form="single_subject_book",
            subject_hint=matched_subjects[0],
            confidence=confidence,
            matched_signals=tuple(matched_signals[:6]),
        )

    if len(matched_subjects) >= 2:
        matched_signals.append("多科文件名")
        return PdfFilenameProfile(
            form="set_paper",
            confidence=0.92,
            matched_signals=tuple(matched_signals[:6]),
        )

    if has_set_paper_marker:
        return PdfFilenameProfile(
            form="set_paper",
            confidence=0.78,
            matched_signals=("套卷标记",),
        )

    return PdfFilenameProfile(form="unknown")


def _nfkc(text: str) -> str:
    return unicodedata.normalize("NFKC", text or "")


def normalize_subject_title_for_compare(kind: SubjectKind, title: str) -> str:
    label = default_subject_title(kind)
    text = re.sub(r"\s+", "", _nfkc(title))
    if not text:
        return label
    previous = None
    while previous != text:
        previous = text
        text = _GENERIC_SECTION_PREFIX.sub("", text)
        text = text.strip("：:·-—_、.．。")
    return text or label


def is_generic_subject_title(kind: SubjectKind, title: str) -> bool:
    normalized = normalize_subject_title_for_compare(kind, title)
    label = default_subject_title(kind)
    return normalized in {label, f"{label}题"}


def preferred_subject_title(kind: SubjectKind, *titles: str) -> str:
    candidates = [title.strip() for title in titles if (title or "").strip()]
    if not candidates:
        return default_subject_title(kind)

    explicit = [title for title in candidates if not is_generic_subject_title(kind, title)]
    if explicit:
        return explicit[0]

    generic = [
        title
        for title in candidates
        if normalize_subject_title_for_compare(kind, title) == default_subject_title(kind)
    ]
    if generic:
        return max(generic, key=len)
    return candidates[0]


def should_merge_subject_sections(kind: SubjectKind, left_title: str, right_title: str) -> bool:
    if kind in {"data", "unknown"}:
        return False
    left = (left_title or "").strip()
    right = (right_title or "").strip()
    if not left or not right:
        return True
    if normalize_subject_title_for_compare(kind, left) == normalize_subject_title_for_compare(kind, right):
        return True
    if is_generic_subject_title(kind, left) or is_generic_subject_title(kind, right):
        return True
    return False


def _clean_text(parts: Iterable[str]) -> str:
    joined = "\n".join(part.strip() for part in parts if (part or "").strip())
    return re.sub(r"\s+", " ", joined).strip()


def _combined_keywords(kind: SubjectKind) -> tuple[str, ...]:
    return tuple(dict.fromkeys((*_KEYWORDS.get(kind, ()), *subject_keywords(kind))))


def _combined_negative_keywords(kind: SubjectKind) -> tuple[str, ...]:
    return tuple(dict.fromkeys(subject_negative_keywords(kind)))


def _combined_structural_markers(kind: SubjectKind) -> tuple[str, ...]:
    return tuple(dict.fromkeys(subject_structural_markers(kind)))


def _append_signal(matched_signals: dict[SubjectKind, list[str]], kind: SubjectKind, label: str) -> None:
    if not label:
        return
    current = matched_signals.setdefault(kind, [])
    if label not in current and len(current) < 10:
        current.append(label)


def _blend_learned_subject_scores(
    *,
    scores: dict[SubjectKind, float],
    matched_signals: dict[SubjectKind, list[str]],
    stem: str,
    options: Iterable[str],
    material_text: str,
    image_count: int,
    material_header: str,
    allow_data: bool,
) -> None:
    prediction = predict_subject_distribution(
        stem=stem,
        options=options,
        material_text=material_text,
        image_count=image_count,
        material_header=material_header,
    )
    if prediction is None or not prediction.probabilities:
        return

    active_kinds = [kind for kind in scores if kind != "data" or allow_data]
    if not active_kinds:
        return
    uniform = 1.0 / len(active_kinds)
    for kind, probability in prediction.probabilities.items():
        if kind not in scores:
            continue
        if kind == "data" and not allow_data:
            continue
        scores[kind] += (float(probability) - uniform) * 4.0
        if probability >= max(0.34, uniform + 0.08):
            _append_signal(
                matched_signals,
                kind,
                f"学习模型:{_DISPLAY_NAMES.get(kind, kind)} {float(probability):.0%}",
            )


def _has_any(text: str, needles: Iterable[str]) -> bool:
    return any(needle and needle in text for needle in needles)


def _sentence_sequence_count(text: str) -> int:
    return len(_ORDERED_SENTENCE_MARKER_RE.findall(text or ""))


def _blank_slot_count(text: str) -> int:
    return len(_BLANK_SLOT_RE.findall(text or ""))


def _average_option_length(option_texts: tuple[str, ...]) -> float:
    lengths = [len((text or "").strip()) for text in option_texts if (text or "").strip()]
    if not lengths:
        return 0.0
    return sum(lengths) / len(lengths)


def _looks_like_statement_option(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    if len(value) >= 12:
        return True
    return _has_any(value, _SENTENCE_END_MARKERS) or _has_any(value, _FACTUAL_OPTION_MARKERS)


def _scenario_option_count(option_texts: tuple[str, ...]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if len(value) < 8:
            continue
        if _CASE_OPTION_PREFIX_RE.search(value) or _has_any(value, _SCENARIO_OPTION_MARKERS):
            count += 1
    return count


def _statement_option_count(option_texts: tuple[str, ...]) -> int:
    return sum(1 for text in option_texts if _looks_like_statement_option(text))


def _combo_option_count(option_texts: tuple[str, ...]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if _looks_like_combo_option(value):
            count += 1
    return count


def _grouping_option_count(option_texts: tuple[str, ...]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if _GROUPING_OPTION_RE.fullmatch(value):
            count += 1
    return count


def _ordering_option_count(option_texts: tuple[str, ...]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if _ORDERING_OPTION_RE.fullmatch(value):
            count += 1
    return count


def _looks_like_combo_option(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    if _COMBO_OPTION_RE.fullmatch(value):
        return True
    circled = "①②③④⑤⑥⑦⑧⑨⑩"
    if all(ch in circled for ch in value) and 2 <= len(value) <= 6:
        order = [circled.index(ch) for ch in value]
        return len(order) == len(set(order)) and order == sorted(order)
    if value.isdigit() and 2 <= len(value) <= 6 and all(ch in "123456" for ch in value):
        order = [int(ch) for ch in value]
        return len(order) == len(set(order)) and order == sorted(order)
    return False


def _pure_numeric_option_count(option_texts: tuple[str, ...]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if not value or not _is_numericish(value):
            continue
        if _looks_like_combo_option(value):
            continue
        if _GROUPING_OPTION_RE.fullmatch(value):
            continue
        if _ORDERING_OPTION_RE.fullmatch(value):
            continue
        count += 1
    return count


def _summary_option_count(option_texts: tuple[str, ...]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if len(value) < 8:
            continue
        if _has_any(value, _SUMMARY_OPTION_MARKERS):
            count += 1
    return count


def _looks_like_analogy_stem(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    return bool(_SHORT_COLON_ANALOGY_RE.fullmatch(value))


def _constraint_clause_count(text: str) -> int:
    return len(_CONSTRAINT_MARKER_RE.findall(text or ""))


def _logic_signal_profile(
    full_text: str,
    *,
    statement_option_count: int,
    blank_signal: bool,
    pure_numeric_option_count: int,
    long_stem: bool,
) -> tuple[bool, bool]:
    strong_logic_signal = _has_any(full_text, _STRONG_LOGIC_MARKERS)
    has_if_like = "如果" in full_text or "若" in full_text
    has_then_like = "那么" in full_text or "则" in full_text
    weak_conditional_logic = (
        (has_if_like and has_then_like)
        or ("只有" in full_text)
        or ("除非" in full_text)
    )
    logic_signal = strong_logic_signal or (
        weak_conditional_logic
        and statement_option_count >= 2
        and not blank_signal
        and pure_numeric_option_count < 2
        and not long_stem
    )
    return logic_signal, strong_logic_signal


def _score_subject_subtype(
    kind: SubjectKind,
    *,
    full_text: str,
    stem_text: str,
    option_texts: tuple[str, ...],
    material_text: str,
    image_count: int,
    matched_signals: dict[SubjectKind, list[str]],
) -> tuple[str | None, float]:
    lowered = full_text.lower()
    best_name: str | None = None
    best_score = 0.0
    sequence_count = _sentence_sequence_count(stem_text)
    analogy_signal = (
        bool(_ANALOGY_PATTERN.search(full_text))
        or ("相当于" in full_text and ("之于" in full_text or "对于" in full_text))
        or _looks_like_analogy_stem(stem_text)
    )
    table_signal = _has_any(full_text, _TABLE_WORDS)
    chart_signal = _has_any(full_text, _CHART_WORDS)
    short_word_options = bool(
        option_texts
        and len(option_texts) >= 4
        and all(1 <= len(text) <= 8 for text in option_texts if text)
    )
    average_option_length = _average_option_length(tuple(option_texts))
    statement_option_count = _statement_option_count(tuple(option_texts))
    scenario_option_count = _scenario_option_count(tuple(option_texts))
    ordering_option_count = _ordering_option_count(tuple(option_texts))
    grouping_option_count = _grouping_option_count(tuple(option_texts))
    blank_signal = bool(_BLANK_SLOT_RE.search(stem_text))
    pure_numeric_option_count = _pure_numeric_option_count(tuple(option_texts))
    long_stem = len(stem_text) >= 42
    logic_signal, strong_logic_signal = _logic_signal_profile(
        full_text,
        statement_option_count=statement_option_count,
        blank_signal=blank_signal,
        pure_numeric_option_count=pure_numeric_option_count,
        long_stem=long_stem,
    )
    constraint_signal = _constraint_clause_count(full_text) >= 2 and _has_any(full_text, _ASSIGNMENT_REASONING_MARKERS)
    set_relation_signal = _has_any(full_text, _SET_RELATION_MARKERS)
    graphic_signal = bool(
        image_count >= 4
        and len(stem_text) <= 40
        and len(option_texts) >= 4
        and all(len(text) <= 10 for text in option_texts if text)
    )
    for subtype in subject_subtypes(kind):
        name = str(subtype.get("name", "")).strip()
        if not name:
            continue
        subtype_score = 0.0
        for keyword in subtype.get("positive_keywords", []) or []:
            value = str(keyword).strip()
            if value and value.lower() in lowered:
                subtype_score += 1.0 if len(value) >= 4 else 0.65
        for marker in subtype.get("structural_markers", []) or []:
            value = str(marker).strip()
            if value and value in stem_text:
                subtype_score += 0.9
        if kind == "verbal" and name == "语句表达":
            if sequence_count >= 3 and _has_any(full_text, _SENTENCE_EXPRESSION_MARKERS):
                subtype_score += 2.0
            if ordering_option_count >= 3:
                subtype_score += 1.2
        elif kind == "verbal" and name == "片段阅读":
            if len(stem_text) >= 40 and _has_any(full_text, _READING_MARKERS):
                subtype_score += 1.9
            if statement_option_count >= 3:
                subtype_score += 0.5
        elif kind == "verbal" and name == "逻辑填空":
            if (_has_any(full_text, _FILL_BLANK_MARKERS) or blank_signal) and (
                short_word_options or average_option_length <= 10
            ):
                subtype_score += 2.4
        elif kind == "reasoning" and name == "定义判断":
            if "是指" in stem_text and ("下列" in full_text or "哪一项" in full_text or "属于" in full_text):
                subtype_score += 2.1
            if scenario_option_count >= 2:
                subtype_score += 0.7
        elif kind == "reasoning" and name == "类比推理":
            if analogy_signal:
                subtype_score += 2.2
        elif kind == "reasoning" and name == "逻辑判断":
            if logic_signal or constraint_signal or set_relation_signal:
                subtype_score += 1.7
            if strong_logic_signal:
                subtype_score += 0.6
            if statement_option_count >= 3:
                subtype_score += 0.5
        elif kind == "reasoning" and name == "图形推理":
            if graphic_signal or ("问号处" in full_text and image_count > 0):
                subtype_score += 2.1
            if grouping_option_count >= 3 and "分为两类" in full_text:
                subtype_score += 1.0
        elif kind == "data" and name == "表格型资料分析":
            if table_signal or (_has_any(material_text, _TABLE_WORDS) and image_count > 0):
                subtype_score += 1.9
        elif kind == "data" and name == "图形型资料分析":
            if chart_signal or (_has_any(material_text, _CHART_WORDS) and image_count > 0):
                subtype_score += 1.9
        elif kind == "data" and name == "综合型资料分析":
            mixed_table_chart = (
                (table_signal and chart_signal)
                or (table_signal and image_count >= 2 and _has_any(material_text, _CHART_WORDS))
                or (chart_signal and image_count >= 2 and _has_any(material_text, _TABLE_WORDS))
            )
            if material_text and mixed_table_chart:
                subtype_score += 2.2
        elif kind == "data" and name == "文字型资料分析":
            if material_text and not table_signal and not chart_signal:
                subtype_score += 1.2
        if subtype_score > best_score:
            best_score = subtype_score
            best_name = name
    if best_name and best_score >= 1.15:
        _append_signal(matched_signals, kind, best_name)
        return best_name, min(2.2, best_score * 0.45)
    return None, 0.0


def _apply_subject_structural_markers(
    *,
    kind: SubjectKind,
    scores: dict[SubjectKind, float],
    matched_signals: dict[SubjectKind, list[str]],
    digit_count: int,
    numeric_option_count: int,
    sequence_count: int,
    long_stem: bool,
    image_count: int,
    material_text: str,
    material_header: str,
    definition_signal: bool,
    analogy_signal: bool,
    logic_signal: bool,
    knowledge_prompt_signal: bool,
    constraint_signal: bool,
    set_relation_signal: bool,
    table_signal: bool,
    chart_signal: bool,
    graphic_signal: bool,
    full_text: str,
) -> None:
    for marker in _combined_structural_markers(kind):
        value = (marker or "").strip()
        if not value:
            continue
        bonus = 0.0
        if value in {"时政理论", "党政理论", "政治理论"}:
            if _has_any(full_text, ("习近平", "新时代", "中国式现代化", "全面从严治党", "高质量发展")):
                bonus = 0.9
        elif value in {"知识判断", "法律常识", "科技常识"}:
            if (
                not definition_signal
                and not analogy_signal
                and not logic_signal
                and not constraint_signal
                and not set_relation_signal
                and (knowledge_prompt_signal or _has_any(full_text, ("宪法", "刑法", "地理", "历史", "科技")))
            ):
                bonus = 0.8
        elif value == "横线填空":
            if _has_any(full_text, _FILL_BLANK_MARKERS):
                bonus = 1.0
        elif value == "片段阅读":
            if long_stem and _has_any(full_text, _READING_MARKERS):
                bonus = 0.9
        elif value in {"语句排序", "句子排序", "衔接填空"}:
            if sequence_count >= 3 and _has_any(full_text, _SENTENCE_EXPRESSION_MARKERS):
                bonus = 1.1
        elif value == "数字密集":
            if digit_count >= 10 and numeric_option_count >= 1:
                bonus = 0.9
        elif value == "公式计算":
            if digit_count >= 6 and numeric_option_count >= 2:
                bonus = 0.75
        elif value == "纯数字选项":
            if numeric_option_count >= 3:
                bonus = 1.0
        elif value in {"定义在前案例在后", "定义匹配"}:
            if definition_signal:
                bonus = 1.0
        elif value in {"图形问号位", "图形矩阵", "图形序列"}:
            if graphic_signal or ("问号处" in full_text and image_count > 0):
                bonus = 1.1
        elif value in {"条件推理", "加强削弱"}:
            if logic_signal:
                bonus = 0.95
        elif value in {"词项配对", "关系映射"}:
            if analogy_signal:
                bonus = 1.0
        elif value == "共享材料":
            if material_text or material_header:
                bonus = 1.0
        elif value in {"表格图表", "图表材料"}:
            if table_signal or chart_signal:
                bonus = 1.0
        elif value == "一段材料带多题":
            if material_text and (table_signal or chart_signal or digit_count >= 10):
                bonus = 0.9
        if bonus > 0:
            scores[kind] += bonus
            _append_signal(matched_signals, kind, value)


def _calibrate_subject_confidence(best_score: float, second_score: float, matched_count: int) -> float:
    settings = confidence_settings()
    display_floor = settings.get("display_floor", 0.18)
    display_cap = settings.get("display_cap", 0.98)
    unknown_cap = settings.get("unknown_cap", 0.34)
    low_margin_cap = settings.get("low_margin_cap", 0.52)

    best_score = max(0.0, best_score)
    second_score = max(0.0, second_score)
    margin = max(0.0, best_score - second_score)
    absolute_strength = min(best_score / 5.8, 1.0)
    gap_strength = min(margin / max(best_score, 1.0), 1.0)
    signal_bonus = min(matched_count, 4) * 0.05
    confidence = display_floor + absolute_strength * 0.38 + gap_strength * 0.32 + signal_bonus
    confidence = min(display_cap, max(0.0, confidence))
    if best_score < 1.35:
        return min(confidence, unknown_cap)
    if best_score < 2.0 and margin < 0.55:
        return min(confidence, low_margin_cap)
    return confidence


def _is_numericish(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    return bool(_NUMERICISH_OPTION.match(value))


def _parse_source_number(value: str | None) -> int | None:
    text = (value or "").strip()
    return int(text) if text.isdigit() else None


def _smooth_subject_kinds(
    kinds: list[SubjectKind],
    confidences: list[float],
) -> list[SubjectKind]:
    if len(kinds) < 3:
        return kinds
    smoothed = list(kinds)
    for index in range(1, len(smoothed) - 1):
        prev_kind = smoothed[index - 1]
        kind = smoothed[index]
        next_kind = smoothed[index + 1]
        if kind != prev_kind and prev_kind == next_kind and confidences[index] < 1.4:
            smoothed[index] = prev_kind
    return smoothed


def _contextualize_subject_kinds(
    kinds: list[SubjectKind],
    confidences: list[float],
) -> list[SubjectKind]:
    if len(kinds) < 3:
        return kinds
    resolved = list(kinds)
    for index in range(len(resolved)):
        left = resolved[index - 1] if index - 1 >= 0 else None
        right = resolved[index + 1] if index + 1 < len(resolved) else None
        if (
            resolved[index] == "unknown"
            and left
            and right
            and left == right
            and left != "unknown"
            and min(confidences[max(0, index - 1)], confidences[min(len(confidences) - 1, index + 1)]) >= 0.55
        ):
            resolved[index] = left
            continue
        if (
            1 <= index < len(resolved) - 1
            and left
            and right
            and left == right
            and resolved[index] != left
            and resolved[index] != "unknown"
            and confidences[index] < 0.7
        ):
            resolved[index] = left
            continue
        if (
            2 <= index < len(resolved) - 2
            and resolved[index - 2] == resolved[index - 1] == resolved[index + 1] == resolved[index + 2]
            and resolved[index - 1] != "unknown"
            and confidences[index] < 0.82
        ):
            resolved[index] = resolved[index - 1]
    return resolved


def _override_gap(
    source_numbers: list[str],
    start: int,
    end: int,
) -> int:
    prev_number = _parse_source_number(source_numbers[start - 1]) if start > 0 else None
    first_number = _parse_source_number(source_numbers[start])
    last_number = _parse_source_number(source_numbers[end - 1])
    next_number = _parse_source_number(source_numbers[end]) if end < len(source_numbers) else None

    gaps: list[int] = []
    if prev_number is not None and first_number is not None and first_number > prev_number:
        gaps.append(first_number - prev_number)
    if next_number is not None and last_number is not None and next_number > last_number:
        gaps.append(next_number - last_number)
    return max(gaps, default=0)


def _allow_override_run(
    default_kind: SubjectKind,
    override_kind: SubjectKind,
    *,
    confidences: list[float],
    source_numbers: list[str],
    start: int,
    end: int,
    strict_default: bool = False,
) -> bool:
    run_len = end - start
    avg_confidence = sum(confidences[start:end]) / max(run_len, 1)
    max_confidence = max(confidences[start:end], default=0.0)
    at_edge = start == 0 or end == len(source_numbers)
    gap = _override_gap(source_numbers, start, end)
    common_sense_reasoning_pair = {default_kind, override_kind} == {"common_sense", "reasoning"}

    if strict_default:
        if run_len >= 6 and avg_confidence >= 0.95:
            return True
        if run_len >= 4 and avg_confidence >= 1.15 and (gap >= 4 or at_edge):
            return True
        if run_len == 1 and max_confidence >= 2.2 and gap >= 10 and at_edge:
            return True
        return False

    if common_sense_reasoning_pair:
        if run_len >= 4 and avg_confidence >= 0.95:
            return True
        if run_len >= 2 and avg_confidence >= 1.2 and (gap >= 3 or at_edge):
            return True
        if run_len == 1 and max_confidence >= 2.1 and gap >= 8:
            return True
        return False

    if run_len >= 3 and avg_confidence >= 0.85:
        return True
    if run_len >= 2 and avg_confidence >= 1.0 and (gap >= 2 or at_edge):
        return True
    if run_len == 1 and max_confidence >= 1.6 and gap >= 6:
        return True
    return False


def resolve_objective_section_kinds(
    *,
    default_kind: SubjectKind,
    inferred_pairs: list[tuple[SubjectKind, float]],
    source_numbers: list[str],
    strong_text_signals: list[bool],
    strict_default: bool = False,
) -> list[SubjectKind]:
    if not inferred_pairs:
        return []

    proposed: list[SubjectKind] = []
    confidences = [confidence for _kind, confidence in inferred_pairs]
    for (inferred_kind, confidence), strong_text_signal in zip(inferred_pairs, strong_text_signals):
        if default_kind == "unknown":
            proposed.append(inferred_kind if inferred_kind != "unknown" else "unknown")
            continue
        if (
            inferred_kind not in {"unknown", default_kind}
            and confidence >= 0.9
            and strong_text_signal
        ):
            proposed.append(inferred_kind)
        else:
            proposed.append(default_kind)

    proposed = _contextualize_subject_kinds(_smooth_subject_kinds(proposed, confidences), confidences)
    if default_kind == "unknown":
        return _contextualize_subject_kinds(proposed, confidences)

    resolved = list(proposed)
    index = 0
    while index < len(resolved):
        run_kind = resolved[index]
        run_end = index + 1
        while run_end < len(resolved) and resolved[run_end] == run_kind:
            run_end += 1
        if run_kind not in {"unknown", default_kind}:
            if not _allow_override_run(
                default_kind,
                run_kind,
                confidences=confidences,
                source_numbers=source_numbers,
                start=index,
                end=run_end,
                strict_default=strict_default,
            ):
                for replace_index in range(index, run_end):
                    resolved[replace_index] = default_kind
        index = run_end

    return _contextualize_subject_kinds(_smooth_subject_kinds(resolved, confidences), confidences)


def infer_subject_diagnostics(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    image_count: int = 0,
    material_header: str = "",
    allow_data: bool = True,
) -> SubjectInferenceDiagnostics:
    stem_text = _clean_text([material_header, material_text, stem])
    option_texts = [text.strip() for text in (options or []) if (text or "").strip()]
    option_blob = _clean_text(option_texts)
    full_text = _clean_text([stem_text, option_blob])
    if not full_text and image_count <= 0:
        return SubjectInferenceDiagnostics(
            kind="unknown",
            margin=0.0,
            confidence=0.0,
            best_score=0.0,
            second_score=0.0,
        )

    scores: dict[SubjectKind, float] = {
        "politics": 0.0,
        "common_sense": 0.0,
        "verbal": 0.0,
        "quant": 0.0,
        "reasoning": 0.0,
        "data": 0.0,
    }
    matched_signals: dict[SubjectKind, list[str]] = {kind: [] for kind in scores}

    lowered = full_text.lower()
    for kind in scores:
        if kind == "data" and not allow_data:
            continue
        keywords = _combined_keywords(kind)
        for keyword in keywords:
            if keyword and keyword.lower() in lowered:
                scores[kind] += 1.25 if len(keyword) >= 4 else 0.8
                _append_signal(matched_signals, kind, keyword)
        for keyword in _combined_negative_keywords(kind):
            if keyword and keyword.lower() in lowered:
                scores[kind] -= 0.8 if len(keyword) >= 4 else 0.45

    digit_count = len(_DIGIT_RE.findall(full_text))
    percent_hits = len(_PERCENT_RE.findall(full_text))
    numeric_option_count = sum(1 for text in option_texts if _is_numericish(text))
    short_word_options = (
        len(option_texts) >= 4
        and all(len(text.strip()) <= 8 for text in option_texts if text.strip())
        and digit_count <= 3
    )
    statement_option_count = _statement_option_count(tuple(option_texts))
    scenario_option_count = _scenario_option_count(tuple(option_texts))
    combo_option_count = _combo_option_count(tuple(option_texts))
    grouping_option_count = _grouping_option_count(tuple(option_texts))
    ordering_option_count = _ordering_option_count(tuple(option_texts))
    pure_numeric_option_count = _pure_numeric_option_count(tuple(option_texts))
    summary_option_count = _summary_option_count(tuple(option_texts))
    average_option_length = _average_option_length(tuple(option_texts))
    long_stem = len(stem_text) >= 42
    definition_signal = any(marker in full_text for marker in _DEFINITION_MARKERS)
    common_sense_signal = any(pattern in full_text for pattern in _COMMON_SENSE_PATTERNS)
    law_title_signal = bool(_LAW_TITLE_RE.search(full_text))
    table_signal = _has_any(full_text, _TABLE_WORDS)
    chart_signal = _has_any(full_text, _CHART_WORDS)
    sequence_count = _sentence_sequence_count(stem_text)
    analogy_signal = (
        bool(_ANALOGY_PATTERN.search(full_text))
        or ("相当于" in full_text and ("之于" in full_text or "对于" in full_text))
        or _looks_like_analogy_stem(stem_text)
    )
    blank_signal = bool(_BLANK_SLOT_RE.search(stem_text))
    blank_slot_count = _blank_slot_count(stem_text)
    logic_signal, strong_logic_signal = _logic_signal_profile(
        full_text,
        statement_option_count=statement_option_count,
        blank_signal=blank_signal,
        pure_numeric_option_count=pure_numeric_option_count,
        long_stem=long_stem,
    )
    knowledge_prompt_signal = common_sense_signal or _has_any(full_text, _LEGAL_SCENARIO_MARKERS)
    constraint_signal = _constraint_clause_count(full_text) >= 2 and _has_any(full_text, _ASSIGNMENT_REASONING_MARKERS)
    set_relation_signal = _has_any(full_text, _SET_RELATION_MARKERS)
    culture_knowledge_signal = _has_any(full_text, _CULTURAL_KNOWLEDGE_MARKERS)
    relational_reasoning_signal = _has_any(full_text, _RELATIONAL_REASONING_MARKERS) and (
        (_has_any(full_text, _RELATIONAL_ENTITY_MARKERS) and scenario_option_count >= 1)
        or "说:" in full_text
        or "说：" in full_text
        or "每人" in full_text
        or "每个" in full_text
    )
    quant_prompt_signal = _has_any(full_text, _QUANT_PROMPT_MARKERS)
    reading_summary_signal = _has_any(full_text, _READING_SUMMARY_MARKERS)
    reasoning_support_signal = _has_any(full_text, _REASONING_SUPPORT_MARKERS)
    quant_graph_signal = _has_any(full_text, _QUANT_GRAPH_MARKERS)
    science_reasoning_signal = _has_any(full_text, _SCIENCE_REASONING_MARKERS)
    cycle_quant_signal = _has_any(full_text, _CYCLE_QUANT_MARKERS)
    graphic_signal = bool(
        image_count >= 4
        and len(stem_text) <= 40
        and len(option_texts) >= 4
        and all(len(text) <= 10 for text in option_texts if text)
    )
    phrase_like_options = bool(
        len(option_texts) >= 4
        and average_option_length <= 10
        and pure_numeric_option_count == 0
    )
    if "是指" in stem_text and ("下列" in full_text or "哪一项" in full_text):
        definition_signal = True

    if pure_numeric_option_count >= 3:
        scores["quant"] += 2.0
        _append_signal(matched_signals, "quant", "纯数字选项")
        if allow_data:
            scores["data"] += 1.0
            _append_signal(matched_signals, "data", "数字/比例密集")
    if digit_count >= 10 and combo_option_count < 3 and grouping_option_count < 3 and ordering_option_count < 3:
        scores["quant"] += 1.2
        _append_signal(matched_signals, "quant", "题干数字密集")
    if digit_count >= 18:
        scores["data"] += 1.0 if allow_data else 0.0
        if allow_data:
            _append_signal(matched_signals, "data", "数据量较大")
    if percent_hits >= 1 and allow_data:
        scores["data"] += 1.5
        _append_signal(matched_signals, "data", "同比/比重/百分点")
    if table_signal and allow_data:
        scores["data"] += 1.6
        _append_signal(matched_signals, "data", "表格材料")
    if chart_signal and allow_data:
        scores["data"] += 1.6
        _append_signal(matched_signals, "data", "图形材料")
    if material_header:
        scores["data"] += 4.0 if allow_data else 0.0
        if allow_data:
            _append_signal(matched_signals, "data", "材料标题")
    if material_text:
        scores["data"] += 2.4 if allow_data else 0.0
        if allow_data:
            _append_signal(matched_signals, "data", "共享材料正文")
    if image_count and allow_data and material_text:
        scores["data"] += 1.0
        _append_signal(matched_signals, "data", "材料带图表")
    if image_count and not material_text and digit_count <= 4:
        scores["reasoning"] += 1.2
        _append_signal(matched_signals, "reasoning", "图形/图片题")
    if graphic_signal:
        scores["reasoning"] += 2.0
        _append_signal(matched_signals, "reasoning", "图形推理结构")
    if law_title_signal and not definition_signal:
        scores["common_sense"] += 1.2
        _append_signal(matched_signals, "common_sense", "法律条文题")
    if long_stem and digit_count <= 8 and not knowledge_prompt_signal:
        scores["verbal"] += 0.9
        _append_signal(matched_signals, "verbal", "长文段题干")
    if _has_any(full_text, _READING_MARKERS):
        scores["verbal"] += 0.8
        _append_signal(matched_signals, "verbal", "片段阅读设问")
    if _has_any(full_text, _FILL_BLANK_MARKERS) and (short_word_options or phrase_like_options):
        scores["verbal"] += 0.8
        _append_signal(matched_signals, "verbal", "逻辑填空设问")
    if blank_signal and (short_word_options or phrase_like_options):
        blank_bonus = 1.8
        if blank_slot_count >= 2:
            blank_bonus += 0.9
        if long_stem:
            blank_bonus += 0.45
        scores["verbal"] += blank_bonus
        scores["common_sense"] -= 0.8
        scores["reasoning"] -= 0.8
        _append_signal(matched_signals, "verbal", "填空结构")
    if sequence_count >= 3 and ordering_option_count >= 3:
        scores["verbal"] += 2.3
        scores["quant"] -= 1.2
        _append_signal(matched_signals, "verbal", "排序型选项")
    if long_stem and summary_option_count >= 2 and not law_title_signal and not definition_signal and not constraint_signal:
        scores["verbal"] += 1.2
        _append_signal(matched_signals, "verbal", "概括型选项")
    if long_stem and statement_option_count >= 3 and digit_count <= 6 and not knowledge_prompt_signal:
        scores["verbal"] += 1.0
        if not strong_logic_signal:
            scores["reasoning"] -= 0.5
        _append_signal(matched_signals, "verbal", "整题阅读型结构")
    if (
        long_stem
        and statement_option_count >= 3
        and _has_any(full_text, _READING_INFERENCE_MARKERS)
        and not definition_signal
        and not law_title_signal
        and not relational_reasoning_signal
    ):
        scores["verbal"] += 1.8
        scores["common_sense"] -= 1.0
        _append_signal(matched_signals, "verbal", "阅读判断问法")
    if _has_any(full_text, ("与原文相符", "与原文不符", "与文意相符", "与文意不符")) and not definition_signal:
        scores["verbal"] += 2.0
        scores["common_sense"] -= 1.2
        _append_signal(matched_signals, "verbal", "原文比对问法")
    if (
        long_stem
        and statement_option_count >= 3
        and reading_summary_signal
        and not definition_signal
        and not law_title_signal
    ):
        scores["verbal"] += 1.7
        scores["common_sense"] -= 0.9
        _append_signal(matched_signals, "verbal", "主旨概括问法")
    if (
        long_stem
        and digit_count <= 12
        and pure_numeric_option_count == 0
        and combo_option_count == 0
        and not table_signal
        and not chart_signal
        and not material_text
        and not knowledge_prompt_signal
        and not strong_logic_signal
    ):
        scores["verbal"] += 0.9
        _append_signal(matched_signals, "verbal", "长文段语义题")
    if (
        long_stem
        and phrase_like_options
        and statement_option_count == 0
        and combo_option_count == 0
        and not definition_signal
        and not analogy_signal
        and not quant_prompt_signal
    ):
        scores["verbal"] += 2.0
        scores["common_sense"] -= 0.8
        _append_signal(matched_signals, "verbal", "隐式语义填空")
    if sequence_count >= 3 and _has_any(full_text, _SENTENCE_EXPRESSION_MARKERS):
        scores["verbal"] += 2.1
        _append_signal(matched_signals, "verbal", "语句表达结构")
    if definition_signal:
        scores["reasoning"] += 3.0
        scores["common_sense"] -= 1.0
        _append_signal(matched_signals, "reasoning", "定义判断结构")
    if definition_signal and scenario_option_count >= 2:
        scores["reasoning"] += 0.9
        scores["common_sense"] -= 0.4
        _append_signal(matched_signals, "reasoning", "案例型选项")
    if analogy_signal:
        scores["reasoning"] += 2.2
        scores["common_sense"] -= 0.6
        _append_signal(matched_signals, "reasoning", "类比推理结构")
    if grouping_option_count >= 3 and ("分为两类" in full_text or image_count > 0):
        scores["reasoning"] += 2.0
        scores["quant"] -= 1.2
        _append_signal(matched_signals, "reasoning", "分组分类选项")
    if relational_reasoning_signal and not definition_signal:
        scores["reasoning"] += 3.0
        scores["common_sense"] -= 1.4
        _append_signal(matched_signals, "reasoning", "关系约束题干")
    if len(option_texts) >= 4 and all(len(text) <= 8 for text in option_texts) and digit_count <= 3:
        scores["reasoning"] += 0.6
        scores["common_sense"] += 0.3
        _append_signal(matched_signals, "reasoning", "短选项对比")
    if scenario_option_count >= 2 and _has_any(full_text, _LEGAL_SCENARIO_MARKERS) and not definition_signal:
        scores["common_sense"] += 1.2
        _append_signal(matched_signals, "common_sense", "法条情景判断")
    if "下列" in full_text and "说法" in full_text and not relational_reasoning_signal:
        scores["common_sense"] += 1.0
        _append_signal(matched_signals, "common_sense", "下列说法")
    if common_sense_signal and not blank_signal and not relational_reasoning_signal:
        scores["common_sense"] += 1.3
        _append_signal(matched_signals, "common_sense", "常识判断问法")
    if (
        knowledge_prompt_signal
        and statement_option_count >= 2
        and not definition_signal
        and not logic_signal
        and not blank_signal
        and not relational_reasoning_signal
    ):
        scores["common_sense"] += 0.8
        _append_signal(matched_signals, "common_sense", "知识问法")
    if (
        culture_knowledge_signal
        and ("正确" in full_text or "错误" in full_text or "不恰当" in full_text or "最可行" in full_text)
        and not relational_reasoning_signal
    ):
        scores["common_sense"] += 0.9
        _append_signal(matched_signals, "common_sense", "文化科技常识")
    if (
        common_sense_signal
        and statement_option_count >= 3
        and not definition_signal
        and not logic_signal
        and not relational_reasoning_signal
    ):
        scores["common_sense"] += 0.7
        _append_signal(matched_signals, "common_sense", "知识判断型选项")
    politics_anchor_count = sum(1 for marker in _POLITICS_ANCHORS if marker and marker in full_text)
    politics_anchor = politics_anchor_count >= 1
    strong_politics_anchor = politics_anchor_count >= 2 or ("习近平" in full_text and "总书记" in full_text)
    if politics_anchor and (statement_option_count >= 3 or combo_option_count >= 3):
        scores["politics"] += 1.1 if strong_politics_anchor else 0.9
        _append_signal(matched_signals, "politics", "政策理论判断")
    if combo_option_count >= 3 and politics_anchor:
        scores["politics"] += 1.0
        scores["quant"] -= 1.1
        _append_signal(matched_signals, "politics", "组合项选项")
    elif combo_option_count >= 3 and (common_sense_signal or law_title_signal or knowledge_prompt_signal):
        scores["common_sense"] += 0.8
        scores["quant"] -= 1.0
        _append_signal(matched_signals, "common_sense", "组合项知识判断")
    if strong_politics_anchor and knowledge_prompt_signal:
        scores["politics"] += 1.5
        scores["common_sense"] -= 0.4
        _append_signal(matched_signals, "politics", "政策文件问法")
    if "作者" in full_text or "文段" in full_text:
        scores["verbal"] += 1.0
        _append_signal(matched_signals, "verbal", "文段/作者意图")
    if logic_signal:
        scores["reasoning"] += 1.4 if strong_logic_signal else 0.7
        _append_signal(matched_signals, "reasoning", "逻辑判断结构")
    if logic_signal and statement_option_count >= 3:
        scores["reasoning"] += 0.6
        _append_signal(matched_signals, "reasoning", "推理型选项")
    if reasoning_support_signal and statement_option_count >= 3:
        scores["reasoning"] += 1.3
        scores["verbal"] -= 0.6
        _append_signal(matched_signals, "reasoning", "支持前提问法")
    if pure_numeric_option_count >= 1 and digit_count >= 6 and not strong_logic_signal:
        scores["quant"] += 0.8
        scores["reasoning"] -= 0.7
        _append_signal(matched_signals, "quant", "计算型题干")
    if quant_prompt_signal and (digit_count >= 4 or pure_numeric_option_count >= 1):
        scores["quant"] += 1.2
        if not strong_logic_signal:
            scores["reasoning"] -= 0.4
        _append_signal(matched_signals, "quant", "数量问法")
    if quant_graph_signal and (image_count > 0 or digit_count >= 4):
        scores["quant"] += 2.6
        scores["reasoning"] -= 0.8
        scores["common_sense"] -= 0.6
        _append_signal(matched_signals, "quant", "图表计算场景")
    if definition_signal and long_stem and summary_option_count >= 2 and scenario_option_count == 0:
        scores["verbal"] += 1.0
        scores["reasoning"] -= 0.6
        _append_signal(matched_signals, "verbal", "概念阐释型文段")
    if (
        definition_signal
        and long_stem
        and statement_option_count >= 3
        and scenario_option_count == 0
        and "根据上述材料" in full_text
    ):
        scores["verbal"] += 3.0
        scores["reasoning"] -= 2.2
        scores["common_sense"] -= 1.5
        _append_signal(matched_signals, "verbal", "概念阐释型文段")
    if (
        "是指" in stem_text
        and "根据上述材料" in full_text
        and not _has_any(full_text, ("符合定义", "不符合定义", "属于", "不属于"))
    ):
        scores["verbal"] += 1.0
        scores["reasoning"] -= 0.6
        _append_signal(matched_signals, "verbal", "概念阐释型文段")
    if relational_reasoning_signal and pure_numeric_option_count >= 3:
        scores["quant"] += 1.8
        scores["reasoning"] -= 0.9
        _append_signal(matched_signals, "quant", "循环位置计算")
    if cycle_quant_signal and pure_numeric_option_count >= 3:
        scores["quant"] += 2.0
        scores["reasoning"] -= 1.0
        _append_signal(matched_signals, "quant", "循环位置计算")
    if science_reasoning_signal and statement_option_count >= 3:
        scores["reasoning"] += 1.6
        scores["verbal"] -= 0.6
        _append_signal(matched_signals, "reasoning", "图示信息判断")
    if constraint_signal:
        scores["reasoning"] += 2.6
        scores["common_sense"] -= 1.0
        _append_signal(matched_signals, "reasoning", "约束条件推理")
    if set_relation_signal:
        scores["reasoning"] += 1.6
        _append_signal(matched_signals, "reasoning", "集合关系推理")

    for kind in scores:
        if kind == "data" and not allow_data:
            continue
        _apply_subject_structural_markers(
            kind=kind,
            scores=scores,
            matched_signals=matched_signals,
            digit_count=digit_count,
            numeric_option_count=pure_numeric_option_count,
            sequence_count=sequence_count,
            long_stem=long_stem,
            image_count=image_count,
            material_text=material_text,
            material_header=material_header,
            definition_signal=definition_signal,
            analogy_signal=analogy_signal,
            logic_signal=logic_signal,
            knowledge_prompt_signal=knowledge_prompt_signal,
            constraint_signal=constraint_signal,
            set_relation_signal=set_relation_signal,
            table_signal=table_signal,
            chart_signal=chart_signal,
            graphic_signal=graphic_signal,
            full_text=full_text,
        )

    best_subtypes: dict[SubjectKind, str] = {}
    for kind in scores:
        if kind == "data" and not allow_data:
            continue
        subtype_name, subtype_bonus = _score_subject_subtype(
            kind,
            full_text=full_text,
            stem_text=stem_text,
            option_texts=tuple(option_texts),
            material_text=material_text,
            image_count=image_count,
            matched_signals=matched_signals,
        )
        if subtype_name:
            scores[kind] += subtype_bonus
            best_subtypes[kind] = subtype_name

    _blend_learned_subject_scores(
        scores=scores,
        matched_signals=matched_signals,
        stem=stem,
        options=option_texts,
        material_text=material_text,
        image_count=image_count,
        material_header=material_header,
        allow_data=allow_data,
    )

    candidates = [
        (kind, score)
        for kind, score in scores.items()
        if kind != "data" or allow_data
    ]
    candidates.sort(key=lambda item: item[1], reverse=True)
    best_kind, best_score = candidates[0]
    second_score = candidates[1][1] if len(candidates) > 1 else 0.0
    margin = best_score - second_score
    confidence = _calibrate_subject_confidence(best_score, second_score, len(matched_signals.get(best_kind, [])))
    if best_score < 1.35:
        return SubjectInferenceDiagnostics(
            kind="unknown",
            margin=margin,
            confidence=confidence,
            best_score=best_score,
            second_score=second_score,
        )
    if best_score < 2.0 and margin < 0.55:
        return SubjectInferenceDiagnostics(
            kind="unknown",
            margin=margin,
            confidence=confidence,
            best_score=best_score,
            second_score=second_score,
        )
    if margin < 0.2:
        return SubjectInferenceDiagnostics(
            kind="unknown",
            margin=margin,
            confidence=confidence,
            best_score=best_score,
            second_score=second_score,
        )
    return SubjectInferenceDiagnostics(
        kind=best_kind,
        margin=margin,
        confidence=confidence,
        best_score=best_score,
        second_score=second_score,
        subtype=best_subtypes.get(best_kind),
        matched_signals=tuple(matched_signals.get(best_kind, ())[:10]),
    )


def infer_subject_from_content(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    image_count: int = 0,
    material_header: str = "",
    allow_data: bool = True,
) -> tuple[SubjectKind, float]:
    diagnostics = infer_subject_diagnostics(
        stem=stem,
        options=options,
        material_text=material_text,
        image_count=image_count,
        material_header=material_header,
        allow_data=allow_data,
    )
    return diagnostics.kind, diagnostics.margin


def infer_document_subject(
    texts: Iterable[str],
    *,
    image_count: int = 0,
    material_header_count: int = 0,
) -> tuple[SubjectKind | None, float]:
    lines = [text.strip() for text in texts if (text or "").strip()]
    if not lines and image_count <= 0:
        return None, 0.0
    material_header_like_count = sum(
        1
        for line in lines
        if re.match(r"^\s*[【\[]?\s*材料\s*[一二三四五六七八九十百千万\d〇零两]?\s*[】\]]?\s*$", _nfkc(line))
    )
    material_prompt_hits = sum(
        1
        for line in lines
        if any(
            marker in _nfkc(line)
            for marker in (
                "根据上述材料",
                "根据下列材料",
                "根据以下材料",
                "根据所给材料",
                "根据材料",
                "根据上述资料",
                "根据下列资料",
                "根据以下资料",
                "根据所给资料",
            )
        )
    )
    material_hits = material_header_like_count + material_prompt_hits
    objective_question_like_count = sum(
        1
        for line in lines[:120]
        if re.match(r"^\d{1,3}\s*[\.．、)\uFF09]", _nfkc(line))
    )
    objective_option_like_count = sum(
        1
        for line in lines[:160]
        if re.match(r"^[ABCD]\s*[\.．、)\uFF09:：]", _nfkc(line), re.IGNORECASE)
    )
    kind, confidence = infer_subject_from_content(
        stem=_clean_text(lines),
        options=(),
        material_text=_clean_text(lines[:8]) if material_hits or material_header_count else "",
        image_count=image_count,
        material_header="材料一" if material_header_count else "",
        allow_data=True,
    )
    sparse_material_headers_in_objective_book = (
        (material_header_count or material_header_like_count >= 1)
        and (material_header_count + material_header_like_count) <= 2
        and objective_question_like_count >= 6
        and objective_option_like_count >= 12
    )
    if kind == "unknown":
        if material_hits == 0 and material_header_count == 0 and objective_question_like_count >= 2 and objective_option_like_count >= 4:
            fallback_kind, fallback_confidence = infer_subject_from_content(
                stem=_clean_text(lines),
                options=(),
                image_count=min(image_count, 1),
                allow_data=False,
            )
            if fallback_kind != "unknown":
                return fallback_kind, max(fallback_confidence, confidence)
        return None, confidence
    if (
        (material_header_count or material_header_like_count >= 1)
        and not sparse_material_headers_in_objective_book
        and (kind == "data" or material_prompt_hits >= 2 or objective_question_like_count <= 3)
    ):
        return "data", max(confidence, 1.0)
    if material_prompt_hits >= 2 and objective_question_like_count <= 3:
        return "data", max(confidence, 1.0)
    if (
        kind == "data"
        and material_hits == 0
        and material_header_count == 0
        and objective_question_like_count >= 2
        and objective_option_like_count >= 4
    ):
        fallback_kind, fallback_confidence = infer_subject_from_content(
            stem=_clean_text(lines),
            options=(),
            image_count=min(image_count, 1),
            allow_data=False,
        )
        if fallback_kind != "unknown":
            return fallback_kind, max(fallback_confidence, min(confidence, 0.65))
    return kind, confidence
