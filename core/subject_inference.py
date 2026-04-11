from __future__ import annotations

import re
from collections.abc import Iterable

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


def default_subject_title(kind: SubjectKind) -> str:
    return _DISPLAY_NAMES.get(kind, "待确认科目")


def _clean_text(parts: Iterable[str]) -> str:
    joined = "\n".join(part.strip() for part in parts if (part or "").strip())
    return re.sub(r"\s+", " ", joined).strip()


def _is_numericish(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    return bool(_NUMERICISH_OPTION.match(value))


def infer_subject_from_content(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    image_count: int = 0,
    material_header: str = "",
    allow_data: bool = True,
) -> tuple[SubjectKind, float]:
    stem_text = _clean_text([material_header, material_text, stem])
    option_texts = [text.strip() for text in (options or []) if (text or "").strip()]
    option_blob = _clean_text(option_texts)
    full_text = _clean_text([stem_text, option_blob])
    if not full_text and image_count <= 0:
        return "unknown", 0.0

    scores: dict[SubjectKind, float] = {
        "politics": 0.0,
        "common_sense": 0.0,
        "verbal": 0.0,
        "quant": 0.0,
        "reasoning": 0.0,
        "data": 0.0,
    }

    lowered = full_text.lower()
    for kind, keywords in _KEYWORDS.items():
        if kind == "data" and not allow_data:
            continue
        for keyword in keywords:
            if keyword and keyword.lower() in lowered:
                scores[kind] += 1.25 if len(keyword) >= 4 else 0.8

    digit_count = len(_DIGIT_RE.findall(full_text))
    percent_hits = len(_PERCENT_RE.findall(full_text))
    numeric_option_count = sum(1 for text in option_texts if _is_numericish(text))
    long_stem = len(stem_text) >= 42

    if numeric_option_count >= 3:
        scores["quant"] += 2.0
        if allow_data:
            scores["data"] += 1.0
    if digit_count >= 10:
        scores["quant"] += 1.2
    if digit_count >= 18:
        scores["data"] += 1.0 if allow_data else 0.0
    if percent_hits >= 1 and allow_data:
        scores["data"] += 1.5
    if material_header:
        scores["data"] += 4.0 if allow_data else 0.0
    if material_text:
        scores["data"] += 2.4 if allow_data else 0.0
    if image_count and allow_data and material_text:
        scores["data"] += 1.0
    if image_count and not material_text and digit_count <= 4:
        scores["reasoning"] += 1.2
    if long_stem and digit_count <= 8:
        scores["verbal"] += 0.9
    if len(option_texts) >= 4 and all(len(text) <= 8 for text in option_texts) and digit_count <= 3:
        scores["reasoning"] += 0.6
        scores["common_sense"] += 0.3
    if "下列" in full_text and "说法" in full_text:
        scores["common_sense"] += 1.0
    if "作者" in full_text or "文段" in full_text:
        scores["verbal"] += 1.0
    if "如果" in full_text and ("那么" in full_text or "则" in full_text):
        scores["reasoning"] += 1.0

    candidates = [
        (kind, score)
        for kind, score in scores.items()
        if kind != "data" or allow_data
    ]
    candidates.sort(key=lambda item: item[1], reverse=True)
    best_kind, best_score = candidates[0]
    second_score = candidates[1][1] if len(candidates) > 1 else 0.0
    confidence = best_score - second_score
    if best_score < 1.35:
        return "unknown", confidence
    if best_score < 2.0 and confidence < 0.55:
        return "unknown", confidence
    if confidence < 0.2:
        return "unknown", confidence
    return best_kind, confidence


def infer_document_subject(
    texts: Iterable[str],
    *,
    image_count: int = 0,
    material_header_count: int = 0,
) -> tuple[SubjectKind | None, float]:
    lines = [text.strip() for text in texts if (text or "").strip()]
    if not lines and image_count <= 0:
        return None, 0.0
    material_hits = sum(1 for line in lines if "材料" in line)
    kind, confidence = infer_subject_from_content(
        stem=_clean_text(lines),
        options=(),
        material_text=_clean_text(lines[:8]) if material_hits or material_header_count else "",
        image_count=image_count,
        material_header="材料一" if material_header_count else "",
        allow_data=True,
    )
    if kind == "unknown":
        return None, confidence
    if material_header_count or material_hits >= 2:
        return "data", max(confidence, 1.0)
    return kind, confidence
