# -*- coding: utf-8 -*-
"""检测并剥离"真题+解析"合订版 / 纯答案解析册中的解析条目。

公考真题的下载版常把【答案】【解析】与题目混在同一个 PDF 里。解析正文会被
切题逻辑误当成题目（题干变成"本题考查…"/"故正确答案为X"，或 0 选项的
".【答案】X【解析】…"），把一份 130 题的卷子炸成 250+ 条垃圾。

本模块在 PDF 工程构建完成后做一次后处理：

- ``clean``：没有解析条目，原样返回（金标准/干净卷必然走这条，零误伤）。
- ``combined``（合订版）：真题在、解析也在 → 删掉解析条目，保留真题。
- ``answer_booklet``（纯答案册）：通篇是解析、几乎没有真题 → 只保留少数可信
  真题，并置 ``is_answer_booklet``，供上层提示用户"这是答案册，请改用试卷文件"。

判定全部基于"按题"信号，且这些信号在真实题干中几乎不出现，因此对干净卷
近乎零误伤（已在 7 份金标准 + 多份外部真卷上验证：解析条目计数恒为 0）。
"""
from __future__ import annotations

import logging
import re
import statistics

from domain.models import ExamProject, QuestionNode

logger = logging.getLogger(__name__)

# 题干层面的强解析信号（真实题干里几乎不出现）。
_EXPLANATION_STEM = re.compile(
    r"【答案】|【解析】"
    r"|本题(主要)?考查"
    r"|故正确答案为\s*[A-D]"
    r"|因此[，,]?\s*选择\s*[A-D]\s*选项"
    r"|^[.\s]*【答案】"
    r"|^\s*\d{1,3}\s*[.、．]\s*解析\b"
    r"|^\s*第[一二三四五六七八九十]+步[，,]"
)
# 解析"选项"是分析长句（"A项错误，…指出…"），而非简短答案。
_OPTION_ANALYSIS = re.compile(
    r"项\s*(正确|错误)|指出[，,:：]|规定[，,:：]|表述\s*(正确|错误)|【解析】"
)

_REAL_STEM_MIN_LEN = 8
_REAL_OPTION_MEDIAN_MAX = 50
_BOOKLET_RATIO = 0.30
_COMBINED_MIN_EXPLANATIONS = 3


def is_explanation_question(question: QuestionNode) -> bool:
    """题干/选项呈现"答案解析"特征，而非真实题目。"""
    stem = (question.stem or "").strip()
    if _EXPLANATION_STEM.search(stem):
        return True
    options = question.options or []
    if options:
        analysisish = sum(
            1 for opt in options if _OPTION_ANALYSIS.search((opt.text or "").strip())
        )
        if analysisish >= max(2, (len(options) + 1) // 2):
            return True
    return False


def is_real_question(question: QuestionNode) -> bool:
    """像一道正常题目：有题干、四个简短选项、且不是解析条目。"""
    if is_explanation_question(question):
        return False
    stem = (question.stem or "").strip()
    options = question.options or []
    if len(stem) < _REAL_STEM_MIN_LEN or len(options) != 4:
        return False
    texts = [(opt.text or "").strip() for opt in options]
    if any(not text for text in texts):
        return False
    if sum(1 for text in texts if _OPTION_ANALYSIS.search(text)) >= 2:
        return False
    if statistics.median(len(text) for text in texts) > _REAL_OPTION_MEDIAN_MAX:
        return False
    return True


def classify_explanation_content(project: ExamProject) -> dict:
    """把整份工程归类为 clean / combined / answer_booklet。"""
    rows = [question for _section, _material, question in project.iter_questions()]
    total = len(rows)
    explanation = sum(1 for question in rows if is_explanation_question(question))
    real = sum(1 for question in rows if is_real_question(question))
    ratio = (explanation / total) if total else 0.0
    if total and ratio >= _BOOKLET_RATIO and real < max(8, 0.15 * total):
        category = "answer_booklet"
    elif explanation >= _COMBINED_MIN_EXPLANATIONS:
        category = "combined"
    else:
        category = "clean"
    return {
        "category": category,
        "total_questions": total,
        "explanation_questions": explanation,
        "real_questions": real,
        "explanation_ratio": round(ratio, 3),
    }


def filter_explanation_questions(project: ExamProject) -> dict:
    """剥离解析条目；返回分类信息（含被删条数与 is_answer_booklet）。

    干净卷（无解析条目）原样返回，不做任何改动。
    """
    info = classify_explanation_content(project)
    category = info["category"]
    info["removed_questions"] = 0
    info["is_answer_booklet"] = category == "answer_booklet"
    if category == "clean":
        return info

    if category == "combined":
        keep = lambda question: not is_explanation_question(question)
    else:  # answer_booklet：整册都是解析，只保留少数可信真题
        keep = is_real_question

    removed = 0
    for section in project.sections:
        before = len(section.questions)
        section.questions = [q for q in section.questions if keep(q)]
        removed += before - len(section.questions)
        for material in section.material_sets:
            mbefore = len(material.questions)
            material.questions = [q for q in material.questions if keep(q)]
            removed += mbefore - len(material.questions)
        section.material_sets = [m for m in section.material_sets if m.questions]
    project.sections = [s for s in project.sections if s.questions or s.material_sets]

    info["removed_questions"] = removed
    if category == "answer_booklet":
        logger.warning(
            "检测到疑似答案/解析册（共 %d 条，仅 %d 条像真题）：已剥离解析条目，"
            "建议改用对应的试卷/题目 PDF。",
            info["total_questions"],
            info["real_questions"],
        )
        project.import_notices.append(
            f"⚠ 这份 PDF 更像『答案/解析册』（共 {info['total_questions']} 条，"
            f"仅 {info['real_questions']} 条像题目）。建议改用对应的『试卷/题目』PDF。"
        )
    else:
        logger.info(
            "剥离解析合订版条目 %d 条，保留约 %d 题。",
            removed,
            info["total_questions"] - removed,
        )
        if removed:
            project.import_notices.append(
                f"已识别为『真题+解析』合订版：自动剥离解析条目 {removed} 条，"
                f"保留约 {info['total_questions'] - removed} 题。"
            )
    return info
