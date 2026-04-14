from __future__ import annotations

import json
import re
import shutil
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
from typing import Iterable


ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.pdf_exam_extract import extract_pdf_line_items_with_metadata
from core.pdf_exam_parse import parse_line_items
from core.subject_inference import infer_subject_diagnostics
from domain.models import ExamProject, MaterialSet, QuestionNode
from ingest.pdf.layout import extract_pdf_text_lines
from ingest.pdf.project_builder import build_project_from_parsed_exam


DATASET_DIR = ROOT / "【02】最新行测5000题题库"
OUTPUT_JSON = ROOT / "data" / "gongkao_corpus_profiles.json"
OUTPUT_MD = ROOT / "docs" / "GONGKAO_CORPUS.md"

DATASETS = (
    {
        "title": "言语理解 2000 题",
        "file": "行测——言语理解2000题.pdf",
        "kind": "verbal",
        "expected_count": 2000,
    },
    {
        "title": "判断推理 2000 题",
        "file": "行测——判断推理（2000题）.pdf",
        "kind": "reasoning",
        "expected_count": 2000,
    },
    {
        "title": "数量关系 850 题",
        "file": "行测——数量关系（850题）.pdf",
        "kind": "quant",
        "expected_count": 850,
    },
    {
        "title": "资料分析 950 题",
        "file": "行测——资料分析（950题）.pdf",
        "kind": "data",
        "expected_count": 950,
    },
)

_COMBO_OPTION_RE = re.compile(r"^\d+项$")
_GROUPING_OPTION_RE = re.compile(r"^(?:[1-6①②③④⑤⑥]+[，,][1-6①②③④⑤⑥]+)$")
_ORDERING_OPTION_RE = re.compile(r"^(?:[1-6]{4,6}|[①②③④⑤⑥]{4,6})$")
_SHORT_COLON_ANALOGY_RE = re.compile(r"^[^:：]{1,8}(?:[:：][^:：]{1,8}){1,2}$")
_CONSTRAINT_MARKER_RE = re.compile(r"(?:\(\d+\)|（\d+）|[①②③④⑤⑥⑦⑧⑨⑩])")
_CASE_OPTION_PREFIX_RE = re.compile(r"^[甲乙丙丁戊己庚辛壬癸](?:某|公司|机关|单位|市民|村民|老师|同学|部门)?")
_BLANK_SLOT_RE = re.compile(r"(?:\(\s*\)|（\s*）|____+|——+|—\s*—)")

_SUMMARY_OPTION_MARKERS = (
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
_POLITICS_ANCHORS = (
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
_LEGAL_SCENARIO_MARKERS = (
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
_ASSIGNMENT_REASONING_MARKERS = (
    "分别",
    "均不相同",
    "各不相同",
    "已知",
    "每人",
    "每个",
    "甲、乙、丙",
    "甲乙丙",
)
_SET_RELATION_MARKERS = (
    "用一个圆来表示",
    "集合",
    "关系符合下图",
    "之间的关系",
)
_CULTURAL_KNOWLEDGE_MARKERS = (
    "诗词",
    "作者",
    "朝代",
    "典故",
    "地貌",
    "颜色",
    "节气",
    "成语",
)
_SENTENCE_END_MARKERS = ("。", "；", "？", "!", "，")


@dataclass
class QuestionRow:
    question: QuestionNode
    section_kind: str
    material: MaterialSet | None = None


def _iter_questions(project: ExamProject) -> Iterable[QuestionRow]:
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                for question in material.questions:
                    yield QuestionRow(question=question, section_kind=section.kind, material=material)
        else:
            for question in section.questions:
                yield QuestionRow(question=question, section_kind=section.kind, material=None)


def _build_project_for_study(pdf_path: Path, document_subject_hint: str) -> ExamProject:
    items, temp_dir, image_regions = extract_pdf_line_items_with_metadata(str(pdf_path))
    try:
        exam = parse_line_items(
            items,
            mode="all",
            document_subject_hint=document_subject_hint,
            source_name=pdf_path.name,
        )
        layout_lines = extract_pdf_text_lines(str(pdf_path))
        return build_project_from_parsed_exam(
            exam,
            source_pdf_path=str(pdf_path),
            layout_lines=layout_lines,
            image_regions=image_regions,
            title=pdf_path.stem,
        )
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)


def _looks_like_statement_option(text: str) -> bool:
    value = (text or "").strip()
    if not value:
        return False
    if len(value) >= 12:
        return True
    return any(marker in value for marker in _SENTENCE_END_MARKERS) or any(
        marker in value for marker in ("我国", "可以", "应当", "不得", "属于", "不属于", "错误", "正确", "符合")
    )


def _scenario_option_count(option_texts: list[str]) -> int:
    count = 0
    for text in option_texts:
        value = (text or "").strip()
        if len(value) < 8:
            continue
        if _CASE_OPTION_PREFIX_RE.search(value) or any(
            marker in value for marker in ("某", "张三", "李四", "王某", "甲公司", "乙公司", "行政机关", "当事人", "行为", "措施")
        ):
            count += 1
    return count


def _summary_option_count(option_texts: list[str]) -> int:
    return sum(1 for text in option_texts if len((text or "").strip()) >= 8 and any(m in (text or "") for m in _SUMMARY_OPTION_MARKERS))


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


def _build_dataset_profile(dataset: dict[str, object]) -> dict[str, object]:
    file_name = str(dataset["file"])
    expected_kind = str(dataset["kind"])
    pdf_path = DATASET_DIR / file_name
    project = _build_project_for_study(pdf_path, expected_kind)

    questions = list(_iter_questions(project))
    signal_counts: Counter[str] = Counter()
    subtype_counts: Counter[str] = Counter()
    feature_counts: Counter[str] = Counter()
    examples: dict[str, list[dict[str, object]]] = defaultdict(list)
    disagreement_examples: list[dict[str, object]] = []

    for row in questions:
        question = row.question
        option_texts = [(opt.text or "").strip() for opt in question.options if (opt.text or "").strip()]
        stem = (question.stem or "").strip()
        image_count = len(question.stem_assets) + sum(1 for opt in question.options if opt.image_path)
        if row.material is not None:
            image_count += len(row.material.body_assets)
        diagnostics = infer_subject_diagnostics(
            stem=stem,
            options=option_texts,
            material_header=(row.material.header if row.material else ""),
            material_text=(row.material.body if row.material else ""),
            image_count=image_count,
            allow_data=(expected_kind == "data"),
        )

        subtype_counts[question.inferred_subtype or diagnostics.subtype or "None"] += 1
        signal_counts.update(diagnostics.matched_signals)

        combo = any(_looks_like_combo_option(text) for text in option_texts)
        grouping = any(_GROUPING_OPTION_RE.fullmatch(text) for text in option_texts)
        ordering = any(_ORDERING_OPTION_RE.fullmatch(text) for text in option_texts)
        short_words = bool(option_texts) and len(option_texts) >= 4 and all(1 <= len(text) <= 8 for text in option_texts)
        statements = sum(1 for text in option_texts if _looks_like_statement_option(text))
        summary = _summary_option_count(option_texts) >= 2
        law_scenario = _scenario_option_count(option_texts) >= 2 and any(marker in stem for marker in _LEGAL_SCENARIO_MARKERS)
        politics_anchor = sum(1 for marker in _POLITICS_ANCHORS if marker in stem) >= 1
        constraint = len(_CONSTRAINT_MARKER_RE.findall(stem)) >= 2 and any(marker in stem for marker in _ASSIGNMENT_REASONING_MARKERS)
        set_relation = any(marker in stem for marker in _SET_RELATION_MARKERS)
        analogy = _SHORT_COLON_ANALOGY_RE.fullmatch(stem) is not None or "相当于" in stem
        culture = any(marker in stem for marker in _CULTURAL_KNOWLEDGE_MARKERS)
        blank_slot = bool(_BLANK_SLOT_RE.search(stem))

        if combo:
            feature_counts["combo_option_questions"] += 1
        if grouping:
            feature_counts["grouping_option_questions"] += 1
        if ordering:
            feature_counts["ordering_option_questions"] += 1
        if short_words:
            feature_counts["short_word_option_questions"] += 1
        if statements >= 3:
            feature_counts["statement_option_questions"] += 1
        if summary:
            feature_counts["summary_option_questions"] += 1
        if blank_slot:
            feature_counts["blank_slot_questions"] += 1
        if law_scenario:
            feature_counts["law_scenario_questions"] += 1
        if politics_anchor:
            feature_counts["politics_anchor_questions"] += 1
        if constraint:
            feature_counts["constraint_reasoning_questions"] += 1
        if set_relation:
            feature_counts["set_relation_questions"] += 1
        if analogy:
            feature_counts["analogy_prompt_questions"] += 1
        if culture:
            feature_counts["culture_knowledge_questions"] += 1
        if image_count:
            feature_counts["image_questions"] += 1
        if any(not (opt.text or "").strip() and not opt.image_path for opt in question.options):
            feature_counts["blank_option_questions"] += 1

        for key, matched in (
            ("combo_option_questions", combo),
            ("grouping_option_questions", grouping),
            ("ordering_option_questions", ordering),
            ("blank_slot_questions", blank_slot),
            ("law_scenario_questions", law_scenario),
            ("constraint_reasoning_questions", constraint),
            ("set_relation_questions", set_relation),
            ("analogy_prompt_questions", analogy),
            ("summary_option_questions", summary),
        ):
            if matched and len(examples[key]) < 2:
                examples[key].append(
                    {
                        "number": question.source_number,
                        "stem": stem[:120],
                        "options": option_texts[:4],
                    }
                )

        if diagnostics.kind != expected_kind and len(disagreement_examples) < 12:
            disagreement_examples.append(
                {
                    "number": question.source_number,
                    "expected": expected_kind,
                    "predicted": diagnostics.kind,
                    "confidence": round(diagnostics.confidence, 4),
                    "signals": list(diagnostics.matched_signals[:8]),
                    "stem": stem[:160],
                }
            )

    question_count = len(questions)
    avg_stem_len = round(sum(len((row.question.stem or "").strip()) for row in questions) / max(1, question_count), 2)
    avg_option_chars = round(
        sum(sum(len((opt.text or "").strip()) for opt in row.question.options) for row in questions) / max(1, question_count),
        2,
    )

    return {
        "title": dataset["title"],
        "file": file_name,
        "kind": expected_kind,
        "expected_count": int(dataset["expected_count"]),
        "parsed_question_count": question_count,
        "avg_stem_length": avg_stem_len,
        "avg_option_chars": avg_option_chars,
        "subtype_distribution": dict(subtype_counts.most_common()),
        "feature_counts": dict(feature_counts.most_common()),
        "top_signals": dict(signal_counts.most_common(20)),
        "examples": dict(examples),
        "disagreement_examples": disagreement_examples,
    }


def _render_markdown(profiles: list[dict[str, object]]) -> str:
    lines = [
        "# 公考题库学习画像",
        "",
        "这份文档由 `scripts/study_gongkao_corpus.py` 自动生成。",
        "",
        "它的目标不是保存题库全文，而是把本地题库沉淀成可复用的结构经验，供分类器、质检器和本地 AI 修复器继续学习。",
        "",
        f"生成时间：{datetime.now().isoformat(timespec='seconds')}",
        "",
    ]
    for profile in profiles:
        lines.append(f"## {profile['title']}")
        lines.append("")
        lines.append(f"- 文件：`{profile['file']}`")
        lines.append(f"- 目标科目：`{profile['kind']}`")
        lines.append(f"- 预期题量：`{profile['expected_count']}`")
        lines.append(f"- 当前解析题量：`{profile['parsed_question_count']}`")
        lines.append(f"- 平均题干长度：`{profile['avg_stem_length']}`")
        lines.append(f"- 平均选项总字数：`{profile['avg_option_chars']}`")
        lines.append("")
        lines.append("### 题型分布")
        lines.append("")
        for name, count in dict(profile["subtype_distribution"]).items():
            lines.append(f"- `{name}`：{count}")
        lines.append("")
        lines.append("### 结构特征")
        lines.append("")
        for name, count in dict(profile["feature_counts"]).items():
            lines.append(f"- `{name}`：{count}")
        lines.append("")
        lines.append("### 高频信号")
        lines.append("")
        for name, count in dict(profile["top_signals"]).items():
            lines.append(f"- `{name}`：{count}")
        disagreement_examples = list(profile.get("disagreement_examples", []))
        if disagreement_examples:
            lines.append("")
            lines.append("### 仍值得继续盯的边界题")
            lines.append("")
            for item in disagreement_examples[:8]:
                lines.append(
                    f"- `Q{item['number']}` 预期 `{item['expected']}`，当前单题推断更像 `{item['predicted']}`，"
                    f"置信度 `{item['confidence']}`，信号：`{' / '.join(item['signals'])}`"
                )
                lines.append(f"  题干：{item['stem']}")
        lines.append("")
    return "\n".join(lines).rstrip() + "\n"


def main() -> None:
    profiles = [_build_dataset_profile(dataset) for dataset in DATASETS]
    payload = {
        "generated_at": datetime.now().isoformat(timespec="seconds"),
        "dataset_root": str(DATASET_DIR),
        "datasets": profiles,
    }

    OUTPUT_JSON.parent.mkdir(parents=True, exist_ok=True)
    OUTPUT_MD.parent.mkdir(parents=True, exist_ok=True)
    with OUTPUT_JSON.open("w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    OUTPUT_MD.write_text(_render_markdown(profiles), encoding="utf-8")
    print(f"Wrote {OUTPUT_JSON}")
    print(f"Wrote {OUTPUT_MD}")


if __name__ == "__main__":
    main()
