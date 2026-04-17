from __future__ import annotations

import argparse
import json
import re
import sys
from dataclasses import asdict, dataclass
from pathlib import Path
from typing import Any

try:
    import fitz
except Exception:  # pragma: no cover - optional dependency
    fitz = None

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.project_quality import annotate_project_quality
from ingest.pdf.project_builder import build_exam_project_from_pdf


QUESTION_COUNT_RE = re.compile(r"（(\d+)题）|\((\d+)题\)|(?<!答案及解析)(\d+)题")


@dataclass
class PdfAuditRow:
    pdf: str
    path: str
    subjects: list[str]
    expected: int | None
    parsed_total: int
    numeric_total: int
    unique_numeric: int
    missing_count: int
    missing_numbers: list[int]
    parser_missing_count: int
    parser_missing_numbers: list[int]
    source_gap_count: int
    source_gap_numbers: list[int]
    duplicate_numbers: list[int]
    out_of_range_numbers: list[int]
    nonnumeric_count: int
    nonnumeric_samples: list[dict[str, Any]]
    flagged_questions: int
    severe_questions: int
    source_defect_questions: int
    top_issue_codes: list[dict[str, Any]]
    severe_samples: list[dict[str, Any]]


def expected_count(path: Path) -> int | None:
    match = QUESTION_COUNT_RE.search(path.name)
    if match:
        for group in match.groups():
            if group:
                return int(group)
    if "模拟卷" in path.name:
        return 120
    if "行政执法卷" in path.name:
        return 130
    return None


def discover_pdfs(root: Path) -> list[Path]:
    candidates = []
    for path in sorted(root.rglob("*.pdf")):
        if "参考答案及解析" in path.name:
            continue
        candidates.append(path)
    return candidates


def _question_number_token(number: int) -> re.Pattern[str]:
    return re.compile(rf"(?<!\d){number}\s*[、.．](?!\d)")


def _page_text_map(pdf_path: Path, page_numbers: set[int]) -> dict[int, str]:
    if fitz is None or not page_numbers or not pdf_path.exists():
        return {}
    try:
        document = fitz.open(pdf_path)
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


def _detect_source_gap_numbers(project: Any, path: Path, missing_numbers: list[int]) -> list[int]:
    if not missing_numbers:
        return []

    numbered_questions = sorted(
        (
            question
            for _section, _material, question in project.iter_questions()
            if question.numeric_source_number is not None
        ),
        key=lambda question: question.numeric_source_number or 0,
    )
    if len(numbered_questions) < 2:
        return []

    page_numbers: set[int] = set()
    for question in numbered_questions:
        for page_number in getattr(question, "page_numbers", []) or []:
            if isinstance(page_number, int) and page_number > 0:
                page_numbers.add(page_number)
    page_text_map = _page_text_map(path, page_numbers)

    source_gap_numbers: list[int] = []
    for missing in missing_numbers:
        previous = next(
            (question for question in reversed(numbered_questions) if (question.numeric_source_number or 0) < missing),
            None,
        )
        nxt = next(
            (question for question in numbered_questions if (question.numeric_source_number or 0) > missing),
            None,
        )
        if previous is None or nxt is None:
            continue
        previous_number = previous.numeric_source_number
        next_number = nxt.numeric_source_number
        if previous_number is None or next_number is None:
            continue
        if previous_number + 1 != missing or next_number - 1 != missing:
            continue

        candidate_pages = {
            page_number
            for question in (previous, nxt)
            for page_number in (getattr(question, "page_numbers", []) or [])
            if isinstance(page_number, int) and page_number > 0
        }
        if not candidate_pages:
            continue

        previous_pattern = _question_number_token(previous_number)
        missing_pattern = _question_number_token(missing)
        next_pattern = _question_number_token(next_number)
        combined_text = "\n".join(page_text_map.get(page_number, "") for page_number in sorted(candidate_pages))
        if not combined_text:
            continue
        if previous_pattern.search(combined_text) and next_pattern.search(combined_text) and not missing_pattern.search(
            combined_text
        ):
            source_gap_numbers.append(missing)

    return source_gap_numbers


def _sequential_section_questions(project: Any) -> list[list[Any]]:
    sequences: list[list[Any]] = []
    for section in project.sections:
        if section.kind == "data":
            questions = [question for material in section.material_sets for question in material.questions]
        else:
            questions = list(section.questions)
        if questions:
            sequences.append(questions)
    return sequences


def _effective_question_pages(sequential_questions: list[Any], index: int) -> tuple[list[int], bool]:
    question = sequential_questions[index]
    direct_pages = [
        page_number
        for page_number in (getattr(question, "page_numbers", []) or [])
        if isinstance(page_number, int) and page_number > 0
    ]
    if direct_pages:
        return sorted(set(direct_pages)), False

    previous_pages: list[int] = []
    next_pages: list[int] = []
    for candidate in reversed(sequential_questions[:index]):
        previous_pages = [
            page_number
            for page_number in (getattr(candidate, "page_numbers", []) or [])
            if isinstance(page_number, int) and page_number > 0
        ]
        if previous_pages:
            break
    for candidate in sequential_questions[index + 1 :]:
        next_pages = [
            page_number
            for page_number in (getattr(candidate, "page_numbers", []) or [])
            if isinstance(page_number, int) and page_number > 0
        ]
        if next_pages:
            break

    if previous_pages and next_pages:
        overlap = sorted(set(previous_pages) & set(next_pages))
        if overlap:
            return overlap, True
        closest_gap, closest_prev, closest_next = min(
            (abs(prev_page - next_page), prev_page, next_page)
            for prev_page in previous_pages
            for next_page in next_pages
        )
        if closest_gap <= 2:
            return sorted({closest_prev, closest_next}), True
        merged = sorted(set(previous_pages + next_pages))
        if merged[-1] - merged[0] <= 2:
            return merged, True
        return sorted({previous_pages[-1], next_pages[0]}), True
    if previous_pages:
        return [previous_pages[-1]], True
    if next_pages:
        return [next_pages[0]], True
    return [], False


def _classify_severe_sample(issue_codes: list[str]) -> str:
    for code in issue_codes:
        if code.startswith("source_"):
            return code
    return issue_codes[0] if issue_codes else "unknown"


def audit_pdf(path: Path) -> PdfAuditRow:
    project = build_exam_project_from_pdf(path, mode="all")
    quality = annotate_project_quality(project)
    numeric_numbers: list[int] = []
    nonnumeric_samples: list[dict[str, Any]] = []
    subjects: set[str] = set()
    issue_counter: dict[str, int] = {}
    severe_samples: list[dict[str, Any]] = []
    effective_pages_by_id: dict[int, tuple[list[int], bool]] = {}

    for sequential_questions in _sequential_section_questions(project):
        for index, question in enumerate(sequential_questions):
            effective_pages_by_id[id(question)] = _effective_question_pages(sequential_questions, index)

    for section, material, question in project.iter_questions():
        subjects.add(section.kind)
        source_number = (question.source_number or "").strip()
        if source_number.isdigit():
            numeric_numbers.append(int(source_number))
        elif len(nonnumeric_samples) < 5:
            nonnumeric_samples.append(
                {
                    "source_number": source_number,
                    "material": material.header if material else "",
                    "stem": (question.stem or "")[:160],
                }
            )
        for issue in question.review_issues:
            issue_counter[issue.code] = issue_counter.get(issue.code, 0) + 1
        if any(issue.severity == "error" for issue in question.review_issues) and len(severe_samples) < 5:
            error_codes = [issue.code for issue in question.review_issues if issue.severity == "error"]
            page_numbers, page_numbers_inferred = effective_pages_by_id.get(id(question), ([], False))
            severe_samples.append(
                {
                    "source_number": source_number,
                    "section": section.kind,
                    "material": material.header if material else "",
                    "issue_codes": error_codes,
                    "defect_type": _classify_severe_sample(error_codes),
                    "page_numbers": page_numbers,
                    "page_numbers_inferred": page_numbers_inferred,
                    "stem": (question.stem or "")[:160],
                }
            )

    expected = expected_count(path)
    unique_numbers = sorted(set(numeric_numbers))
    missing_numbers = (
        [value for value in range(1, expected + 1) if value not in unique_numbers]
        if expected is not None
        else []
    )
    source_gap_numbers = _detect_source_gap_numbers(project, path, missing_numbers)
    parser_missing_numbers = [value for value in missing_numbers if value not in set(source_gap_numbers)]
    duplicate_numbers = sorted({value for value in numeric_numbers if numeric_numbers.count(value) > 1})
    out_of_range_numbers = (
        [value for value in unique_numbers if value < 1 or value > expected]
        if expected is not None
        else []
    )
    return PdfAuditRow(
        pdf=path.name,
        path=str(path),
        subjects=sorted(subjects),
        expected=expected,
        parsed_total=project.question_count,
        numeric_total=len(numeric_numbers),
        unique_numeric=len(unique_numbers),
        missing_count=len(missing_numbers),
        missing_numbers=missing_numbers,
        parser_missing_count=len(parser_missing_numbers),
        parser_missing_numbers=parser_missing_numbers,
        source_gap_count=len(source_gap_numbers),
        source_gap_numbers=source_gap_numbers,
        duplicate_numbers=duplicate_numbers,
        out_of_range_numbers=out_of_range_numbers,
        nonnumeric_count=project.question_count - len(numeric_numbers),
        nonnumeric_samples=nonnumeric_samples,
        flagged_questions=quality.flagged_questions,
        severe_questions=quality.severe_questions,
        source_defect_questions=quality.source_defect_questions,
        top_issue_codes=[
            {"code": code, "count": count}
            for code, count in sorted(issue_counter.items(), key=lambda item: (-item[1], item[0]))[:10]
        ],
        severe_samples=severe_samples,
    )


def main() -> None:
    parser = argparse.ArgumentParser(description="Audit parse coverage for local PDF corpus.")
    parser.add_argument(
        "--root",
        type=Path,
        default=ROOT,
        help="Workspace root to scan for PDFs.",
    )
    parser.add_argument(
        "--json-output",
        type=Path,
        default=None,
        help="Optional JSON output path.",
    )
    args = parser.parse_args()

    rows = [audit_pdf(path) for path in discover_pdfs(args.root)]
    payload = [asdict(row) for row in rows]
    text = json.dumps(payload, ensure_ascii=False, indent=2)
    print(text)

    if args.json_output is not None:
        args.json_output.parent.mkdir(parents=True, exist_ok=True)
        args.json_output.write_text(text, encoding="utf-8")


if __name__ == "__main__":
    main()
