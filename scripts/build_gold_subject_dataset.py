from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
import sys
from typing import Any, Iterable

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.pdf_exam_extract import extract_pdf_line_items
from core.pdf_exam_parse import _preprocess_line_items, parse_quant_block
from core.learned_subject_model import build_subject_feature_record
from ingest.pdf.project_builder import build_exam_project_from_pdf


DEFAULT_CATALOG = ROOT / "data" / "gold_pdf_catalog.json"
DEFAULT_OUTPUT = ROOT / "data" / "datasets" / "subject_gold.jsonl"


def _clean_text(parts: Iterable[str]) -> str:
    return "\n".join(part.strip() for part in parts if (part or "").strip()).strip()


def _material_body_text(material) -> str:
    return _clean_text(material.body_lines)


def _rich_line_text(line) -> str:
    return "".join(text for text, _img in line.parts).strip()


def _fallback_rows_from_pdf(
    pdf_path: Path,
    *,
    source_pdf: str,
    source_form: str,
    forced_subject: str,
):
    raw, _ = extract_pdf_line_items(str(pdf_path))
    processed = _preprocess_line_items(raw)
    questions = parse_quant_block(processed, 0, len(processed))
    for question in questions:
        option_texts = [_rich_line_text(line) for line in question.option_lines if _rich_line_text(line)]
        stem = _clean_text(_rich_line_text(line) for line in question.stem_lines)
        image_count = sum(1 for line in question.stem_lines for _text, image_path in line.parts if image_path)
        image_count += sum(1 for line in question.option_lines for _text, image_path in line.parts if image_path)
        yield {
            "source_pdf": source_pdf,
            "source_form": source_form,
            "subject": forced_subject,
            "label_source": "catalog_subject",
            "question_no": question.source_number,
            "stem": stem,
            "options": option_texts,
            "material_header": "",
            "material_text": "",
            "image_count": image_count,
            "page_numbers": [],
            "feature_record": build_subject_feature_record(
                stem=stem,
                options=option_texts,
                image_count=image_count,
            ),
        }


def _expand_range_specs(range_specs: Iterable[str]) -> dict[str, str]:
    expanded: dict[str, str] = {}
    for item in range_specs:
        spec = str(item or "").strip()
        if not spec:
            continue
        if "-" in spec:
            start_text, end_text = spec.split("-", 1)
            try:
                start = int(start_text.strip())
                end = int(end_text.strip())
            except ValueError:
                continue
            if end < start:
                start, end = end, start
            for number in range(start, end + 1):
                expanded[str(number)] = ""
        else:
            try:
                expanded[str(int(spec))] = ""
            except ValueError:
                continue
    return expanded


def _catalog_subject_map(entry: dict[str, Any]) -> dict[str, str]:
    subject_by_number: dict[str, str] = {}
    sections = entry.get("sections") or {}
    if not isinstance(sections, dict):
        return {}
    for subject, range_specs in sections.items():
        expanded = _expand_range_specs(range_specs or [])
        for number in expanded:
            subject_by_number[number] = str(subject)
    return subject_by_number


def _catalog_subject_hint(entry: dict[str, Any]) -> str | None:
    if str(entry.get("form", "")).strip() != "single_subject_book":
        return None
    subject = str(entry.get("subject", "")).strip()
    return subject or None


def _iter_project_rows(
    project,
    source_pdf: str,
    source_form: str,
    forced_subject: str | None,
    explicit_subjects: dict[str, str] | None = None,
):
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                material_text = _material_body_text(material)
                material_images = len(material.body_assets)
                for question in material.questions:
                    if explicit_subjects:
                        subject = explicit_subjects.get(str(question.source_number or "").strip())
                        if not subject:
                            continue
                        label_source = "catalog_section_ranges"
                    else:
                        subject = forced_subject or "data"
                        label_source = "catalog_subject" if forced_subject else "parsed_section"
                    option_texts = [option.text for option in question.options]
                    image_count = len(question.stem_assets) + material_images + sum(
                        1 for option in question.options if option.image_path
                    )
                    yield {
                        "source_pdf": source_pdf,
                        "source_form": source_form,
                        "subject": subject,
                        "label_source": label_source,
                        "question_no": question.source_number,
                        "stem": question.stem,
                        "options": option_texts,
                        "material_header": material.header,
                        "material_text": material_text,
                        "image_count": image_count,
                        "page_numbers": list(question.page_numbers),
                        "feature_record": build_subject_feature_record(
                            stem=question.stem,
                            options=option_texts,
                            material_text=material_text,
                            material_header=material.header,
                            image_count=image_count,
                        ),
                    }
        else:
            for question in section.questions:
                if explicit_subjects:
                    subject = explicit_subjects.get(str(question.source_number or "").strip())
                    if not subject:
                        continue
                    label_source = "catalog_section_ranges"
                else:
                    subject = forced_subject or section.kind
                    label_source = "catalog_subject" if forced_subject else "parsed_section"
                if subject == "unknown":
                    continue
                option_texts = [option.text for option in question.options]
                image_count = len(question.stem_assets) + sum(1 for option in question.options if option.image_path)
                yield {
                    "source_pdf": source_pdf,
                    "source_form": source_form,
                    "subject": subject,
                    "label_source": label_source,
                    "question_no": question.source_number,
                    "stem": question.stem,
                    "options": option_texts,
                    "material_header": "",
                    "material_text": "",
                    "image_count": image_count,
                    "page_numbers": list(question.page_numbers),
                    "feature_record": build_subject_feature_record(
                        stem=question.stem,
                        options=option_texts,
                        image_count=image_count,
                    ),
                }


def build_dataset(catalog_path: Path, output_path: Path, *, forms: set[str] | None = None) -> dict[str, Any]:
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    rows: list[dict[str, Any]] = []
    subject_counter: Counter[str] = Counter()
    pdf_counter: Counter[str] = Counter()

    for entry in catalog.get("pdfs", []) or []:
        source_form = str(entry.get("form", "unknown")).strip() or "unknown"
        if forms and source_form not in forms:
            continue
        rel_path = Path(str(entry.get("path", "")))
        pdf_path = ROOT / rel_path
        if not pdf_path.exists():
            continue
        project = build_exam_project_from_pdf(
            str(pdf_path),
            mode="all",
            document_subject_hint=_catalog_subject_hint(entry),
        )
        forced_subject = str(entry.get("subject", "")).strip() or None
        explicit_subjects = _catalog_subject_map(entry)
        row_iter = _iter_project_rows(
            project,
            rel_path.as_posix(),
            source_form,
            forced_subject,
            explicit_subjects=explicit_subjects or None,
        )
        emitted = False
        for row in row_iter:
            rows.append(row)
            subject_counter.update([row["subject"]])
            pdf_counter.update([row["source_pdf"]])
            emitted = True
        if not emitted and forced_subject:
            for row in _fallback_rows_from_pdf(
                pdf_path,
                source_pdf=rel_path.as_posix(),
                source_form=source_form,
                forced_subject=forced_subject,
            ):
                rows.append(row)
                subject_counter.update([row["subject"]])
                pdf_counter.update([row["source_pdf"]])

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")

    summary = {
        "catalog": str(catalog_path),
        "output": str(output_path),
        "question_count": len(rows),
        "pdf_count": len(pdf_counter),
        "subject_distribution": dict(subject_counter),
        "pdf_distribution": dict(pdf_counter),
    }
    summary_path = output_path.with_suffix(".summary.json")
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="从金标准 PDF 语料构建科目分类训练集")
    parser.add_argument("--catalog", type=Path, default=DEFAULT_CATALOG)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--forms",
        nargs="*",
        default=None,
        help="仅构建指定来源 form 的样本，例如：single_subject_book set_paper",
    )
    args = parser.parse_args()

    forms = {str(item).strip() for item in (args.forms or []) if str(item).strip()} or None
    summary = build_dataset(args.catalog, args.output, forms=forms)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
