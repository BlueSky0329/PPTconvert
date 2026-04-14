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


def _iter_project_rows(project, source_pdf: str, source_form: str, forced_subject: str | None):
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                subject = forced_subject or "data"
                material_text = _material_body_text(material)
                material_images = len(material.body_assets)
                for question in material.questions:
                    option_texts = [option.text for option in question.options]
                    image_count = len(question.stem_assets) + material_images + sum(
                        1 for option in question.options if option.image_path
                    )
                    yield {
                        "source_pdf": source_pdf,
                        "source_form": source_form,
                        "subject": subject,
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
            subject = forced_subject or section.kind
            if subject == "unknown":
                continue
            for question in section.questions:
                option_texts = [option.text for option in question.options]
                image_count = len(question.stem_assets) + sum(1 for option in question.options if option.image_path)
                yield {
                    "source_pdf": source_pdf,
                    "source_form": source_form,
                    "subject": subject,
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


def build_dataset(catalog_path: Path, output_path: Path) -> dict[str, Any]:
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    rows: list[dict[str, Any]] = []
    subject_counter: Counter[str] = Counter()
    pdf_counter: Counter[str] = Counter()

    for entry in catalog.get("pdfs", []) or []:
        rel_path = Path(str(entry.get("path", "")))
        pdf_path = ROOT / rel_path
        if not pdf_path.exists():
            continue
        project = build_exam_project_from_pdf(str(pdf_path), mode="all")
        forced_subject = str(entry.get("subject", "")).strip() or None
        source_form = str(entry.get("form", "unknown")).strip() or "unknown"
        row_iter = _iter_project_rows(project, rel_path.as_posix(), source_form, forced_subject)
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
    args = parser.parse_args()

    summary = build_dataset(args.catalog, args.output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
