from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.pdf_exam_extract import (
    _is_decorative_text_record,
    _is_noise_model_candidate,
    _is_noise_text_line,
    iter_pdf_text_line_records,
)
from core.pdf_noise_model import build_pdf_noise_feature_record


DEFAULT_CATALOG = ROOT / "data" / "gold_pdf_catalog.json"
DEFAULT_OUTPUT = ROOT / "data" / "datasets" / "pdf_noise_text.jsonl"


def _label_record(record: dict) -> tuple[str, str]:
    text = str(record.get("text") or "")
    if _is_noise_text_line(text):
        return "noise", "rule_text_noise"
    if _is_decorative_text_record(record):
        return "noise", "rule_layout_noise"
    return "content", "default_content"


def build_dataset(catalog_path: Path, output_path: Path) -> dict:
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    rows: list[dict] = []
    label_counter: Counter[str] = Counter()
    source_counter: Counter[str] = Counter()
    pdf_counter: Counter[str] = Counter()

    for entry in catalog.get("pdfs", []) or []:
        rel_path = Path(str(entry.get("path", "")))
        pdf_path = ROOT / rel_path
        if not pdf_path.exists():
            continue
        source_pdf = rel_path.as_posix()
        emitted = 0
        for record in iter_pdf_text_line_records(str(pdf_path)):
            label, label_source = _label_record(record)
            if label != "noise" and not _is_noise_model_candidate(record):
                continue
            rows.append(
                {
                    "source_pdf": source_pdf,
                    "label": label,
                    "label_source": label_source,
                    "text": str(record.get("text") or ""),
                    "page_number": int(record.get("page_number") or 1),
                    "bbox": [
                        float(record.get("x0") or 0.0),
                        float(record.get("y0") or 0.0),
                        float(record.get("x1") or 0.0),
                        float(record.get("y1") or 0.0),
                    ],
                    "feature_record": build_pdf_noise_feature_record(
                        str(record.get("text") or ""),
                        x0=float(record.get("x0") or 0.0),
                        y0=float(record.get("y0") or 0.0),
                        x1=float(record.get("x1") or 0.0),
                        y1=float(record.get("y1") or 0.0),
                        page_width=float(record.get("page_width") or 0.0),
                        page_height=float(record.get("page_height") or 0.0),
                        page_number=int(record.get("page_number") or 1),
                        line_index_in_block=int(record.get("line_index_in_block") or 0),
                        line_count_in_block=int(record.get("line_count_in_block") or 1),
                    ),
                }
            )
            label_counter.update([label])
            source_counter.update([label_source])
            emitted += 1
        if emitted:
            pdf_counter.update([source_pdf] * emitted)

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")

    summary = {
        "catalog": str(catalog_path),
        "output": str(output_path),
        "row_count": len(rows),
        "pdf_count": len(pdf_counter),
        "label_distribution": dict(label_counter),
        "label_source_distribution": dict(source_counter),
        "pdf_distribution": dict(pdf_counter),
    }
    summary_path = output_path.with_suffix(".summary.json")
    summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="从金标准 PDF 构建文本噪声训练集")
    parser.add_argument("--catalog", type=Path, default=DEFAULT_CATALOG)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    summary = build_dataset(args.catalog, args.output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
