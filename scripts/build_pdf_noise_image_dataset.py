from __future__ import annotations

import argparse
import json
from collections import Counter
import os
from pathlib import Path
import shutil
import sys
import tempfile

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from PIL import Image

import core.pdf_exam_extract as pdf_exam_extract
from core.pdf_exam_extract import (
    _extract_image_from_block,
    _is_decorative_image_block,
    _merge_page_image_blocks,
    _order_page_blocks,
    require_fitz,
)
from core.pdf_noise_image_model import (
    build_pdf_noise_image_meta,
    build_pdf_noise_image_visual_stats,
    is_background_like_pdf_image_meta,
    is_watermark_like_pdf_image_meta,
)


DEFAULT_CATALOG = ROOT / "data" / "gold_pdf_catalog.json"
DEFAULT_OUTPUT = ROOT / "data" / "datasets" / "pdf_noise_images.jsonl"


def _iter_image_records_from_pdf(pdf_path: Path, *, source_pdf: str, assets_dir: Path):
    require_fitz()
    doc = pdf_exam_extract.fitz.open(str(pdf_path))
    image_index = 0
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            data = page.get_text("dict")
            blocks = _order_page_blocks(page, _merge_page_image_blocks(page, list(data.get("blocks") or [])))
            for block in blocks:
                if block.get("type") != 1:
                    continue
                saved_path = _extract_image_from_block(doc, block, str(assets_dir), f"p{page_index + 1}", image_index)
                image_index += 1
                if not saved_path:
                    continue
                bbox = tuple(float(value) for value in (block.get("bbox") or (0.0, 0.0, 0.0, 0.0)))
                x0, y0, x1, y1 = bbox
                with Image.open(saved_path) as image:
                    image_width, image_height = image.size
                visual_stats = build_pdf_noise_image_visual_stats(saved_path)
                label = "noise" if _is_decorative_image_block(page, block) else "content"
                meta_record = build_pdf_noise_image_meta(
                    x0=x0,
                    y0=y0,
                    x1=x1,
                    y1=y1,
                    page_width=float(page.rect.width or 1.0),
                    page_height=float(page.rect.height or 1.0),
                    image_width=float(image_width),
                    image_height=float(image_height),
                    page_number=page_index + 1,
                    gray_mean=float(visual_stats.get("gray_mean", 0.0)),
                    gray_stddev=float(visual_stats.get("gray_stddev", 0.0)),
                    gray_range=float(visual_stats.get("gray_range", 0.0)),
                    edge_density=float(visual_stats.get("edge_density", 0.0)),
                    bright_ratio=float(visual_stats.get("bright_ratio", 0.0)),
                    dark_ratio=float(visual_stats.get("dark_ratio", 0.0)),
                )
                if label != "noise" and is_background_like_pdf_image_meta(meta_record):
                    label = "noise"
                    label_source = "rule_background_image"
                elif label != "noise" and is_watermark_like_pdf_image_meta(meta_record):
                    label = "noise"
                    label_source = "rule_watermark_image"
                else:
                    label_source = "rule_decorative_image" if label == "noise" else "default_content_image"
                yield {
                    "source_pdf": source_pdf,
                    "label": label,
                    "label_source": label_source,
                    "image_path": saved_path,
                    "page_number": page_index + 1,
                    "bbox": [x0, y0, x1, y1],
                    "page_size": [float(page.rect.width or 1.0), float(page.rect.height or 1.0)],
                    "image_size": [int(image_width), int(image_height)],
                    "meta_record": meta_record,
                }
    finally:
        doc.close()


def build_dataset(catalog_path: Path, output_path: Path) -> dict:
    require_fitz()
    catalog = json.loads(catalog_path.read_text(encoding="utf-8"))
    rows: list[dict] = []
    label_counter: Counter[str] = Counter()
    source_counter: Counter[str] = Counter()
    pdf_counter: Counter[str] = Counter()

    output_path.parent.mkdir(parents=True, exist_ok=True)
    assets_dir = output_path.parent / f"{output_path.stem}_assets"
    summary_path = output_path.with_suffix(".summary.json")

    with tempfile.TemporaryDirectory(prefix=f"{output_path.stem}_", dir=str(output_path.parent)) as tmp_dir:
        staging_root = Path(tmp_dir)
        staging_assets_dir = staging_root / assets_dir.name
        staging_output_path = staging_root / output_path.name
        staging_summary_path = staging_root / summary_path.name
        staging_assets_dir.mkdir(parents=True, exist_ok=True)

        for entry in catalog.get("pdfs", []) or []:
            rel_path = Path(str(entry.get("path", "")))
            pdf_path = ROOT / rel_path
            if not pdf_path.exists():
                continue
            source_pdf = rel_path.as_posix()
            emitted = 0
            for row in _iter_image_records_from_pdf(pdf_path, source_pdf=source_pdf, assets_dir=staging_assets_dir):
                row = dict(row)
                row["image_path"] = str(assets_dir / Path(str(row["image_path"])).name)
                rows.append(row)
                label_counter.update([row["label"]])
                source_counter.update([row["label_source"]])
                emitted += 1
            if emitted:
                pdf_counter.update([source_pdf] * emitted)

        with staging_output_path.open("w", encoding="utf-8") as fh:
            for row in rows:
                fh.write(json.dumps(row, ensure_ascii=False) + "\n")

        summary = {
            "catalog": str(catalog_path),
            "output": str(output_path),
            "assets_dir": str(assets_dir),
            "row_count": len(rows),
            "pdf_count": len({row["source_pdf"] for row in rows}),
            "label_distribution": dict(label_counter),
            "label_source_distribution": dict(source_counter),
            "pdf_distribution": dict(pdf_counter),
        }
        staging_summary_path.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

        if output_path.exists():
            output_path.unlink()
        if summary_path.exists():
            summary_path.unlink()
        if assets_dir.exists():
            shutil.rmtree(assets_dir)

        os.replace(staging_output_path, output_path)
        os.replace(staging_summary_path, summary_path)
        os.replace(staging_assets_dir, assets_dir)

    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="从金标准 PDF 构建图片噪声训练集")
    parser.add_argument("--catalog", type=Path, default=DEFAULT_CATALOG)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    summary = build_dataset(args.catalog, args.output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
