from __future__ import annotations

import os

from core import pdf_exam_extract as legacy_extract
from domain.models import MaterialSet, PageRegion


def crop_page_regions(
    pdf_path: str,
    regions: list[PageRegion],
    crop_dir: str,
    *,
    prefix: str,
    dpi: int = 180,
    margin: float = 16.0,
) -> list[str]:
    if not pdf_path or not regions:
        return []
    legacy_extract.require_fitz()
    if legacy_extract.fitz is None:  # pragma: no cover
        return []

    os.makedirs(crop_dir, exist_ok=True)
    paths: list[str] = []
    doc = legacy_extract.fitz.open(pdf_path)
    try:
        for index, region in enumerate(regions, 1):
            page = doc[region.page_number - 1]
            padded = region.padded(margin=margin)
            clip = legacy_extract.fitz.Rect(
                padded.x0,
                padded.y0,
                padded.x1,
                padded.y1,
            )
            pix = page.get_pixmap(clip=clip, dpi=dpi)
            out_path = os.path.join(
                crop_dir,
                f"{prefix}_p{region.page_number}_{index}.png",
            )
            pix.save(out_path)
            paths.append(out_path)
    finally:
        doc.close()
    return paths


def crop_material_regions(
    pdf_path: str,
    material: MaterialSet,
    crop_dir: str,
    *,
    dpi: int = 180,
    margin: float = 16.0,
) -> list[str]:
    return crop_page_regions(
        pdf_path,
        material.body_regions,
        crop_dir,
        prefix=material.material_id,
        dpi=dpi,
        margin=margin,
    )
