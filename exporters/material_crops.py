from __future__ import annotations

import os

from core import pdf_exam_extract as legacy_extract
from domain.models import MaterialSet


def crop_material_regions(
    pdf_path: str,
    material: MaterialSet,
    crop_dir: str,
    *,
    dpi: int = 180,
    margin: float = 16.0,
) -> list[str]:
    if not pdf_path or not material.body_regions:
        return []
    legacy_extract.require_fitz()
    if legacy_extract.fitz is None:  # pragma: no cover
        return []

    os.makedirs(crop_dir, exist_ok=True)
    paths: list[str] = []
    doc = legacy_extract.fitz.open(pdf_path)
    try:
        for index, region in enumerate(material.body_regions, 1):
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
                f"{material.material_id}_p{region.page_number}_{index}.png",
            )
            pix.save(out_path)
            paths.append(out_path)
    finally:
        doc.close()
    return paths
