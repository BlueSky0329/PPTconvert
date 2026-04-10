from __future__ import annotations

from dataclasses import dataclass
import unicodedata

from core import pdf_exam_extract as legacy_extract


@dataclass
class PageTextLine:
    text: str
    page_number: int
    x0: float
    y0: float
    x1: float
    y1: float
    block_x0: float | None = None
    block_y0: float | None = None
    block_x1: float | None = None
    block_y1: float | None = None


def _line_text(line: dict) -> str:
    parts: list[str] = []
    for span in line.get("spans") or []:
        text = span.get("text") or ""
        if text:
            parts.append(text)
    return "".join(parts).strip()


def _normalize_text(text: str) -> str:
    return unicodedata.normalize("NFKC", (text or "").strip())


def extract_pdf_text_lines(pdf_path: str) -> list[PageTextLine]:
    legacy_extract.require_fitz()
    if legacy_extract.fitz is None:  # pragma: no cover
        return []

    lines: list[PageTextLine] = []
    doc = legacy_extract.fitz.open(pdf_path)
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            data = page.get_text("dict")
            blocks = legacy_extract._order_page_blocks(page, list(data.get("blocks") or []))
            for block in blocks:
                if block.get("type") != 0:
                    continue
                block_x0, block_y0, block_x1, block_y1 = [
                    float(value) for value in (block.get("bbox") or (0.0, 0.0, 0.0, 0.0))
                ]
                page_lines = sorted(block.get("lines") or [], key=legacy_extract._line_sort_key)
                for line in page_lines:
                    text = _normalize_text(_line_text(line))
                    if not text or legacy_extract._is_noise_text_line(text):
                        continue
                    x0, y0, x1, y1 = [float(value) for value in (line.get("bbox") or block.get("bbox"))]
                    lines.append(
                        PageTextLine(
                            text=text,
                            page_number=page_index + 1,
                            x0=x0,
                            y0=y0,
                            x1=x1,
                            y1=y1,
                            block_x0=block_x0,
                            block_y0=block_y0,
                            block_x1=block_x1,
                            block_y1=block_y1,
                        )
                    )
    finally:
        doc.close()
    return lines
