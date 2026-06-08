# -*- coding: utf-8 -*-
"""扫描件 OCR 兜底。

原生 PDF 没有文字层（扫描件 / 拍照件）时，``extract_pdf_line_items`` 抽不出文字，
切题结果是 0 题。本模块用已装的本地 RapidOCR 逐页抽出文字行，合成与原生抽取
同形的 ``(text, None)`` 段，喂给同一套切题逻辑。

判定基于"逐页原生文字量"，对干净卷零误触：干净卷每页上千字，扫描比例为 0，
永远走原生路径；只有过半页面几乎无文字时才整份按扫描件 OCR。
"""
from __future__ import annotations

import logging

from core.pdf_ocr_engine import ocr_pdf_pages, synthesize_text_segments

logger = logging.getLogger(__name__)

# 一页原生文字少于此字数，视为"无文字层"。
_SCANNED_PAGE_TEXT_MAX = 50
# 抽样页中无文字层占比达到此值，整份按扫描件处理。
_SCANNED_DOC_RATIO = 0.5
# 判定时最多抽查多少页（大书不必逐页读，抽样即可可靠区分）。
_SCAN_PROBE_PAGES = 16


def document_scanned_ratio(pdf_path) -> tuple[float, int]:
    """返回 (抽样页中无文字层占比, 总页数)。打开失败返回 (0.0, 0)。"""
    try:
        import fitz
    except Exception:
        return 0.0, 0
    try:
        doc = fitz.open(pdf_path)
    except Exception:
        return 0.0, 0
    try:
        pages = doc.page_count
        if pages <= 0:
            return 0.0, 0
        if pages <= _SCAN_PROBE_PAGES:
            sample = list(range(pages))
        else:
            step = pages / _SCAN_PROBE_PAGES
            sample = sorted({int(i * step) for i in range(_SCAN_PROBE_PAGES)})
        low = 0
        for index in sample:
            try:
                text = doc[index].get_text("text")
            except Exception:
                text = ""
            if len(text.strip()) < _SCANNED_PAGE_TEXT_MAX:
                low += 1
        return (low / len(sample)), pages
    finally:
        doc.close()


def looks_scanned(pdf_path) -> bool:
    """这份 PDF 是否像扫描件（无文字层）。"""
    ratio, pages = document_scanned_ratio(pdf_path)
    return pages > 0 and ratio >= _SCANNED_DOC_RATIO


def ocr_line_items(pdf_path, *, dpi: int = 200) -> list[tuple[str, str | None]]:
    """整份 PDF 逐页 OCR，合成 ``(text, None)`` 行序列（形状同原生抽取）。"""
    _ratio, pages = document_scanned_ratio(pdf_path)
    if pages <= 0:
        return []
    ocr_pages = ocr_pdf_pages(pdf_path, range(1, pages + 1), dpi=dpi)
    segments = synthesize_text_segments(ocr_pages)
    logger.info("扫描件 OCR：%d 页，合成 %d 行文字段。", pages, len(segments))
    return segments
