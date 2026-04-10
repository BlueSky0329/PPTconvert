"""从 PDF 按阅读顺序提取文本块与图片（PyMuPDF）。"""

from __future__ import annotations

from dataclasses import dataclass
import logging
import os
import re
import tempfile
import unicodedata
from typing import Iterator

LOGGER = logging.getLogger(__name__)

try:
    import fitz  # PyMuPDF
except ImportError as exc:  # pragma: no cover
    fitz = None  # type: ignore
    _IMPORT_ERROR = exc
else:
    _IMPORT_ERROR = None


@dataclass(frozen=True)
class ExtractedImageRegion:
    path: str
    page_number: int
    x0: float
    y0: float
    x1: float
    y1: float


def require_fitz():
    if fitz is None:
        raise RuntimeError(
            "需要安装 PyMuPDF：pip install pymupdf"
        ) from _IMPORT_ERROR


_PAGE_NOISE_LINE = re.compile(r"^\s*第\s*\d+\s*页\s*[,，]\s*共\s*\d+\s*页\s*$")


def _block_sort_key(block: dict) -> tuple[float, float]:
    bbox = block.get("bbox") or (0.0, 0.0, 0.0, 0.0)
    return (float(bbox[1]), float(bbox[0]))


def _line_sort_key(line: dict) -> tuple[float, float]:
    bbox = line.get("bbox") or (0.0, 0.0, 0.0, 0.0)
    return (float(bbox[1]), float(bbox[0]))


def _block_bbox(block: dict) -> tuple[float, float, float, float]:
    bbox = block.get("bbox") or (0.0, 0.0, 0.0, 0.0)
    return tuple(float(value) for value in bbox)  # type: ignore[return-value]


def _block_text_content(block: dict) -> str:
    if block.get("type") != 0:
        return ""
    parts: list[str] = []
    for line in block.get("lines") or []:
        line_text = _line_text(line)
        if line_text:
            parts.append(line_text)
    return "".join(parts).strip()


def _block_has_substantive_content(block: dict) -> bool:
    if block.get("type") == 1:
        return True
    return bool(_block_text_content(block))


def _bbox_key(bbox: tuple[float, float, float, float]) -> tuple[float, float, float, float]:
    return tuple(round(value, 1) for value in bbox)  # type: ignore[return-value]


def _merge_page_image_blocks(page, blocks: list[dict]) -> list[dict]:
    merged = list(blocks or [])
    seen_boxes = set()
    seen_image_keys = set()
    for block in merged:
        if block.get("type") != 1:
            continue
        bbox = _block_bbox(block)
        bbox_key = _bbox_key(bbox)
        seen_boxes.add(bbox_key)
        seen_image_keys.add((int(block.get("xref") or 0), bbox_key))

    try:
        image_infos = page.get_image_info(xrefs=True)
    except Exception:
        LOGGER.warning("补充页面图片信息失败", exc_info=True)
        return merged

    for info in image_infos:
        xref = int(info.get("xref") or 0)
        bbox_raw = info.get("bbox") or ()
        if xref <= 0 or len(bbox_raw) != 4:
            continue
        bbox = tuple(float(value) for value in bbox_raw)  # type: ignore[assignment]
        bbox_key = _bbox_key(bbox)
        if bbox_key in seen_boxes or (xref, bbox_key) in seen_image_keys:
            continue
        merged.append(
            {
                "type": 1,
                "bbox": bbox,
                "xref": xref,
                "ext": "png",
            }
        )
        seen_boxes.add(bbox_key)
        seen_image_keys.add((xref, bbox_key))
    return merged


def _column_side(block: dict, page_width: float) -> str:
    x0, _y0, x1, _y1 = _block_bbox(block)
    width = max(0.0, x1 - x0)
    center = page_width / 2.0
    gutter = page_width * 0.08
    if width > page_width * 0.62:
        return "full"
    if x1 <= center + gutter:
        return "left"
    if x0 >= center - gutter:
        return "right"
    return "full"


def _order_page_blocks(page, blocks: list[dict]) -> list[dict]:
    ordered = sorted(blocks or [], key=_block_sort_key)
    candidates = [
        block
        for block in ordered
        if block.get("type") in (0, 1) and _block_has_substantive_content(block)
    ]
    if len(candidates) < 4:
        return ordered

    page_width = float(page.rect.width or 1.0)
    left = [block for block in candidates if _column_side(block, page_width) == "left"]
    right = [block for block in candidates if _column_side(block, page_width) == "right"]
    if len(left) < 2 or len(right) < 2:
        return ordered

    column_blocks = left + right
    column_top = min(_block_bbox(block)[1] for block in column_blocks)
    column_bottom = max(_block_bbox(block)[3] for block in column_blocks)
    tolerance = max(16.0, page_width * 0.015)

    top_full: list[dict] = []
    middle_full: list[dict] = []
    bottom_full: list[dict] = []
    for block in ordered:
        if block in left or block in right:
            continue
        _x0, y0, _x1, y1 = _block_bbox(block)
        if y1 <= column_top + tolerance:
            top_full.append(block)
        elif y0 >= column_bottom - tolerance:
            bottom_full.append(block)
        else:
            middle_full.append(block)

    return [
        *sorted(top_full, key=_block_sort_key),
        *sorted(left, key=_block_sort_key),
        *sorted(middle_full, key=_block_sort_key),
        *sorted(right, key=_block_sort_key),
        *sorted(bottom_full, key=_block_sort_key),
    ]


def _is_noise_text_line(text: str) -> bool:
    s = unicodedata.normalize("NFKC", (text or "").strip())
    if not s:
        return True
    if s.startswith("· 本试卷由") and "生成" in s:
        return True
    if _PAGE_NOISE_LINE.match(s):
        return True
    if "国家公务员录用考试" in s and "行测" in s:
        return True
    if s.startswith("正确答案:") or s.startswith("正确答案："):
        return True
    if s.startswith("你的答案:") or s.startswith("你的答案："):
        return True
    return False


def _is_decorative_image_block(page, block: dict) -> bool:
    bbox = block.get("bbox") or (0.0, 0.0, 0.0, 0.0)
    x0, y0, x1, y1 = [float(v) for v in bbox]
    width = max(0.0, x1 - x0)
    height = max(0.0, y1 - y0)
    page_w = float(page.rect.width or 1.0)
    page_h = float(page.rect.height or 1.0)

    if width >= page_w * 0.9 and height >= page_h * 0.75:
        return True
    if y1 <= 90 and width >= page_w * 0.7 and height <= 40:
        return True
    if height <= 5:
        return True
    return False


def _save_image_bytes(data: bytes, ext: str, out_dir: str, prefix: str, index: int) -> str:
    ext = (ext or "png").lower().lstrip(".")
    if ext not in ("png", "jpg", "jpeg", "bmp", "gif", "tiff", "webp"):
        ext = "png"
    if ext == "jpeg":
        ext = "jpg"
    path = os.path.join(out_dir, f"{prefix}_{index:04d}.{ext}")
    with open(path, "wb") as f:
        f.write(data)
    return path


def _extract_image_from_block(doc, block: dict, out_dir: str, prefix: str, index: int) -> str | None:
    """从 type==1 的块写出图片文件。"""
    img = block.get("image")
    if img:
        ext = str(block.get("ext") or "png")
        try:
            return _save_image_bytes(bytes(img), ext, out_dir, prefix, index)
        except Exception:
            LOGGER.warning("写出内嵌图片失败", exc_info=True)
            return None

    xref = block.get("xref") or 0
    if xref <= 0:
        return None
    try:
        pix = fitz.Pixmap(doc, xref)
        if pix.n - pix.alpha > 3:  # CMYK
            pix = fitz.Pixmap(fitz.csRGB, pix)
        path = os.path.join(out_dir, f"{prefix}_{index:04d}.png")
        pix.save(path)
        pix = None
        return path
    except Exception:
        LOGGER.warning("从 xref=%s 提取图片失败", xref, exc_info=True)
        return None


def _line_text(line: dict) -> str:
    parts: list[str] = []
    for span in line.get("spans") or []:
        t = span.get("text") or ""
        if t:
            parts.append(t)
    return "".join(parts).strip()


def extract_pdf_line_items_with_metadata(
    pdf_path: str,
    temp_dir: str | None = None,
) -> tuple[list[tuple[str, str | None]], str, dict[str, ExtractedImageRegion]]:
    """
    按页内块顺序提取：每个元素为 (text_line, image_path)。
    image_path 非空表示该行应插入对应图片（text 通常为空）。
    返回 (行列表, 临时目录)；临时目录内含抽取的图片，调用方可视情况 shutil.rmtree。
    """
    require_fitz()
    out_dir = temp_dir or tempfile.mkdtemp(prefix="pptconvert_pdf_")
    os.makedirs(out_dir, exist_ok=True)

    raw_segments: list[tuple[str, str | None]] = []
    image_regions: dict[str, ExtractedImageRegion] = {}
    doc = fitz.open(pdf_path)
    img_counter = 0
    try:
        for page_index in range(len(doc)):
            page = doc[page_index]
            prefix = f"p{page_index + 1}"
            data = page.get_text("dict")
            blocks = _order_page_blocks(page, _merge_page_image_blocks(page, list(data.get("blocks") or [])))
            for block in blocks:
                btype = block.get("type")
                if btype == 0:
                    line_parts: list[str] = []
                    lines = sorted(block.get("lines") or [], key=_line_sort_key)
                    for line in lines:
                        lt = _line_text(line)
                        if lt and not _is_noise_text_line(lt):
                            line_parts.append(lt)
                    if line_parts:
                        raw_segments.append(("\n".join(line_parts), None))
                elif btype == 1:
                    if _is_decorative_image_block(page, block):
                        continue
                    path = _extract_image_from_block(doc, block, out_dir, prefix, img_counter)
                    img_counter += 1
                    if path:
                        x0, y0, x1, y1 = [float(value) for value in (block.get("bbox") or (0.0, 0.0, 0.0, 0.0))]
                        image_regions[path] = ExtractedImageRegion(
                            path=path,
                            page_number=page_index + 1,
                            x0=x0,
                            y0=y0,
                            x1=x1,
                            y1=y1,
                        )
                        raw_segments.append(("", path))
    finally:
        doc.close()

    lines = segments_to_lines(raw_segments)
    return lines, out_dir, image_regions


def extract_pdf_line_items(pdf_path: str, temp_dir: str | None = None) -> tuple[list[tuple[str, str | None]], str]:
    lines, out_dir, _image_regions = extract_pdf_line_items_with_metadata(pdf_path, temp_dir)
    return lines, out_dir


def iter_page_segments(pdf_path: str, temp_dir: str | None = None) -> Iterator[tuple[str, str | None]]:
    """兼容旧名：委托 extract_pdf_line_items。"""
    items, _ = extract_pdf_line_items(pdf_path, temp_dir)
    yield from items


def segments_to_lines(segments: list[tuple[str, str | None]]) -> list[tuple[str, str | None]]:
    """
    将相邻的纯文本段合并为行列表；图片作为独立「行」占位 (空字符串, path)。
    """
    lines: list[tuple[str, str | None]] = []
    buf: list[str] = []

    def flush_text():
        nonlocal buf
        if not buf:
            return
        chunk = "\n".join(buf).strip()
        if chunk:
            for part in chunk.splitlines():
                t = unicodedata.normalize("NFKC", part.strip())
                if t:
                    lines.append((t, None))
        buf = []

    for text, img in segments:
        if img:
            flush_text()
            lines.append(("", img))
            continue
        if text and text.strip():
            buf.append(text.strip())

    flush_text()
    return lines
