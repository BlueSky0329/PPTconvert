"""RapidOCR 可选适配：给扫描版 PDF 提供文字回收能力。

设计原则：
- rapidocr-onnxruntime 延迟导入，未安装只告警不抛异常
- 页级 JSON 缓存，按 (pdf 摘要, 页号, dpi, 引擎指纹) 键控
- OCRLine 保留 bbox，供后续按阅读顺序重新合成文本行
"""

from __future__ import annotations

from dataclasses import dataclass
import hashlib
import json
import logging
import os
from pathlib import Path
from typing import Any, Iterable

LOGGER = logging.getLogger(__name__)

_ROOT = Path(__file__).resolve().parent.parent
_CACHE_DIR = _ROOT / "data" / "cache" / "ocr"
_DISABLE_ENV = "PPTCONVERT_DISABLE_OCR"

_ENGINE_FINGERPRINT = "rapidocr-onnx-v1"

_ENGINE_CACHE: Any = None
_ENGINE_LOAD_FAILED = False


@dataclass(frozen=True)
class OCRLine:
    text: str
    bbox: tuple[float, float, float, float]
    confidence: float
    page_number: int

    def to_dict(self) -> dict[str, Any]:
        return {
            "text": self.text,
            "bbox": list(self.bbox),
            "confidence": self.confidence,
            "page_number": self.page_number,
        }

    @classmethod
    def from_dict(cls, data: dict[str, Any]) -> "OCRLine":
        raw_box = data.get("bbox") or (0.0, 0.0, 0.0, 0.0)
        try:
            bbox = tuple(float(v) for v in raw_box)
        except (TypeError, ValueError):
            bbox = (0.0, 0.0, 0.0, 0.0)
        if len(bbox) != 4:
            bbox = (0.0, 0.0, 0.0, 0.0)
        try:
            confidence = float(data.get("confidence") or 0.0)
        except (TypeError, ValueError):
            confidence = 0.0
        try:
            page_number = int(data.get("page_number") or 0)
        except (TypeError, ValueError):
            page_number = 0
        return cls(
            text=str(data.get("text") or ""),
            bbox=bbox,  # type: ignore[arg-type]
            confidence=confidence,
            page_number=page_number,
        )


def _ocr_disabled_via_env() -> bool:
    value = os.environ.get(_DISABLE_ENV, "").strip().lower()
    return value in {"1", "true", "yes"}


def _load_engine() -> Any:
    global _ENGINE_CACHE, _ENGINE_LOAD_FAILED
    if _ENGINE_CACHE is not None:
        return _ENGINE_CACHE
    if _ENGINE_LOAD_FAILED:
        return None
    if _ocr_disabled_via_env():
        return None
    try:
        from rapidocr_onnxruntime import RapidOCR  # type: ignore[import-not-found]
    except ImportError:
        _ENGINE_LOAD_FAILED = True
        LOGGER.info("rapidocr-onnxruntime 未安装，OCR 能力不可用")
        return None
    try:
        _ENGINE_CACHE = RapidOCR()
    except Exception as exc:  # noqa: BLE001 — 初始化失败路径多样，统一降级
        _ENGINE_LOAD_FAILED = True
        LOGGER.warning("RapidOCR 初始化失败: %s", exc)
        return None
    return _ENGINE_CACHE


def is_ocr_available() -> bool:
    return _load_engine() is not None


def reset_engine_cache() -> None:
    global _ENGINE_CACHE, _ENGINE_LOAD_FAILED
    _ENGINE_CACHE = None
    _ENGINE_LOAD_FAILED = False


def _fingerprint_pdf(pdf_path: Path) -> str:
    stat = pdf_path.stat()
    raw = f"{pdf_path.resolve()}|{stat.st_size}|{int(stat.st_mtime)}".encode("utf-8")
    return hashlib.md5(raw).hexdigest()


def _cache_path(pdf_fingerprint: str, page_number: int, dpi: int) -> Path:
    name = f"{pdf_fingerprint}_p{page_number}_dpi{dpi}_{_ENGINE_FINGERPRINT}.json"
    return _CACHE_DIR / name


def _load_cache(cache_file: Path) -> list[OCRLine] | None:
    if not cache_file.exists():
        return None
    try:
        payload = json.loads(cache_file.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return None
    if not isinstance(payload, dict):
        return None
    if payload.get("engine_fingerprint") != _ENGINE_FINGERPRINT:
        return None
    raw_lines = payload.get("lines")
    if not isinstance(raw_lines, list):
        return None
    return [OCRLine.from_dict(item) for item in raw_lines if isinstance(item, dict)]


def _write_cache(cache_file: Path, page_number: int, lines: list[OCRLine]) -> None:
    try:
        cache_file.parent.mkdir(parents=True, exist_ok=True)
        payload = {
            "engine_fingerprint": _ENGINE_FINGERPRINT,
            "page_number": page_number,
            "lines": [line.to_dict() for line in lines],
        }
        cache_file.write_text(json.dumps(payload, ensure_ascii=False), encoding="utf-8")
    except OSError as exc:
        LOGGER.warning("写 OCR 缓存失败 %s: %s", cache_file, exc)


def _render_page_image(pdf_path: Path, page_number: int, dpi: int) -> bytes | None:
    try:
        import fitz  # type: ignore[import-not-found]
    except ImportError:
        LOGGER.warning("PyMuPDF 未安装，无法渲染页面给 OCR 使用")
        return None
    try:
        with fitz.open(pdf_path) as doc:
            if page_number < 1 or page_number > len(doc):
                return None
            page = doc[page_number - 1]
            zoom = dpi / 72.0
            matrix = fitz.Matrix(zoom, zoom)
            pix = page.get_pixmap(matrix=matrix, alpha=False)
            return pix.tobytes("png")
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("渲染页面 %d 失败: %s", page_number, exc)
        return None


def _parse_rapidocr_result(result: Any, page_number: int) -> list[OCRLine]:
    if not result:
        return []
    lines: list[OCRLine] = []
    for entry in result:
        if not isinstance(entry, (list, tuple)) or len(entry) < 3:
            continue
        raw_box, text, conf = entry[0], entry[1], entry[2]
        if text is None:
            continue
        text_str = str(text).strip()
        if not text_str:
            continue
        xs: list[float] = []
        ys: list[float] = []
        if isinstance(raw_box, (list, tuple)):
            for point in raw_box:
                if isinstance(point, (list, tuple)) and len(point) >= 2:
                    try:
                        xs.append(float(point[0]))
                        ys.append(float(point[1]))
                    except (TypeError, ValueError):
                        continue
        if len(xs) < 2 or len(ys) < 2:
            continue
        try:
            conf_val = float(conf)
        except (TypeError, ValueError):
            conf_val = 0.0
        lines.append(
            OCRLine(
                text=text_str,
                bbox=(min(xs), min(ys), max(xs), max(ys)),
                confidence=conf_val,
                page_number=page_number,
            )
        )
    return lines


def _run_ocr(engine: Any, image_bytes: bytes, page_number: int) -> list[OCRLine]:
    try:
        import io

        import numpy as np  # type: ignore[import-not-found]
        from PIL import Image
    except ImportError as exc:
        LOGGER.warning("OCR 依赖缺失 (numpy/Pillow): %s", exc)
        return []
    try:
        image = Image.open(io.BytesIO(image_bytes)).convert("RGB")
        array = np.array(image)
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("OCR 图像解码失败: %s", exc)
        return []
    try:
        output = engine(array)
    except Exception as exc:  # noqa: BLE001
        LOGGER.warning("OCR 推理失败 (page %d): %s", page_number, exc)
        return []
    if isinstance(output, tuple) and output:
        result = output[0]
    else:
        result = output
    return _parse_rapidocr_result(result, page_number)


def ocr_pdf_page(
    pdf_path: str | os.PathLike[str],
    page_number: int,
    *,
    dpi: int = 200,
    use_cache: bool = True,
) -> list[OCRLine]:
    """对单页 PDF 做 OCR，返回行级结果。page_number 从 1 开始。

    OCR 引擎不可用、PDF 不存在、渲染或推理失败时返回空列表，不抛异常。
    """
    path = Path(pdf_path)
    if not path.exists():
        return []
    if page_number < 1:
        return []
    engine = _load_engine()
    if engine is None:
        return []

    fingerprint = _fingerprint_pdf(path)
    cache_file = _cache_path(fingerprint, page_number, dpi)
    if use_cache:
        cached = _load_cache(cache_file)
        if cached is not None:
            return cached

    image_bytes = _render_page_image(path, page_number, dpi)
    if image_bytes is None:
        return []

    lines = _run_ocr(engine, image_bytes, page_number)

    if use_cache:
        _write_cache(cache_file, page_number, lines)

    return lines


def ocr_pdf_pages(
    pdf_path: str | os.PathLike[str],
    page_numbers: Iterable[int],
    *,
    dpi: int = 200,
    use_cache: bool = True,
) -> dict[int, list[OCRLine]]:
    result: dict[int, list[OCRLine]] = {}
    for page_number in page_numbers:
        result[page_number] = ocr_pdf_page(
            pdf_path,
            page_number,
            dpi=dpi,
            use_cache=use_cache,
        )
    return result


def synthesize_text_segments(
    pages: dict[int, list[OCRLine]],
    *,
    min_confidence: float = 0.3,
    line_band_px: float = 8.0,
) -> list[tuple[str, str | None]]:
    """把按页聚合的 OCRLine 转成 ``extract_pdf_line_items`` 风格的 ``(text, None)`` 段。

    - 每页内按 y 轴分行（容差 ``line_band_px``），同一行内按 x 轴从左到右拼接
    - 各页按页号升序输出
    - 过滤置信度低于 ``min_confidence`` 的行
    """
    segments: list[tuple[str, str | None]] = []
    for page_number in sorted(pages.keys()):
        lines = pages[page_number]
        if not lines:
            continue
        band = max(line_band_px, 1.0)
        filtered = [
            line
            for line in lines
            if line.confidence >= min_confidence and line.text.strip()
        ]
        if not filtered:
            continue
        filtered.sort(key=lambda line: (round(line.bbox[1] / band), line.bbox[0]))
        grouped: list[list[OCRLine]] = []
        current: list[OCRLine] = []
        current_band: float | None = None
        for line in filtered:
            key = round(line.bbox[1] / band)
            if current_band is None or key == current_band:
                current.append(line)
            else:
                grouped.append(current)
                current = [line]
            current_band = key
        if current:
            grouped.append(current)
        for row in grouped:
            row.sort(key=lambda item: item.bbox[0])
            text = " ".join(item.text.strip() for item in row if item.text.strip()).strip()
            if text:
                segments.append((text, None))
    return segments


def clear_ocr_cache(pdf_path: str | os.PathLike[str] | None = None) -> int:
    """清理 OCR 缓存。传入 pdf_path 只清该文件；不传则全清。返回删除条数。"""
    if not _CACHE_DIR.exists():
        return 0
    if pdf_path is None:
        pattern = "*.json"
    else:
        path = Path(pdf_path)
        if not path.exists():
            return 0
        pattern = f"{_fingerprint_pdf(path)}_*.json"
    removed = 0
    for cache_file in _CACHE_DIR.glob(pattern):
        try:
            cache_file.unlink()
            removed += 1
        except OSError:
            continue
    return removed
