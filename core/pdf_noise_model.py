from __future__ import annotations

import math
import os
import pickle
import unicodedata
from dataclasses import dataclass
from pathlib import Path
from typing import Any


_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_MODEL_PATH = _ROOT / "data" / "models" / "pdf_noise_text_classifier.pkl"
_TRUSTED_MODEL_DIR = _DEFAULT_MODEL_PATH.parent.resolve()
_CACHE_KEY: tuple[str, float] | None = None
_CACHE_BUNDLE: dict[str, Any] | None = None
_ENABLE_MODEL_ENV = "PPTCONVERT_ENABLE_PICKLED_PDF_NOISE_MODEL"
_FORCE_MODEL_ENV = "PPTCONVERT_FORCE_PDF_NOISE_MODEL"
_TRUST_MODEL_ENV = "PPTCONVERT_TRUST_PDF_NOISE_MODEL"
_MODEL_PATH_ENV = "PPTCONVERT_PDF_NOISE_MODEL"

_AD_KEYWORDS: tuple[str, ...] = (
    "微信",
    "加微",
    "扫码",
    "关注",
    "资料",
    "购买",
    "领取",
    "答案",
    "解析",
    "广告",
)
_PAGE_KEYWORDS: tuple[str, ...] = (
    "第",
    "页",
    "共",
)
_HEADER_KEYWORDS: tuple[str, ...] = (
    "国家公务员录用考试",
    "行政职业能力测验",
    "行测",
)


@dataclass(frozen=True)
class PdfNoisePrediction:
    probabilities: dict[str, float]

    @property
    def best_label(self) -> str | None:
        if not self.probabilities:
            return None
        return max(self.probabilities, key=self.probabilities.get)

    @property
    def best_confidence(self) -> float:
        best_label = self.best_label
        if best_label is None:
            return 0.0
        return float(self.probabilities.get(best_label, 0.0))


try:
    from sklearn.base import BaseEstimator, TransformerMixin
except Exception:  # pragma: no cover - sklearn is optional at runtime
    class BaseEstimator:  # type: ignore[override]
        pass

    class TransformerMixin:  # type: ignore[override]
        pass


class TextFieldExtractor(BaseEstimator, TransformerMixin):
    def __init__(self, field: str):
        self.field = field

    def fit(self, X, y=None):
        return self

    def transform(self, X):
        return [str((row or {}).get(self.field, "")) for row in X]


class MetaFieldExtractor(BaseEstimator, TransformerMixin):
    def __init__(self, field: str):
        self.field = field

    def fit(self, X, y=None):
        return self

    def transform(self, X):
        return [dict((row or {}).get(self.field, {}) or {}) for row in X]


def _clean_text(text: str) -> str:
    return unicodedata.normalize("NFKC", (text or "").strip())


def _safe_ratio(numerator: float, denominator: float) -> float:
    if denominator <= 0:
        return 0.0
    return float(numerator) / float(denominator)


def _contains_any(text: str, needles: tuple[str, ...]) -> bool:
    return any(needle in text for needle in needles if needle)


def _keyword_count(text: str, needles: tuple[str, ...]) -> int:
    return sum(1 for needle in needles if needle and needle in text)


def _is_default_model_path(path: Path) -> bool:
    try:
        return path.resolve() == _DEFAULT_MODEL_PATH.resolve()
    except Exception:
        return False


def _pickle_model_loading_enabled(path: Path) -> bool:
    configured = os.environ.get(_ENABLE_MODEL_ENV, "").strip().lower()
    if configured in {"0", "false", "no"}:
        return False
    if configured in {"1", "true", "yes"}:
        return True
    return _is_default_model_path(path)


def _is_trusted_model_path(path: Path) -> bool:
    try:
        resolved = path.resolve()
    except Exception:
        return False
    if os.environ.get(_TRUST_MODEL_ENV, "").strip() == "1":
        return True
    return resolved.parent == _TRUSTED_MODEL_DIR


def default_model_path() -> Path:
    configured = os.environ.get(_MODEL_PATH_ENV, "").strip()
    return Path(configured) if configured else _DEFAULT_MODEL_PATH


def build_pdf_noise_feature_text(
    text: str,
    *,
    x0: float = 0.0,
    y0: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    page_width: float = 0.0,
    page_height: float = 0.0,
) -> str:
    normalized = _clean_text(text)
    center_x = (float(x0) + float(x1)) / 2.0
    zone = "middle"
    if _safe_ratio(y0, page_height) <= 0.14:
        zone = "top"
    elif _safe_ratio(y1, page_height) >= 0.86:
        zone = "bottom"
    side = "center"
    if _safe_ratio(x1, page_width) <= 0.45:
        side = "left"
    elif _safe_ratio(x0, page_width) >= 0.55:
        side = "right"
    centered = "yes" if abs(center_x - page_width / 2.0) <= max(12.0, page_width * 0.08) else "no"
    return "\n".join(
        part
        for part in (
            f"[ZONE] {zone}",
            f"[SIDE] {side}",
            f"[CENTERED] {centered}",
            f"[TEXT] {normalized}" if normalized else "",
        )
        if part
    )


def build_pdf_noise_feature_meta(
    text: str,
    *,
    x0: float = 0.0,
    y0: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    page_width: float = 0.0,
    page_height: float = 0.0,
    page_number: int = 1,
    line_index_in_block: int = 0,
    line_count_in_block: int = 1,
) -> dict[str, float]:
    normalized = _clean_text(text)
    width = max(0.0, float(x1) - float(x0))
    height = max(0.0, float(y1) - float(y0))
    center_x = (float(x0) + float(x1)) / 2.0
    text_length = len(normalized)
    digit_count = sum(ch.isdigit() for ch in normalized)
    cjk_count = sum("\u4e00" <= ch <= "\u9fff" for ch in normalized)
    punctuation_count = sum(not ch.isalnum() and not ("\u4e00" <= ch <= "\u9fff") for ch in normalized)
    whitespace_count = sum(ch.isspace() for ch in normalized)
    top_ratio = _safe_ratio(y0, page_height)
    bottom_ratio = _safe_ratio(y1, page_height)
    center_offset_ratio = _safe_ratio(abs(center_x - page_width / 2.0), page_width)
    width_ratio = _safe_ratio(width, page_width)
    height_ratio = _safe_ratio(height, page_height)
    return {
        "page_number": float(page_number),
        "line_index_in_block": float(line_index_in_block),
        "line_count_in_block": float(max(1, line_count_in_block)),
        "line_index_ratio": _safe_ratio(line_index_in_block, max(1, line_count_in_block - 1)),
        "top_ratio": top_ratio,
        "bottom_ratio": bottom_ratio,
        "width_ratio": width_ratio,
        "height_ratio": height_ratio,
        "center_offset_ratio": center_offset_ratio,
        "text_length": float(text_length),
        "digit_count": float(digit_count),
        "cjk_count": float(cjk_count),
        "punctuation_count": float(punctuation_count),
        "whitespace_count": float(whitespace_count),
        "is_numeric_only": 1.0 if normalized.isdigit() else 0.0,
        "single_char": 1.0 if text_length == 1 else 0.0,
        "short_text": 1.0 if 0 < text_length <= 18 else 0.0,
        "top_band": 1.0 if top_ratio <= 0.14 else 0.0,
        "bottom_band": 1.0 if bottom_ratio >= 0.86 else 0.0,
        "edge_band": 1.0 if top_ratio <= 0.14 or bottom_ratio >= 0.86 else 0.0,
        "centered": 1.0 if center_offset_ratio <= 0.08 else 0.0,
        "ad_keyword_count": float(_keyword_count(normalized, _AD_KEYWORDS)),
        "page_keyword_count": float(_keyword_count(normalized, _PAGE_KEYWORDS)),
        "header_keyword_count": float(_keyword_count(normalized, _HEADER_KEYWORDS)),
        "contains_ad_keyword": 1.0 if _contains_any(normalized, _AD_KEYWORDS) else 0.0,
        "contains_page_keyword": 1.0 if _contains_any(normalized, _PAGE_KEYWORDS) else 0.0,
        "contains_header_keyword": 1.0 if _contains_any(normalized, _HEADER_KEYWORDS) else 0.0,
    }


def build_pdf_noise_feature_record(
    text: str,
    *,
    x0: float = 0.0,
    y0: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    page_width: float = 0.0,
    page_height: float = 0.0,
    page_number: int = 1,
    line_index_in_block: int = 0,
    line_count_in_block: int = 1,
) -> dict[str, Any]:
    return {
        "text": build_pdf_noise_feature_text(
            text,
            x0=x0,
            y0=y0,
            x1=x1,
            y1=y1,
            page_width=page_width,
            page_height=page_height,
        ),
        "meta": build_pdf_noise_feature_meta(
            text,
            x0=x0,
            y0=y0,
            x1=x1,
            y1=y1,
            page_width=page_width,
            page_height=page_height,
            page_number=page_number,
            line_index_in_block=line_index_in_block,
            line_count_in_block=line_count_in_block,
        ),
    }


def _softmax(values: list[float]) -> list[float]:
    if not values:
        return []
    peak = max(values)
    exps = [math.exp(value - peak) for value in values]
    total = sum(exps) or 1.0
    return [value / total for value in exps]


def _binary_decision_probs(score: float) -> tuple[float, float]:
    positive = 1.0 / (1.0 + math.exp(-score))
    return 1.0 - positive, positive


def _load_bundle(model_path: Path | None = None) -> dict[str, Any] | None:
    global _CACHE_BUNDLE, _CACHE_KEY
    path = Path(model_path or default_model_path())
    if not path.exists():
        return None
    if not _pickle_model_loading_enabled(path):
        return None
    if not _is_trusted_model_path(path):
        return None
    key = (str(path.resolve()), path.stat().st_mtime)
    if _CACHE_KEY == key:
        return _CACHE_BUNDLE
    try:
        with path.open("rb") as fh:
            bundle = pickle.load(fh)
    except Exception:
        return None
    if not isinstance(bundle, dict) or "model" not in bundle or "labels" not in bundle:
        return None
    if bundle.get("ready_for_runtime") is False and os.environ.get(_FORCE_MODEL_ENV, "").strip() != "1":
        return None
    _CACHE_KEY = key
    _CACHE_BUNDLE = bundle
    return bundle


def predict_pdf_noise_distribution(
    text: str,
    *,
    x0: float = 0.0,
    y0: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    page_width: float = 0.0,
    page_height: float = 0.0,
    page_number: int = 1,
    line_index_in_block: int = 0,
    line_count_in_block: int = 1,
    model_path: str | os.PathLike[str] | None = None,
) -> PdfNoisePrediction | None:
    bundle = _load_bundle(Path(model_path) if model_path else None)
    if bundle is None:
        return None
    model = bundle.get("model")
    labels = [str(label) for label in bundle.get("labels", [])]
    if not model or not labels:
        return None
    record = build_pdf_noise_feature_record(
        text,
        x0=x0,
        y0=y0,
        x1=x1,
        y1=y1,
        page_width=page_width,
        page_height=page_height,
        page_number=page_number,
        line_index_in_block=line_index_in_block,
        line_count_in_block=line_count_in_block,
    )
    try:
        if hasattr(model, "predict_proba"):
            probs = list(model.predict_proba([record])[0])
        elif hasattr(model, "decision_function"):
            raw = model.decision_function([record])
            if getattr(raw, "ndim", 1) > 1:
                values = [float(value) for value in raw[0]]
                probs = _softmax(values)
            elif len(labels) == 2:
                neg, pos = _binary_decision_probs(float(raw[0]))
                probs = [neg, pos]
            else:
                return None
        else:
            return None
    except Exception:
        return None
    if len(probs) != len(labels):
        return None
    probabilities = {
        label: max(0.0, min(1.0, float(prob)))
        for label, prob in zip(labels, probs)
    }
    return PdfNoisePrediction(probabilities=probabilities)
