from __future__ import annotations

import math
import os
import pickle
from dataclasses import dataclass
from pathlib import Path
from typing import Any, Iterable, Optional


_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_MODEL_PATH = _ROOT / "data" / "models" / "subject_classifier.pkl"
_TRUSTED_MODEL_DIR = _DEFAULT_MODEL_PATH.parent.resolve()
_CACHE_KEY: tuple[str, float] | None = None
_CACHE_BUNDLE: dict[str, Any] | None = None
_ENABLE_MODEL_ENV = "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL"


def _clean_text(parts: Iterable[str]) -> str:
    return "\n".join(part.strip() for part in parts if (part or "").strip()).strip()


def _count_numericish(values: Iterable[str]) -> int:
    count = 0
    for value in values:
        text = (value or "").strip()
        if not text:
            continue
        digits = sum(ch.isdigit() for ch in text)
        if digits >= max(1, len(text) // 3):
            count += 1
    return count


def _statement_option_count(values: Iterable[str]) -> int:
    count = 0
    for value in values:
        text = (value or "").strip()
        if len(text) >= 12 or any(marker in text for marker in ("正确", "错误", "符合", "不符合", "可以推出")):
            count += 1
    return count


def _blank_slot_count(text: str) -> int:
    return text.count("____") + text.count("（  ）") + text.count("()")


@dataclass(frozen=True)
class LearnedSubjectPrediction:
    probabilities: dict[str, float]

    @property
    def best_kind(self) -> str | None:
        if not self.probabilities:
            return None
        return max(self.probabilities, key=self.probabilities.get)

    @property
    def best_confidence(self) -> float:
        best_kind = self.best_kind
        if best_kind is None:
            return 0.0
        return float(self.probabilities.get(best_kind, 0.0))


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
    if os.environ.get("PPTCONVERT_TRUST_SUBJECT_MODEL", "").strip() == "1":
        return True
    return resolved.parent == _TRUSTED_MODEL_DIR


def default_model_path() -> Path:
    configured = os.environ.get("PPTCONVERT_SUBJECT_MODEL", "").strip()
    return Path(configured) if configured else _DEFAULT_MODEL_PATH


def build_subject_feature_text(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    material_header: str = "",
) -> str:
    option_texts = [text.strip() for text in (options or []) if (text or "").strip()]
    option_blob = "\n".join(f"[OPT] {text}" for text in option_texts)
    return _clean_text(
        [
            f"[MATERIAL_HEADER] {material_header}" if material_header.strip() else "",
            f"[MATERIAL] {material_text}" if material_text.strip() else "",
            f"[STEM] {stem}" if stem.strip() else "",
            option_blob,
        ]
    )


def build_subject_feature_meta(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    image_count: int = 0,
    material_header: str = "",
) -> dict[str, float]:
    option_texts = [text.strip() for text in (options or []) if (text or "").strip()]
    stem_text = _clean_text([material_header, material_text, stem])
    option_lengths = [len(text) for text in option_texts]
    return {
        "image_count": float(image_count),
        "option_count": float(len(option_texts)),
        "stem_length": float(len(stem_text)),
        "material_length": float(len((material_text or "").strip())),
        "material_header": 1.0 if (material_header or "").strip() else 0.0,
        "digit_count": float(sum(ch.isdigit() for ch in stem_text)),
        "numeric_option_count": float(_count_numericish(option_texts)),
        "statement_option_count": float(_statement_option_count(option_texts)),
        "avg_option_length": float(sum(option_lengths) / len(option_lengths)) if option_lengths else 0.0,
        "blank_slot_count": float(_blank_slot_count(stem_text)),
        "long_stem": 1.0 if len(stem_text) >= 42 else 0.0,
    }


def build_subject_feature_record(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    image_count: int = 0,
    material_header: str = "",
) -> dict[str, Any]:
    return {
        "text": build_subject_feature_text(
            stem=stem,
            options=options,
            material_text=material_text,
            material_header=material_header,
        ),
        "meta": build_subject_feature_meta(
            stem=stem,
            options=options,
            material_text=material_text,
            image_count=image_count,
            material_header=material_header,
        ),
    }


def _softmax(values: list[float]) -> list[float]:
    if not values:
        return []
    peak = max(values)
    exps = [math.exp(value - peak) for value in values]
    total = sum(exps) or 1.0
    return [value / total for value in exps]


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
    if bundle.get("ready_for_runtime") is False and os.environ.get("PPTCONVERT_FORCE_SUBJECT_MODEL", "").strip() != "1":
        return None
    _CACHE_KEY = key
    _CACHE_BUNDLE = bundle
    return bundle


def predict_subject_distribution(
    *,
    stem: str = "",
    options: Iterable[str] | None = None,
    material_text: str = "",
    image_count: int = 0,
    material_header: str = "",
    model_path: str | os.PathLike[str] | None = None,
) -> LearnedSubjectPrediction | None:
    bundle = _load_bundle(Path(model_path) if model_path else None)
    if bundle is None:
        return None
    model = bundle.get("model")
    labels = [str(label) for label in bundle.get("labels", [])]
    if not model or not labels:
        return None
    record = build_subject_feature_record(
        stem=stem,
        options=options,
        material_text=material_text,
        image_count=image_count,
        material_header=material_header,
    )
    try:
        if hasattr(model, "predict_proba"):
            probs = list(model.predict_proba([record])[0])
        elif hasattr(model, "decision_function"):
            raw = model.decision_function([record])
            values = raw[0] if getattr(raw, "ndim", 1) > 1 else raw
            probs = _softmax([float(value) for value in values])
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
    return LearnedSubjectPrediction(probabilities=probabilities)
