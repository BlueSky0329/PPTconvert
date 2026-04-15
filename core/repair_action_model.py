from __future__ import annotations

import math
import os
import pickle
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from core.repair_action_features import build_repair_state_record
from domain.models import MaterialSet, QuestionNode, Section


_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_ACTION_MODEL_PATH = _ROOT / "data" / "models" / "repair_action_classifier.pkl"
_DEFAULT_FAMILY_MODEL_PATH = _ROOT / "data" / "models" / "repair_action_family_classifier.pkl"
_TRUSTED_MODEL_DIR = _DEFAULT_ACTION_MODEL_PATH.parent.resolve()
_ENABLE_MODEL_ENV = "PPTCONVERT_ENABLE_PICKLED_REPAIR_MODEL"
_FORCE_MODEL_ENV = "PPTCONVERT_FORCE_REPAIR_MODEL"
_TRUST_MODEL_ENV = "PPTCONVERT_TRUST_REPAIR_MODEL"
_ACTION_MODEL_ENV = "PPTCONVERT_REPAIR_MODEL"
_FAMILY_MODEL_ENV = "PPTCONVERT_REPAIR_FAMILY_MODEL"
_CACHE_KEYS: dict[str, tuple[str, float] | None] = {"action": None, "action_family": None}
_CACHE_BUNDLES: dict[str, dict[str, Any] | None] = {"action": None, "action_family": None}


@dataclass(frozen=True)
class RepairActionPrediction:
    probabilities: dict[str, float]
    label_field: str

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


def clear_repair_model_cache() -> None:
    for key in tuple(_CACHE_KEYS):
        _CACHE_KEYS[key] = None
        _CACHE_BUNDLES[key] = None


def _default_path_for_label_field(label_field: str) -> Path:
    return _DEFAULT_FAMILY_MODEL_PATH if label_field == "action_family" else _DEFAULT_ACTION_MODEL_PATH


def default_model_path(label_field: str = "action") -> Path:
    if label_field == "action_family":
        configured = os.environ.get(_FAMILY_MODEL_ENV, "").strip()
        return Path(configured) if configured else _DEFAULT_FAMILY_MODEL_PATH
    configured = os.environ.get(_ACTION_MODEL_ENV, "").strip()
    return Path(configured) if configured else _DEFAULT_ACTION_MODEL_PATH


def _is_default_model_path(path: Path, label_field: str) -> bool:
    try:
        return path.resolve() == _default_path_for_label_field(label_field).resolve()
    except Exception:
        return False


def _pickle_model_loading_enabled(path: Path, label_field: str) -> bool:
    configured = os.environ.get(_ENABLE_MODEL_ENV, "").strip().lower()
    if configured in {"0", "false", "no"}:
        return False
    if configured in {"1", "true", "yes"}:
        return True
    return _is_default_model_path(path, label_field)


def _is_trusted_model_path(path: Path) -> bool:
    try:
        resolved = path.resolve()
    except Exception:
        return False
    if os.environ.get(_TRUST_MODEL_ENV, "").strip() == "1":
        return True
    return resolved.parent == _TRUSTED_MODEL_DIR


def _softmax(values: list[float]) -> list[float]:
    if not values:
        return []
    peak = max(values)
    exps = [math.exp(value - peak) for value in values]
    total = sum(exps) or 1.0
    return [value / total for value in exps]


def _load_bundle(label_field: str, model_path: Path | None = None) -> dict[str, Any] | None:
    global _CACHE_BUNDLES, _CACHE_KEYS
    path = Path(model_path or default_model_path(label_field))
    if not path.exists():
        return None
    if not _pickle_model_loading_enabled(path, label_field):
        return None
    if not _is_trusted_model_path(path):
        return None
    key = (str(path.resolve()), path.stat().st_mtime)
    if _CACHE_KEYS.get(label_field) == key:
        return _CACHE_BUNDLES.get(label_field)
    try:
        with path.open("rb") as fh:
            bundle = pickle.load(fh)
    except Exception:
        return None
    if not isinstance(bundle, dict) or "model" not in bundle or "labels" not in bundle:
        return None
    bundle_label_field = str(bundle.get("label_field") or label_field)
    if bundle.get("label_field") and bundle_label_field != label_field:
        return None
    if bundle.get("ready_for_runtime") is False and os.environ.get(_FORCE_MODEL_ENV, "").strip() != "1":
        return None
    _CACHE_KEYS[label_field] = key
    _CACHE_BUNDLES[label_field] = bundle
    return bundle


def predict_repair_action_distribution(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
    label_field: str = "action",
    model_path: str | os.PathLike[str] | None = None,
) -> RepairActionPrediction | None:
    bundle = _load_bundle(label_field, Path(model_path) if model_path else None)
    if bundle is None:
        return None
    model = bundle.get("model")
    labels = [str(label) for label in bundle.get("labels", [])]
    if not model or not labels:
        return None
    record = build_repair_state_record(
        section=section,
        material=material,
        question=question,
        previous_question=previous_question,
        next_question=next_question,
    )
    sample = {
        "text": str(record.get("text", "")),
        "meta": dict(record.get("meta") or {}),
    }
    try:
        if hasattr(model, "predict_proba"):
            probs = list(model.predict_proba([sample])[0])
        elif hasattr(model, "decision_function"):
            raw = model.decision_function([sample])
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
    return RepairActionPrediction(probabilities=probabilities, label_field=label_field)
