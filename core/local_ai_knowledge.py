from __future__ import annotations

import json
from functools import lru_cache
from pathlib import Path
from typing import Any


_ROOT = Path(__file__).resolve().parent.parent
_KNOWLEDGE_PATH = _ROOT / "data" / "local_ai_knowledge.json"
_CORPUS_PROFILE_PATH = _ROOT / "data" / "gongkao_corpus_profiles.json"


@lru_cache(maxsize=1)
def load_local_ai_knowledge() -> dict[str, Any]:
    if not _KNOWLEDGE_PATH.exists():
        return {}
    with _KNOWLEDGE_PATH.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data if isinstance(data, dict) else {}


@lru_cache(maxsize=1)
def load_corpus_profiles() -> dict[str, Any]:
    if not _CORPUS_PROFILE_PATH.exists():
        return {}
    with _CORPUS_PROFILE_PATH.open("r", encoding="utf-8") as fh:
        data = json.load(fh)
    return data if isinstance(data, dict) else {}


def subject_profile(kind: str) -> dict[str, Any]:
    return dict(load_local_ai_knowledge().get("subject_profiles", {}).get(kind, {}) or {})


def subject_keywords(kind: str) -> tuple[str, ...]:
    profile = subject_profile(kind)
    values = profile.get("positive_keywords", []) or []
    return tuple(str(value) for value in values if str(value).strip())


def subject_negative_keywords(kind: str) -> tuple[str, ...]:
    profile = subject_profile(kind)
    values = profile.get("negative_keywords", []) or []
    return tuple(str(value) for value in values if str(value).strip())


def subject_structural_markers(kind: str) -> tuple[str, ...]:
    profile = subject_profile(kind)
    values = profile.get("structural_markers", []) or []
    return tuple(str(value) for value in values if str(value).strip())


def subject_subtypes(kind: str) -> list[dict[str, Any]]:
    profile = subject_profile(kind)
    values = profile.get("subtypes", []) or []
    return [dict(value) for value in values if isinstance(value, dict)]


def official_sources() -> list[dict[str, Any]]:
    values = load_local_ai_knowledge().get("official_sources", []) or []
    return [dict(value) for value in values if isinstance(value, dict)]


def public_sources() -> list[dict[str, Any]]:
    values = load_local_ai_knowledge().get("public_sources", []) or []
    return [dict(value) for value in values if isinstance(value, dict)]


def corpus_datasets() -> list[dict[str, Any]]:
    values = load_corpus_profiles().get("datasets", []) or []
    return [dict(value) for value in values if isinstance(value, dict)]


def confidence_settings() -> dict[str, float]:
    raw = load_local_ai_knowledge().get("confidence", {}) or {}
    settings: dict[str, float] = {}
    for key, value in raw.items():
        try:
            settings[str(key)] = float(value)
        except (TypeError, ValueError):
            continue
    return settings
