"""多步修复轨迹策略模型：给定当前损坏状态，预测最优修复动作序列。

和 repair_action_model.py 的区别在于：
- repair_action_model 只预测单步最佳动作
- 本模块预测有序的多步修复方案，并为每步附带置信度
"""
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
_DEFAULT_MODEL_PATH = _ROOT / "data" / "models" / "repair_trajectory_policy.pkl"
_TRUSTED_MODEL_DIR = _DEFAULT_MODEL_PATH.parent.resolve()
_CACHE_KEY: tuple[str, float] | None = None
_CACHE_BUNDLE: dict[str, Any] | None = None
_ENABLE_MODEL_ENV = "PPTCONVERT_ENABLE_TRAJECTORY_MODEL"
_FORCE_MODEL_ENV = "PPTCONVERT_FORCE_TRAJECTORY_MODEL"
_TRUST_MODEL_ENV = "PPTCONVERT_TRUST_TRAJECTORY_MODEL"
_MODEL_PATH_ENV = "PPTCONVERT_TRAJECTORY_MODEL"

ACTION_PRIORITY: dict[str, int] = {
    "renumber_current_question": 0,
    "split_embedded_next_question": 1,
    "move_spilled_option_back": 2,
    "move_data_intro_back_to_material": 3,
    "move_data_assets_to_material": 4,
    "reassign_stem_image_to_options": 5,
}


@dataclass(frozen=True)
class TrajectoryStep:
    action: str
    confidence: float
    action_family: str = ""


@dataclass(frozen=True)
class TrajectoryPlan:
    steps: list[TrajectoryStep]

    @property
    def is_empty(self) -> bool:
        return not self.steps

    @property
    def primary_action(self) -> str | None:
        return self.steps[0].action if self.steps else None

    @property
    def primary_confidence(self) -> float:
        return self.steps[0].confidence if self.steps else 0.0

    def actions_in_order(self) -> list[str]:
        return [step.action for step in self.steps]


FAMILY_MAP: dict[str, str] = {
    "renumber_current_question": "boundary_repair",
    "split_embedded_next_question": "boundary_repair",
    "move_spilled_option_back": "boundary_repair",
    "move_data_intro_back_to_material": "material_repair",
    "move_data_assets_to_material": "material_repair",
    "reassign_stem_image_to_options": "asset_repair",
}


def _is_default_model_path(path: Path) -> bool:
    try:
        return path.resolve() == _DEFAULT_MODEL_PATH.resolve()
    except Exception:
        return False


def _model_loading_enabled(path: Path) -> bool:
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


def clear_trajectory_model_cache() -> None:
    global _CACHE_KEY, _CACHE_BUNDLE
    _CACHE_KEY = None
    _CACHE_BUNDLE = None


def default_model_path() -> Path:
    configured = os.environ.get(_MODEL_PATH_ENV, "").strip()
    return Path(configured) if configured else _DEFAULT_MODEL_PATH


def _margin_to_confidence(values: list[float]) -> list[float]:
    """将线性分类边界值映射为逐标签置信度。

    轨迹策略不是单选题，而是“同一状态下可能需要多步动作”的排序问题。
    对 decision_function 结果做逐标签 sigmoid，比 softmax 更符合这里的多标签近似语义。
    """
    if not values:
        return []
    scores: list[float] = []
    for value in values:
        bounded = max(min(float(value), 30.0), -30.0)
        scores.append(1.0 / (1.0 + math.exp(-bounded)))
    return scores


def _load_bundle(model_path: Path | None = None) -> dict[str, Any] | None:
    global _CACHE_BUNDLE, _CACHE_KEY
    path = Path(model_path or default_model_path())
    if not path.exists():
        return None
    if not _model_loading_enabled(path):
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


def predict_trajectory_plan(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
    model_path: str | os.PathLike[str] | None = None,
    min_confidence: float = 0.25,
    max_steps: int = 4,
) -> TrajectoryPlan | None:
    bundle = _load_bundle(Path(model_path) if model_path else None)
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
            probs = _margin_to_confidence([float(v) for v in values])
        else:
            return None
    except Exception:
        return None

    if len(probs) != len(labels):
        return None

    # 按概率降序排列，取置信度超过阈值的动作
    scored = sorted(zip(labels, probs), key=lambda x: x[1], reverse=True)
    steps: list[TrajectoryStep] = []
    for action_label, prob in scored:
        if prob < min_confidence or len(steps) >= max_steps:
            break
        steps.append(TrajectoryStep(
            action=action_label,
            confidence=max(0.0, min(1.0, float(prob))),
            action_family=FAMILY_MAP.get(action_label, "other_repair"),
        ))

    if not steps:
        return None

    # 按修复优先级排序
    steps.sort(key=lambda s: ACTION_PRIORITY.get(s.action, 99))
    return TrajectoryPlan(steps=steps)


def build_rule_based_trajectory_plan(
    *,
    section: Section,
    material: MaterialSet | None,
    question: QuestionNode,
    previous_question: QuestionNode | None = None,
    next_question: QuestionNode | None = None,
) -> TrajectoryPlan:
    """基于规则的修复轨迹方案（不依赖模型），用于没有模型时的兜底。"""
    steps: list[TrajectoryStep] = []

    # 题号检查
    if previous_question is not None:
        prev_num = previous_question.numeric_source_number
        cur_num = question.numeric_source_number
        if prev_num is not None and cur_num is not None and cur_num != prev_num + 1:
            steps.append(TrajectoryStep(
                action="renumber_current_question",
                confidence=0.85,
                action_family="boundary_repair",
            ))

    # 题干中嵌入下一题的题号
    stem = (question.stem or "").strip()
    if next_question is not None:
        next_num = (next_question.source_number or "").strip()
        if next_num and f"{next_num}." in stem and not (next_question.stem or "").strip():
            steps.append(TrajectoryStep(
                action="split_embedded_next_question",
                confidence=0.8,
                action_family="boundary_repair",
            ))

    # 选项不足 + 前题有串选项
    if previous_question is not None and len(previous_question.options) < 4:
        if any(letter in stem[:8] for letter in ("A.", "B.", "C.", "D.")):
            steps.append(TrajectoryStep(
                action="move_spilled_option_back",
                confidence=0.75,
                action_family="boundary_repair",
            ))

    # 资料分析材料回收
    if material is not None and section.kind == "data":
        first_q = material.questions[0] if material.questions else None
        if first_q is question:
            if not (material.body or "").strip() and not material.body_lines and len(stem) > 40:
                steps.append(TrajectoryStep(
                    action="move_data_intro_back_to_material",
                    confidence=0.7,
                    action_family="material_repair",
                ))
            if question.stem_assets and not material.body_assets:
                steps.append(TrajectoryStep(
                    action="move_data_assets_to_material",
                    confidence=0.7,
                    action_family="material_repair",
                ))

    # 图片选项归位
    if len(question.options) == 4 and question.stem_assets:
        image_opts = sum(1 for o in question.options if o.image_path)
        text_opts = sum(1 for o in question.options if (o.text or "").strip())
        if image_opts == 0 and text_opts == 0 and len(question.stem_assets) >= 4:
            steps.append(TrajectoryStep(
                action="reassign_stem_image_to_options",
                confidence=0.8,
                action_family="asset_repair",
            ))

    steps.sort(key=lambda s: ACTION_PRIORITY.get(s.action, 99))
    return TrajectoryPlan(steps=steps)
