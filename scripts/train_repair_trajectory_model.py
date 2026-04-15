"""训练多步修复轨迹策略模型。

输入：repair_trajectories.jsonl（多步轨迹数据集）
输出：repair_trajectory_policy.pkl（多标签分类器，预测给定状态下哪些动作需要执行）
"""
from __future__ import annotations

import argparse
import json
import pickle
from collections import Counter, defaultdict
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.learned_subject_model import MetaFieldExtractor, TextFieldExtractor
from scripts.train_local_subject_model import _choose_grouped_split

DEFAULT_DATASET = ROOT / "data" / "datasets" / "repair_trajectories.jsonl"
DEFAULT_MODEL = ROOT / "data" / "models" / "repair_trajectory_policy.pkl"

ALL_ACTIONS = [
    "move_data_assets_to_material",
    "move_data_intro_back_to_material",
    "move_spilled_option_back",
    "reassign_stem_image_to_options",
    "renumber_current_question",
    "split_embedded_next_question",
]


def _rebalance_training_split(samples: list[dict], labels: list[str]) -> tuple[list[dict], list[str], dict[str, int], dict[str, int]]:
    grouped: dict[str, list[int]] = defaultdict(list)
    for index, label in enumerate(labels):
        grouped[label].append(index)

    before = {label: len(indexes) for label, indexes in grouped.items()}
    if not grouped:
        return samples, labels, before, before

    majority_count = max(before.values())
    majority_cap = min(1800, majority_count)
    target_counts: dict[str, int] = {}
    for label, count in before.items():
        target_counts[label] = min(majority_cap, max(count, max(72, count * 8)))

    balanced_samples: list[dict] = []
    balanced_labels: list[str] = []
    after: dict[str, int] = {}
    for label, indexes in grouped.items():
        target = target_counts[label]
        count = len(indexes)
        repeats, remainder = divmod(target, count)
        selected_indexes = indexes * repeats + indexes[:remainder]
        for index in selected_indexes:
            balanced_samples.append(samples[index])
            balanced_labels.append(label)
        after[label] = len(selected_indexes)
    return balanced_samples, balanced_labels, before, after


def _load_trajectories(path: Path) -> list[dict]:
    rows: list[dict] = []
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            rows.append(json.loads(line))
    return rows


def _flatten_to_step_samples(trajectories: list[dict]) -> tuple[list[dict], list[str], list[str]]:
    """将轨迹展开为 (state_record, primary_action, source_pdf) 的训练样本。

    只取每条轨迹的第一步作为训练标签，因为模型在运行时按步递进。
    """
    samples: list[dict] = []
    labels: list[str] = []
    groups: list[str] = []
    for traj in trajectories:
        steps = traj.get("steps") or []
        if not steps:
            continue
        first_step = steps[0]
        state = dict(first_step.get("state_record") or {})
        sample = {
            "text": str(state.get("text", "")),
            "meta": dict(state.get("meta") or {}),
        }
        samples.append(sample)
        labels.append(str(first_step.get("action", "")))
        groups.append(str(traj.get("source_pdf", "")))
    return samples, labels, groups


def _choose_trajectory_split(labels: list[str], groups: list[str]) -> tuple[list[int], list[int], bool, str]:
    """优先按 PDF 分组评估；单来源小样本时退化为确定性 holdout。"""
    unique_groups = list(dict.fromkeys(groups))
    if len(unique_groups) >= 2:
        train_idx, test_idx, full_label_coverage = _choose_grouped_split(labels, groups)
        split_strategy = "grouped_by_pdf" if full_label_coverage else "grouped_by_pdf_partial_labels"
        return train_idx, test_idx, full_label_coverage, split_strategy

    total = len(labels)
    if total < 2:
        raise SystemExit("轨迹训练样本过少，至少需要两条样本。")

    label_indexes: dict[str, list[int]] = defaultdict(list)
    for idx, label in enumerate(labels):
        label_indexes[label].append(idx)

    test_idx: list[int] = []
    for indexes in label_indexes.values():
        if len(indexes) >= 2:
            test_idx.append(indexes[-1])

    if not test_idx:
        fallback_count = max(1, round(total * 0.2))
        test_idx = list(range(total - fallback_count, total))

    test_idx = sorted(dict.fromkeys(test_idx))
    train_idx = [idx for idx in range(total) if idx not in set(test_idx)]
    if not train_idx or not test_idx:
        raise SystemExit("无法为轨迹训练集构造有效切分。")

    full_label_coverage = {labels[idx] for idx in train_idx} == set(labels)
    return train_idx, test_idx, full_label_coverage, "single_source_holdout"


def train_trajectory_model(dataset_path: Path, model_path: Path, *, rebalance_train: bool = False) -> dict:
    try:
        from sklearn.base import clone
        from sklearn.feature_extraction import DictVectorizer
        from sklearn.feature_extraction.text import HashingVectorizer
        from sklearn.linear_model import SGDClassifier
        from sklearn.metrics import classification_report
        from sklearn.pipeline import FeatureUnion, Pipeline
        from sklearn.svm import LinearSVC
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "缺少训练依赖。请先安装 requirements-ml.txt 再运行轨迹策略训练脚本。"
        ) from exc

    trajectories = _load_trajectories(dataset_path)
    if not trajectories:
        raise SystemExit("轨迹训练集为空，无法训练。")

    samples, labels, groups = _flatten_to_step_samples(trajectories)
    if not samples:
        raise SystemExit("没有有效的训练样本。")

    train_idx, test_idx, full_label_coverage, split_strategy = _choose_trajectory_split(labels, groups)

    feature_pipeline = FeatureUnion(
        transformer_list=[
            (
                "text",
                Pipeline(steps=[
                    ("selector", TextFieldExtractor("text")),
                    ("hashing", HashingVectorizer(
                        analyzer="char", ngram_range=(2, 4),
                        n_features=2**16, alternate_sign=False, norm="l2",
                    )),
                ]),
            ),
            (
                "meta",
                Pipeline(steps=[
                    ("selector", MetaFieldExtractor("meta")),
                    ("vect", DictVectorizer()),
                ]),
            ),
        ]
    )

    X_train = [samples[idx] for idx in train_idx]
    y_train = [labels[idx] for idx in train_idx]
    X_test = [samples[idx] for idx in test_idx]
    y_test = [labels[idx] for idx in test_idx]
    train_distribution_before = dict(Counter(y_train))
    train_distribution_after = dict(train_distribution_before)
    if rebalance_train:
        X_train, y_train, train_distribution_before, train_distribution_after = _rebalance_training_split(X_train, y_train)

    if len(samples) >= 8000:
        classifier = SGDClassifier(
            loss="log_loss",
            class_weight="balanced",
            alpha=1e-6,
            max_iter=2500,
            tol=1e-3,
            early_stopping=True,
            validation_fraction=0.1,
            n_iter_no_change=6,
            random_state=42,
        )
    else:
        classifier = LinearSVC(
            class_weight="balanced",
            max_iter=12000,
            tol=1e-4,
            dual="auto",
            random_state=42,
        )
    model = Pipeline(steps=[
        ("features", feature_pipeline),
        ("classifier", classifier),
    ])
    model.fit(X_train, y_train)
    predictions = model.predict(X_test)
    all_labels = sorted({str(l) for l in labels})
    eval_labels = sorted({str(label) for label in y_test} | {str(label) for label in predictions})
    report = classification_report(
        y_test, predictions,
        labels=eval_labels or all_labels, output_dict=True, zero_division=0,
    )

    final_model = clone(model)
    final_model.fit(samples, labels)
    macro_f1 = report.get("macro avg", {}).get("f1-score", 0.0)
    ready_for_runtime = bool(split_strategy.startswith("grouped_by_pdf") and full_label_coverage and macro_f1 >= 0.55)

    bundle = {
        "version": 1,
        "labels": all_labels,
        "label_field": "action",
        "model": final_model,
        "metrics": report,
        "dataset": str(dataset_path),
        "trajectory_count": len(trajectories),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
        "full_label_coverage": full_label_coverage,
        "ready_for_runtime": ready_for_runtime,
        "classifier": classifier.__class__.__name__,
        "rebalanced_train_split": bool(rebalance_train),
        "train_distribution_before": train_distribution_before,
        "train_distribution_after": train_distribution_after,
    }

    model_path.parent.mkdir(parents=True, exist_ok=True)
    with model_path.open("wb") as fh:
        pickle.dump(bundle, fh)

    metrics_path = model_path.with_suffix(".metrics.json")
    metrics_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "model_path": str(model_path),
        "trajectory_count": len(trajectories),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
        "full_label_coverage": full_label_coverage,
        "macro_f1": macro_f1,
        "ready_for_runtime": ready_for_runtime,
        "classifier": classifier.__class__.__name__,
        "rebalanced_train_split": bool(rebalance_train),
        "train_distribution_before": train_distribution_before,
        "train_distribution_after": train_distribution_after,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="训练多步修复轨迹策略模型")
    parser.add_argument("--dataset", type=Path, default=DEFAULT_DATASET)
    parser.add_argument("--model", type=Path, default=DEFAULT_MODEL)
    parser.add_argument("--rebalance-train", action="store_true", help="对训练集做标签重采样")
    args = parser.parse_args()

    summary = train_trajectory_model(args.dataset, args.model, rebalance_train=bool(args.rebalance_train))
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
