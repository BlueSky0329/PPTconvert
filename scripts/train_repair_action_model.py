from __future__ import annotations

import argparse
import json
import pickle
from collections import defaultdict
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.learned_subject_model import MetaFieldExtractor, TextFieldExtractor
from scripts.train_local_subject_model import _choose_grouped_split


DEFAULT_DATASET = ROOT / "data" / "datasets" / "repair_actions.jsonl"
DEFAULT_MODEL = ROOT / "data" / "models" / "repair_action_classifier.pkl"
DEFAULT_FAMILY_MODEL = ROOT / "data" / "models" / "repair_action_family_classifier.pkl"


def _load_rows(path: Path) -> list[dict]:
    rows: list[dict] = []
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            rows.append(json.loads(line))
    return rows


def _sample_from_row(row: dict) -> dict:
    state = dict(row.get("state_record") or {})
    return {
        "text": str(state.get("text", "")),
        "meta": dict(state.get("meta") or {}),
    }


def _label_from_row(row: dict, label_field: str) -> str:
    return str(row.get(label_field) or "")


def _rebalance_training_split(samples: list[dict], labels: list[str]) -> tuple[list[dict], list[str], dict[str, int], dict[str, int]]:
    grouped: dict[str, list[int]] = defaultdict(list)
    for index, label in enumerate(labels):
        grouped[label].append(index)

    before = {label: len(indexes) for label, indexes in grouped.items()}
    if not grouped:
        return samples, labels, before, before

    majority_count = max(before.values())
    majority_cap = min(1500, majority_count)
    target_counts: dict[str, int] = {}
    for label, count in before.items():
        target_counts[label] = min(majority_cap, max(count, max(96, count * 12)))

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


def train_repair_action_model(dataset_path: Path, model_path: Path, *, label_field: str = "action") -> dict:
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
            "缺少训练依赖。请先安装 requirements-ml.txt 再运行修复动作训练脚本。"
        ) from exc

    rows = _load_rows(dataset_path)
    if not rows:
        raise SystemExit("修复动作训练集为空，无法训练。")

    labels = [_label_from_row(row, label_field) for row in rows]
    if not all(labels):
        raise SystemExit(f"修复动作训练集缺少标签字段：{label_field}")
    groups = [str(row["source_pdf"]) for row in rows]
    samples = [_sample_from_row(row) for row in rows]

    train_idx, test_idx, full_label_coverage = _choose_grouped_split(labels, groups)
    split_strategy = "grouped_by_pdf" if full_label_coverage else "grouped_by_pdf_partial_labels"

    feature_pipeline = FeatureUnion(
        transformer_list=[
            (
                "text",
                Pipeline(
                    steps=[
                        ("selector", TextFieldExtractor("text")),
                        (
                            "hashing",
                            HashingVectorizer(
                                analyzer="char",
                                ngram_range=(2, 4),
                                n_features=2**17,
                                alternate_sign=False,
                                norm="l2",
                            ),
                        ),
                    ]
                ),
            ),
            (
                "meta",
                Pipeline(
                    steps=[
                        ("selector", MetaFieldExtractor("meta")),
                        ("vect", DictVectorizer()),
                    ]
                ),
            ),
        ]
    )

    X_train = [samples[idx] for idx in train_idx]
    y_train = [labels[idx] for idx in train_idx]
    X_test = [samples[idx] for idx in test_idx]
    y_test = [labels[idx] for idx in test_idx]
    X_train, y_train, train_distribution_before, train_distribution_after = _rebalance_training_split(X_train, y_train)

    classifier = LinearSVC(
        class_weight="balanced",
        max_iter=30000,
        tol=1e-4,
        random_state=42,
    )

    model = Pipeline(
        steps=[
            ("features", feature_pipeline),
            ("classifier", classifier),
        ]
    )
    model.fit(X_train, y_train)
    predictions = model.predict(X_test)
    all_sorted_labels = sorted({str(label) for label in labels})
    report = classification_report(
        y_test,
        predictions,
        labels=all_sorted_labels,
        output_dict=True,
        zero_division=0,
    )

    final_model = clone(model)
    final_model.fit(samples, labels)
    macro_f1 = report.get("macro avg", {}).get("f1-score", 0.0)
    ready_for_runtime = bool(full_label_coverage and macro_f1 >= 0.55)

    bundle = {
        "version": 1,
        "labels": all_sorted_labels,
        "label_field": label_field,
        "model": final_model,
        "metrics": report,
        "dataset": str(dataset_path),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
        "full_label_coverage": full_label_coverage,
        "ready_for_runtime": ready_for_runtime,
    }

    model_path.parent.mkdir(parents=True, exist_ok=True)
    with model_path.open("wb") as fh:
        pickle.dump(bundle, fh)

    metrics_path = model_path.with_suffix(".metrics.json")
    metrics_path.write_text(json.dumps(report, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "model_path": str(model_path),
        "metrics_path": str(metrics_path),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
        "full_label_coverage": full_label_coverage,
        "macro_f1": macro_f1,
        "ready_for_runtime": ready_for_runtime,
        "label_count": len(all_sorted_labels),
        "label_field": label_field,
        "train_distribution_before": train_distribution_before,
        "train_distribution_after": train_distribution_after,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="训练本地公务员题目修复动作模型")
    parser.add_argument("--dataset", type=Path, default=DEFAULT_DATASET)
    parser.add_argument("--model", type=Path)
    parser.add_argument("--label-field", choices=["action", "action_family"], default="action")
    args = parser.parse_args()

    model_path = args.model or (DEFAULT_FAMILY_MODEL if args.label_field == "action_family" else DEFAULT_MODEL)
    summary = train_repair_action_model(args.dataset, model_path, label_field=args.label_field)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
