from __future__ import annotations

import argparse
from itertools import combinations
import json
import pickle
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.pdf_noise_model import MetaFieldExtractor, TextFieldExtractor


DEFAULT_DATASET = ROOT / "data" / "datasets" / "pdf_noise_text.jsonl"
DEFAULT_MODEL = ROOT / "data" / "models" / "pdf_noise_text_classifier.pkl"


def _load_rows(path: Path) -> list[dict]:
    rows: list[dict] = []
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            rows.append(json.loads(line))
    return rows


def _choose_grouped_split(labels: list[str], groups: list[str]) -> tuple[list[int], list[int], bool]:
    unique_groups = list(dict.fromkeys(groups))
    if len(unique_groups) < 2:
        raise SystemExit("至少需要两份不同来源 PDF 才能做按 PDF 分组评估。")

    all_labels = set(labels)
    target_ratio = 0.2
    total_rows = len(labels)
    target_test_rows = total_rows * target_ratio
    target_group_count = max(1, round(len(unique_groups) * target_ratio))
    best_choice: tuple[tuple[float, float, float], list[int], list[int], bool] | None = None

    for size in range(1, len(unique_groups)):
        for group_combo in combinations(unique_groups, size):
            test_groups = set(group_combo)
            train_idx = [index for index, group in enumerate(groups) if group not in test_groups]
            test_idx = [index for index, group in enumerate(groups) if group in test_groups]
            if not train_idx or not test_idx:
                continue

            train_labels = {labels[idx] for idx in train_idx}
            test_labels = {labels[idx] for idx in test_idx}
            if len(train_labels) < 2:
                continue

            full_coverage = train_labels == all_labels and test_labels == all_labels
            coverage_score = (len(train_labels) + len(test_labels)) / (2 * max(len(all_labels), 1))
            row_balance = -abs(len(test_idx) - target_test_rows)
            group_balance = -abs(size - target_group_count)
            key = (
                1.0 if full_coverage else 0.0,
                coverage_score,
                row_balance + group_balance * 0.1,
            )
            if best_choice is None or key > best_choice[0]:
                best_choice = (key, train_idx, test_idx, full_coverage)

    if best_choice is None:
        raise SystemExit("无法为当前语料构造有效的按 PDF 分组评估切分。")
    return best_choice[1], best_choice[2], best_choice[3]


def train_pdf_noise_model(dataset_path: Path, model_path: Path) -> dict:
    try:
        from sklearn.base import clone
        from sklearn.feature_extraction import DictVectorizer
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.linear_model import LogisticRegression
        from sklearn.metrics import classification_report
        from sklearn.pipeline import FeatureUnion, Pipeline
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "缺少训练依赖。请先安装 requirements-ml.txt 再运行训练脚本。"
        ) from exc

    rows = _load_rows(dataset_path)
    if not rows:
        raise SystemExit("训练集为空，无法训练。")

    labels = [str(row["label"]) for row in rows]
    groups = [str(row["source_pdf"]) for row in rows]
    samples = [row["feature_record"] for row in rows]

    train_idx, test_idx, full_label_coverage = _choose_grouped_split(labels, groups)
    split_strategy = "grouped_by_pdf" if full_label_coverage else "grouped_by_pdf_partial_labels"

    feature_pipeline = FeatureUnion(
        transformer_list=[
            (
                "text",
                Pipeline(
                    steps=[
                        ("selector", TextFieldExtractor("text")),
                        ("tfidf", TfidfVectorizer(analyzer="char_wb", ngram_range=(2, 5), min_df=2)),
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

    model = Pipeline(
        steps=[
            ("features", feature_pipeline),
            (
                "classifier",
                LogisticRegression(
                    class_weight="balanced",
                    max_iter=2000,
                    solver="liblinear",
                ),
            ),
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
    macro_f1 = float(report.get("macro avg", {}).get("f1-score", 0.0))
    noise_precision = float(report.get("noise", {}).get("precision", 0.0))
    noise_recall = float(report.get("noise", {}).get("recall", 0.0))
    ready_for_runtime = bool(full_label_coverage and macro_f1 >= 0.85 and noise_precision >= 0.98 and noise_recall >= 0.7)

    bundle = {
        "version": 1,
        "labels": all_sorted_labels,
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
        "noise_precision": noise_precision,
        "noise_recall": noise_recall,
        "ready_for_runtime": ready_for_runtime,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="训练 PDF 文本噪声分类模型")
    parser.add_argument("--dataset", type=Path, default=DEFAULT_DATASET)
    parser.add_argument("--model", type=Path, default=DEFAULT_MODEL)
    args = parser.parse_args()

    summary = train_pdf_noise_model(args.dataset, args.model)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
