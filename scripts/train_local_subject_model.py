from __future__ import annotations

import argparse
import json
import pickle
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.learned_subject_model import MetaFieldExtractor, TextFieldExtractor


DEFAULT_DATASET = ROOT / "data" / "datasets" / "subject_gold.jsonl"
DEFAULT_MODEL = ROOT / "data" / "models" / "subject_classifier.pkl"


def _load_rows(path: Path) -> list[dict]:
    rows: list[dict] = []
    with path.open("r", encoding="utf-8") as fh:
        for line in fh:
            line = line.strip()
            if not line:
                continue
            rows.append(json.loads(line))
    return rows


def _covers_all_labels(indexes, labels) -> bool:
    return {labels[idx] for idx in indexes} == set(labels)


def train_subject_model(dataset_path: Path, model_path: Path) -> dict:
    try:
        from sklearn.base import clone
        from sklearn.feature_extraction import DictVectorizer
        from sklearn.feature_extraction.text import TfidfVectorizer
        from sklearn.metrics import classification_report
        from sklearn.model_selection import GroupShuffleSplit, StratifiedShuffleSplit
        from sklearn.pipeline import FeatureUnion, Pipeline
        from sklearn.svm import LinearSVC
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "缺少训练依赖。请先安装 requirements-ml.txt 再运行训练脚本。"
        ) from exc

    rows = _load_rows(dataset_path)
    if not rows:
        raise SystemExit("训练集为空，无法训练。")

    labels = [row["subject"] for row in rows]
    groups = [row["source_pdf"] for row in rows]

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

    samples = [row["feature_record"] for row in rows]
    split_strategy = "grouped_by_pdf"
    splitter = GroupShuffleSplit(n_splits=1, test_size=0.2, random_state=42)
    train_idx, test_idx = next(splitter.split(samples, labels, groups))
    if not _covers_all_labels(train_idx, labels) or not _covers_all_labels(test_idx, labels):
        split_strategy = "stratified_by_question"
        stratified = StratifiedShuffleSplit(n_splits=1, test_size=0.2, random_state=42)
        train_idx, test_idx = next(stratified.split(samples, labels))

    X_train = [samples[idx] for idx in train_idx]
    y_train = [labels[idx] for idx in train_idx]
    X_test = [samples[idx] for idx in test_idx]
    y_test = [labels[idx] for idx in test_idx]

    model = Pipeline(
        steps=[
            ("features", feature_pipeline),
            (
                "classifier",
                LinearSVC(
                    class_weight="balanced",
                    dual="auto",
                    max_iter=5000,
                ),
            ),
        ]
    )
    model.fit(X_train, y_train)
    predictions = model.predict(X_test)
    report = classification_report(y_test, predictions, output_dict=True, zero_division=0)
    final_model = clone(model)
    final_model.fit(samples, labels)
    macro_f1 = report.get("macro avg", {}).get("f1-score", 0.0)
    ready_for_runtime = bool(macro_f1 >= 0.45)

    bundle = {
        "version": 1,
        "labels": sorted({str(label) for label in labels}),
        "model": final_model,
        "metrics": report,
        "dataset": str(dataset_path),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
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
        "macro_f1": macro_f1,
        "ready_for_runtime": ready_for_runtime,
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="训练本地公务员题目科目分类模型")
    parser.add_argument("--dataset", type=Path, default=DEFAULT_DATASET)
    parser.add_argument("--model", type=Path, default=DEFAULT_MODEL)
    args = parser.parse_args()

    summary = train_subject_model(args.dataset, args.model)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
