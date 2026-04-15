from __future__ import annotations

import argparse
from itertools import combinations
import json
from pathlib import Path
import random
import sys

from PIL import Image
import torch
from torch import nn
from torch.utils.data import DataLoader, Dataset

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from core.pdf_noise_image_model import (
    _IMAGE_SIZE,
    _META_KEYS,
    PdfNoiseImageNet,
    build_pdf_noise_image_meta,
    load_image_tensor,
)


DEFAULT_DATASET = ROOT / "data" / "datasets" / "pdf_noise_images.jsonl"
DEFAULT_MODEL = ROOT / "data" / "models" / "pdf_noise_image_classifier.pt"


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


def _set_seed(seed: int) -> None:
    random.seed(seed)
    torch.manual_seed(seed)
    if torch.cuda.is_available():
        torch.cuda.manual_seed_all(seed)


def _row_meta_vector(row: dict) -> torch.Tensor:
    meta = row.get("meta_record")
    if not isinstance(meta, dict):
        bbox = row.get("bbox") or [0.0, 0.0, 0.0, 0.0]
        page_size = row.get("page_size") or [1.0, 1.0]
        image_size = row.get("image_size") or [0.0, 0.0]
        meta = build_pdf_noise_image_meta(
            x0=float(bbox[0]),
            y0=float(bbox[1]),
            x1=float(bbox[2]),
            y1=float(bbox[3]),
            page_width=float(page_size[0]),
            page_height=float(page_size[1]),
            image_width=float(image_size[0]),
            image_height=float(image_size[1]),
            page_number=int(row.get("page_number") or 1),
        )
    return torch.tensor([float(meta.get(key, 0.0)) for key in _META_KEYS], dtype=torch.float32)


class PdfNoiseImageDataset(Dataset):
    def __init__(self, rows: list[dict], label_to_index: dict[str, int], image_size: int):
        self.rows = rows
        self.label_to_index = label_to_index
        self.image_size = image_size

    def __len__(self) -> int:
        return len(self.rows)

    def __getitem__(self, index: int):
        row = self.rows[index]
        image_path = Path(str(row["image_path"]))
        image_tensor = load_image_tensor(image_path, image_size=self.image_size)
        meta_tensor = _row_meta_vector(row)
        label = self.label_to_index[str(row["label"])]
        return image_tensor, meta_tensor, label


def _build_optimizer(model: PdfNoiseImageNet, learning_rate: float):
    try:
        return torch.optim.Adam(model.parameters(), lr=learning_rate)
    except Exception:
        return None


def _apply_sgd_step(model: PdfNoiseImageNet, learning_rate: float) -> None:
    with torch.no_grad():
        for parameter in model.parameters():
            if parameter.grad is None:
                continue
            parameter.add_(parameter.grad, alpha=-learning_rate)


def _collect_predictions(
    model: PdfNoiseImageNet,
    loader: DataLoader,
    device: torch.device,
) -> tuple[list[int], list[int]]:
    predictions: list[int] = []
    targets: list[int] = []
    model.eval()
    with torch.no_grad():
        for image_tensor, meta_tensor, label_tensor in loader:
            logits = model(image_tensor.to(device), meta_tensor.to(device))
            preds = logits.argmax(dim=1).detach().cpu().tolist()
            predictions.extend(int(value) for value in preds)
            targets.extend(int(value) for value in label_tensor.tolist())
    return predictions, targets


def train_pdf_noise_image_model(
    dataset_path: Path,
    model_path: Path,
    *,
    image_size: int = _IMAGE_SIZE,
    epochs: int = 8,
    batch_size: int = 64,
    learning_rate: float = 1e-3,
    seed: int = 42,
) -> dict:
    try:
        from sklearn.metrics import classification_report
    except ModuleNotFoundError as exc:
        raise SystemExit(
            "缺少训练依赖。请先安装 requirements-ml.txt 再运行训练脚本。"
        ) from exc

    _set_seed(seed)
    rows = _load_rows(dataset_path)
    if not rows:
        raise SystemExit("训练集为空，无法训练。")

    labels = [str(row["label"]) for row in rows]
    groups = [str(row["source_pdf"]) for row in rows]
    train_idx, test_idx, full_label_coverage = _choose_grouped_split(labels, groups)
    split_strategy = "grouped_by_pdf" if full_label_coverage else "grouped_by_pdf_partial_labels"

    all_sorted_labels = sorted({str(label) for label in labels})
    label_to_index = {label: index for index, label in enumerate(all_sorted_labels)}
    index_to_label = {index: label for label, index in label_to_index.items()}

    train_rows = [rows[idx] for idx in train_idx]
    test_rows = [rows[idx] for idx in test_idx]
    train_dataset = PdfNoiseImageDataset(train_rows, label_to_index, image_size=image_size)
    test_dataset = PdfNoiseImageDataset(test_rows, label_to_index, image_size=image_size)

    train_loader = DataLoader(train_dataset, batch_size=batch_size, shuffle=True, num_workers=0)
    test_loader = DataLoader(test_dataset, batch_size=batch_size, shuffle=False, num_workers=0)

    train_label_counts = {label: 0 for label in all_sorted_labels}
    for row in train_rows:
        train_label_counts[str(row["label"])] += 1
    class_weights = torch.tensor(
        [1.0 / max(1, train_label_counts[label]) for label in all_sorted_labels],
        dtype=torch.float32,
    )

    device = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model = PdfNoiseImageNet(meta_dim=len(_META_KEYS), num_classes=len(all_sorted_labels)).to(device)
    optimizer = _build_optimizer(model, learning_rate)
    criterion = nn.CrossEntropyLoss(weight=class_weights.to(device))

    best_state = None
    best_macro_f1 = -1.0
    best_report: dict | None = None

    for _epoch in range(max(1, epochs)):
        model.train()
        for image_tensor, meta_tensor, label_tensor in train_loader:
            if optimizer is not None:
                optimizer.zero_grad(set_to_none=True)
            else:
                model.zero_grad(set_to_none=True)
            logits = model(image_tensor.to(device), meta_tensor.to(device))
            loss = criterion(logits, label_tensor.to(device))
            loss.backward()
            if optimizer is not None:
                optimizer.step()
            else:
                _apply_sgd_step(model, learning_rate)

        predictions, targets = _collect_predictions(model, test_loader, device)
        target_labels = [index_to_label[value] for value in targets]
        predicted_labels = [index_to_label[value] for value in predictions]
        report = classification_report(
            target_labels,
            predicted_labels,
            labels=all_sorted_labels,
            output_dict=True,
            zero_division=0,
        )
        macro_f1 = float(report.get("macro avg", {}).get("f1-score", 0.0))
        if macro_f1 > best_macro_f1:
            best_macro_f1 = macro_f1
            best_state = {key: value.detach().cpu() for key, value in model.state_dict().items()}
            best_report = report

    if best_state is None or best_report is None:
        raise SystemExit("图片噪声模型训练失败，未产出有效模型。")

    model.load_state_dict(best_state)
    noise_precision = float(best_report.get("noise", {}).get("precision", 0.0))
    noise_recall = float(best_report.get("noise", {}).get("recall", 0.0))
    ready_for_runtime = bool(full_label_coverage and best_macro_f1 >= 0.9 and noise_precision >= 0.97 and noise_recall >= 0.8)

    bundle = {
        "version": 1,
        "labels": all_sorted_labels,
        "meta_keys": list(_META_KEYS),
        "meta_dim": len(_META_KEYS),
        "image_size": image_size,
        "state_dict": best_state,
        "metrics": best_report,
        "dataset": str(dataset_path),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
        "full_label_coverage": full_label_coverage,
        "ready_for_runtime": ready_for_runtime,
        "device": str(device),
        "epochs": epochs,
        "batch_size": batch_size,
        "learning_rate": learning_rate,
    }
    model_path.parent.mkdir(parents=True, exist_ok=True)
    torch.save(bundle, model_path)

    metrics_path = model_path.with_suffix(".metrics.json")
    metrics_path.write_text(json.dumps(best_report, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "model_path": str(model_path),
        "metrics_path": str(metrics_path),
        "train_size": len(train_idx),
        "test_size": len(test_idx),
        "split_strategy": split_strategy,
        "full_label_coverage": full_label_coverage,
        "macro_f1": best_macro_f1,
        "noise_precision": noise_precision,
        "noise_recall": noise_recall,
        "ready_for_runtime": ready_for_runtime,
        "device": str(device),
    }


def main() -> None:
    parser = argparse.ArgumentParser(description="训练 PDF 图片噪声分类模型")
    parser.add_argument("--dataset", type=Path, default=DEFAULT_DATASET)
    parser.add_argument("--model", type=Path, default=DEFAULT_MODEL)
    parser.add_argument("--image-size", type=int, default=_IMAGE_SIZE)
    parser.add_argument("--epochs", type=int, default=8)
    parser.add_argument("--batch-size", type=int, default=64)
    parser.add_argument("--learning-rate", type=float, default=1e-3)
    parser.add_argument("--seed", type=int, default=42)
    args = parser.parse_args()

    summary = train_pdf_noise_image_model(
        args.dataset,
        args.model,
        image_size=args.image_size,
        epochs=args.epochs,
        batch_size=args.batch_size,
        learning_rate=args.learning_rate,
        seed=args.seed,
    )
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
