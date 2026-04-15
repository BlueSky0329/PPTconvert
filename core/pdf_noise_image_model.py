from __future__ import annotations

import os
from dataclasses import dataclass
from pathlib import Path
from typing import Any

from PIL import Image, ImageFilter, ImageOps, ImageStat
import torch
from torch import nn


_ROOT = Path(__file__).resolve().parent.parent
_DEFAULT_MODEL_PATH = _ROOT / "data" / "models" / "pdf_noise_image_classifier.pt"
_TRUSTED_MODEL_DIR = _DEFAULT_MODEL_PATH.parent.resolve()
_CACHE_KEY: tuple[str, float] | None = None
_CACHE_BUNDLE: dict[str, Any] | None = None
_CACHE_MODEL: nn.Module | None = None
_CACHE_DEVICE: str | None = None
_ENABLE_MODEL_ENV = "PPTCONVERT_ENABLE_PDF_NOISE_IMAGE_MODEL"
_FORCE_MODEL_ENV = "PPTCONVERT_FORCE_PDF_NOISE_IMAGE_MODEL"
_TRUST_MODEL_ENV = "PPTCONVERT_TRUST_PDF_NOISE_IMAGE_MODEL"
_MODEL_PATH_ENV = "PPTCONVERT_PDF_NOISE_IMAGE_MODEL"
_IMAGE_SIZE = 96
_META_KEYS: tuple[str, ...] = (
    "page_number",
    "page_top_ratio",
    "page_bottom_ratio",
    "width_ratio",
    "height_ratio",
    "image_aspect_ratio",
    "center_offset_ratio",
    "top_band",
    "bottom_band",
    "left_edge",
    "right_edge",
    "corner_like",
    "banner_like",
    "small_image",
    "page_cover_like",
    "gray_mean",
    "gray_stddev",
    "gray_range",
    "edge_density",
    "bright_ratio",
    "dark_ratio",
)


@dataclass(frozen=True)
class PdfNoiseImagePrediction:
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


class PdfNoiseImageNet(nn.Module):
    def __init__(self, *, meta_dim: int, num_classes: int) -> None:
        super().__init__()
        self.image_encoder = nn.Sequential(
            nn.Conv2d(3, 24, kernel_size=3, stride=1, padding=1),
            nn.ReLU(inplace=True),
            nn.MaxPool2d(2),
            nn.Conv2d(24, 48, kernel_size=3, stride=1, padding=1),
            nn.ReLU(inplace=True),
            nn.MaxPool2d(2),
            nn.Conv2d(48, 96, kernel_size=3, stride=1, padding=1),
            nn.ReLU(inplace=True),
            nn.MaxPool2d(2),
            nn.AdaptiveAvgPool2d((1, 1)),
        )
        self.meta_encoder = nn.Sequential(
            nn.Linear(meta_dim, 32),
            nn.ReLU(inplace=True),
            nn.Linear(32, 32),
            nn.ReLU(inplace=True),
        )
        self.classifier = nn.Sequential(
            nn.Linear(96 + 32, 64),
            nn.ReLU(inplace=True),
            nn.Dropout(0.1),
            nn.Linear(64, num_classes),
        )

    def forward(self, image_tensor: torch.Tensor, meta_tensor: torch.Tensor) -> torch.Tensor:
        image_features = self.image_encoder(image_tensor).flatten(1)
        meta_features = self.meta_encoder(meta_tensor)
        combined = torch.cat([image_features, meta_features], dim=1)
        return self.classifier(combined)


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


def default_model_path() -> Path:
    configured = os.environ.get(_MODEL_PATH_ENV, "").strip()
    return Path(configured) if configured else _DEFAULT_MODEL_PATH


def _choose_device() -> torch.device:
    if torch.cuda.is_available():
        return torch.device("cuda")
    return torch.device("cpu")


def _safe_ratio(numerator: float, denominator: float) -> float:
    if denominator <= 0:
        return 0.0
    return float(numerator) / float(denominator)


def build_pdf_noise_image_visual_stats(image: Image.Image | str | os.PathLike[str]) -> dict[str, float]:
    if isinstance(image, Image.Image):
        rgb = image.convert("RGB")
    else:
        with Image.open(image) as loaded:
            return build_pdf_noise_image_visual_stats(loaded)

    gray = rgb.convert("L")
    stat = ImageStat.Stat(gray)
    histogram = gray.histogram()
    total_pixels = max(1, int(sum(histogram)))
    minimum, maximum = gray.getextrema()
    edge_stat = ImageStat.Stat(gray.filter(ImageFilter.FIND_EDGES))
    return {
        "gray_mean": _safe_ratio(float(stat.mean[0]), 255.0),
        "gray_stddev": _safe_ratio(float(stat.stddev[0]), 255.0),
        "gray_range": _safe_ratio(float(maximum - minimum), 255.0),
        "edge_density": _safe_ratio(float(edge_stat.mean[0]), 255.0),
        "bright_ratio": _safe_ratio(float(sum(histogram[235:])), float(total_pixels)),
        "dark_ratio": _safe_ratio(float(sum(histogram[:20])), float(total_pixels)),
    }


def build_pdf_noise_image_meta(
    *,
    x0: float = 0.0,
    y0: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    page_width: float = 0.0,
    page_height: float = 0.0,
    image_width: float = 0.0,
    image_height: float = 0.0,
    page_number: int = 1,
    gray_mean: float = 0.0,
    gray_stddev: float = 0.0,
    gray_range: float = 0.0,
    edge_density: float = 0.0,
    bright_ratio: float = 0.0,
    dark_ratio: float = 0.0,
) -> dict[str, float]:
    width = max(0.0, float(x1) - float(x0))
    height = max(0.0, float(y1) - float(y0))
    center_x = (float(x0) + float(x1)) / 2.0
    center_offset_ratio = _safe_ratio(abs(center_x - page_width / 2.0), page_width)
    width_ratio = _safe_ratio(width, page_width)
    height_ratio = _safe_ratio(height, page_height)
    page_top_ratio = _safe_ratio(y0, page_height)
    page_bottom_ratio = _safe_ratio(y1, page_height)
    image_aspect_ratio = _safe_ratio(image_width, image_height) if image_height > 0 else 0.0
    left_edge = 1.0 if _safe_ratio(x0, page_width) <= 0.2 else 0.0
    right_edge = 1.0 if _safe_ratio(x1, page_width) >= 0.8 else 0.0
    top_band = 1.0 if page_top_ratio <= 0.16 else 0.0
    bottom_band = 1.0 if page_bottom_ratio >= 0.84 else 0.0
    corner_like = 1.0 if (top_band or bottom_band) and (left_edge or right_edge) else 0.0
    banner_like = 1.0 if page_top_ratio <= 0.18 and width_ratio >= 0.55 and height_ratio <= 0.2 else 0.0
    small_image = 1.0 if width_ratio <= 0.18 and height_ratio <= 0.14 else 0.0
    page_cover_like = 1.0 if width_ratio >= 0.72 and height_ratio >= 0.55 and center_offset_ratio <= 0.16 else 0.0
    return {
        "page_number": float(page_number),
        "page_top_ratio": page_top_ratio,
        "page_bottom_ratio": page_bottom_ratio,
        "width_ratio": width_ratio,
        "height_ratio": height_ratio,
        "image_aspect_ratio": float(image_aspect_ratio),
        "center_offset_ratio": center_offset_ratio,
        "top_band": top_band,
        "bottom_band": bottom_band,
        "left_edge": left_edge,
        "right_edge": right_edge,
        "corner_like": corner_like,
        "banner_like": banner_like,
        "small_image": small_image,
        "page_cover_like": page_cover_like,
        "gray_mean": float(gray_mean),
        "gray_stddev": float(gray_stddev),
        "gray_range": float(gray_range),
        "edge_density": float(edge_density),
        "bright_ratio": float(bright_ratio),
        "dark_ratio": float(dark_ratio),
    }


def build_pdf_noise_image_meta_vector(
    *,
    keys: tuple[str, ...] | list[str] | None = None,
    **kwargs: float | int,
) -> torch.Tensor:
    meta = build_pdf_noise_image_meta(**kwargs)
    feature_keys = tuple(keys or _META_KEYS)
    return torch.tensor([float(meta.get(key, 0.0)) for key in feature_keys], dtype=torch.float32)


def is_background_like_pdf_image_meta(meta: dict[str, float]) -> bool:
    page_cover_like = float(meta.get("page_cover_like", 0.0)) >= 1.0
    top_ratio = float(meta.get("page_top_ratio", 0.0))
    bottom_ratio = float(meta.get("page_bottom_ratio", 0.0))
    gray_mean = float(meta.get("gray_mean", 0.0))
    gray_stddev = float(meta.get("gray_stddev", 0.0))
    gray_range = float(meta.get("gray_range", 0.0))
    edge_density = float(meta.get("edge_density", 0.0))
    extreme_tone = gray_mean <= 0.12 or gray_mean >= 0.88
    low_variance = gray_stddev <= 0.06 and gray_range <= 0.12 and edge_density <= 0.08
    margin_spanning = top_ratio <= 0.18 and bottom_ratio >= 0.82
    return page_cover_like and margin_spanning and extreme_tone and low_variance


def is_watermark_like_pdf_image_meta(meta: dict[str, float]) -> bool:
    page_cover_like = float(meta.get("page_cover_like", 0.0)) >= 1.0
    top_ratio = float(meta.get("page_top_ratio", 0.0))
    bottom_ratio = float(meta.get("page_bottom_ratio", 0.0))
    gray_mean = float(meta.get("gray_mean", 0.0))
    gray_stddev = float(meta.get("gray_stddev", 0.0))
    gray_range = float(meta.get("gray_range", 0.0))
    edge_density = float(meta.get("edge_density", 0.0))
    bright_ratio = float(meta.get("bright_ratio", 0.0))
    dark_ratio = float(meta.get("dark_ratio", 0.0))
    margin_spanning = top_ratio <= 0.18 and bottom_ratio >= 0.82
    very_bright = gray_mean >= 0.86 and bright_ratio >= 0.72
    faint_contrast = gray_stddev <= 0.08 and gray_range <= 0.28 and edge_density <= 0.05
    sparse_dark = dark_ratio <= 0.025
    return page_cover_like and margin_spanning and very_bright and faint_contrast and sparse_dark


def load_image_tensor(image_path: str | os.PathLike[str], image_size: int = _IMAGE_SIZE) -> torch.Tensor:
    with Image.open(image_path) as image:
        fitted = ImageOps.fit(
            image.convert("RGB"),
            (image_size, image_size),
            method=Image.Resampling.BILINEAR,
        )
        byte_tensor = torch.frombuffer(bytearray(fitted.tobytes()), dtype=torch.uint8)
        tensor = byte_tensor.view(image_size, image_size, 3).permute(2, 0, 1).float() / 255.0
    return tensor


def _load_bundle(model_path: Path | None = None) -> tuple[dict[str, Any] | None, nn.Module | None, torch.device]:
    global _CACHE_BUNDLE, _CACHE_KEY, _CACHE_MODEL, _CACHE_DEVICE
    path = Path(model_path or default_model_path())
    device = _choose_device()
    if not path.exists():
        return None, None, device
    if not _model_loading_enabled(path):
        return None, None, device
    if not _is_trusted_model_path(path):
        return None, None, device
    key = (str(path.resolve()), path.stat().st_mtime)
    if _CACHE_KEY == key and _CACHE_BUNDLE is not None and _CACHE_MODEL is not None and _CACHE_DEVICE == str(device):
        return _CACHE_BUNDLE, _CACHE_MODEL, device
    try:
        bundle = torch.load(path, map_location=device)
    except Exception:
        return None, None, device
    if not isinstance(bundle, dict) or "state_dict" not in bundle or "labels" not in bundle:
        return None, None, device
    if bundle.get("ready_for_runtime") is False and os.environ.get(_FORCE_MODEL_ENV, "").strip() != "1":
        return None, None, device
    labels = [str(label) for label in bundle.get("labels", [])]
    if not labels:
        return None, None, device
    model = PdfNoiseImageNet(
        meta_dim=int(bundle.get("meta_dim") or len(_META_KEYS)),
        num_classes=len(labels),
    )
    try:
        model.load_state_dict(bundle["state_dict"])
    except Exception:
        return None, None, device
    model.to(device)
    model.eval()
    _CACHE_KEY = key
    _CACHE_BUNDLE = bundle
    _CACHE_MODEL = model
    _CACHE_DEVICE = str(device)
    return bundle, model, device


def predict_pdf_noise_image_distribution(
    image_path: str | os.PathLike[str],
    *,
    x0: float = 0.0,
    y0: float = 0.0,
    x1: float = 0.0,
    y1: float = 0.0,
    page_width: float = 0.0,
    page_height: float = 0.0,
    page_number: int = 1,
    model_path: str | os.PathLike[str] | None = None,
) -> PdfNoiseImagePrediction | None:
    bundle, model, device = _load_bundle(Path(model_path) if model_path else None)
    if bundle is None or model is None:
        return None
    labels = [str(label) for label in bundle.get("labels", [])]
    meta_keys = [str(key) for key in (bundle.get("meta_keys") or list(_META_KEYS))]
    image_size = int(bundle.get("image_size") or _IMAGE_SIZE)
    try:
        with Image.open(image_path) as image:
            image_width, image_height = image.size
            visual_stats = build_pdf_noise_image_visual_stats(image)
        image_tensor = load_image_tensor(image_path, image_size=image_size).unsqueeze(0).to(device)
        meta_tensor = build_pdf_noise_image_meta_vector(
            keys=meta_keys,
            x0=x0,
            y0=y0,
            x1=x1,
            y1=y1,
            page_width=page_width,
            page_height=page_height,
            image_width=image_width,
            image_height=image_height,
            page_number=page_number,
            gray_mean=float(visual_stats.get("gray_mean", 0.0)),
            gray_stddev=float(visual_stats.get("gray_stddev", 0.0)),
            gray_range=float(visual_stats.get("gray_range", 0.0)),
            edge_density=float(visual_stats.get("edge_density", 0.0)),
            bright_ratio=float(visual_stats.get("bright_ratio", 0.0)),
            dark_ratio=float(visual_stats.get("dark_ratio", 0.0)),
        ).unsqueeze(0).to(device)
    except Exception:
        return None
    with torch.no_grad():
        logits = model(image_tensor, meta_tensor)
        probs = torch.softmax(logits, dim=1)[0].detach().cpu().tolist()
    if len(probs) != len(labels):
        return None
    probabilities = {
        label: max(0.0, min(1.0, float(prob)))
        for label, prob in zip(labels, probs)
    }
    return PdfNoiseImagePrediction(probabilities=probabilities)
