import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

from PIL import Image, ImageDraw
import torch

import core.pdf_noise_image_model as pdf_noise_image_model
from core.pdf_noise_image_model import (
    PdfNoiseImagePrediction,
    build_pdf_noise_image_meta,
    build_pdf_noise_image_visual_stats,
    is_background_like_pdf_image_meta,
    is_watermark_like_pdf_image_meta,
    predict_pdf_noise_image_distribution,
)


class DummyPdfNoiseImageNet(torch.nn.Module):
    def __init__(self, *, meta_dim: int, num_classes: int) -> None:
        super().__init__()
        self.meta_dim = meta_dim
        self.num_classes = num_classes
        self.linear = torch.nn.Linear(meta_dim, num_classes)

    def forward(self, image_tensor: torch.Tensor, meta_tensor: torch.Tensor) -> torch.Tensor:
        scores = self.linear(meta_tensor)
        top_band = meta_tensor[:, 7]
        scores[:, 1] = scores[:, 1] + top_band * 6.0
        scores[:, 0] = scores[:, 0] + (1.0 - top_band) * 3.0
        return scores


class PdfNoiseImageModelTest(unittest.TestCase):
    def tearDown(self):
        pdf_noise_image_model._CACHE_KEY = None
        pdf_noise_image_model._CACHE_BUNDLE = None
        pdf_noise_image_model._CACHE_MODEL = None
        pdf_noise_image_model._CACHE_DEVICE = None

    def test_build_pdf_noise_image_meta_marks_banner_like_top_image(self):
        meta = build_pdf_noise_image_meta(
            x0=24,
            y0=8,
            x1=560,
            y1=72,
            page_width=596,
            page_height=842,
            image_width=536,
            image_height=64,
            page_number=1,
        )

        self.assertEqual(meta["top_band"], 1.0)
        self.assertEqual(meta["banner_like"], 1.0)

    def test_visual_stats_detect_uniform_background(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "blank.png"
            Image.new("RGB", (128, 128), color="black").save(image_path)

            stats = build_pdf_noise_image_visual_stats(image_path)
            meta = build_pdf_noise_image_meta(
                x0=0,
                y0=92,
                x1=596,
                y1=782,
                page_width=596,
                page_height=842,
                image_width=128,
                image_height=128,
                page_number=1,
                **stats,
            )

        self.assertLessEqual(stats["gray_stddev"], 0.01)
        self.assertTrue(is_background_like_pdf_image_meta(meta))

    def test_visual_stats_detect_faint_watermark(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "watermark.png"
            image = Image.new("RGB", (256, 256), color=(248, 248, 248))
            draw = ImageDraw.Draw(image)
            for offset in range(-80, 280, 28):
                draw.line((offset, 0, offset + 120, 256), fill=(212, 212, 212), width=6)
            image.save(image_path)

            stats = build_pdf_noise_image_visual_stats(image_path)
            meta = build_pdf_noise_image_meta(
                x0=4,
                y0=88,
                x1=592,
                y1=784,
                page_width=596,
                page_height=842,
                image_width=256,
                image_height=256,
                page_number=1,
                **stats,
            )

        self.assertGreaterEqual(stats["bright_ratio"], 0.72)
        self.assertLessEqual(stats["dark_ratio"], 0.025)
        self.assertTrue(is_watermark_like_pdf_image_meta(meta))

    def test_predict_pdf_noise_image_distribution_loads_ready_default_bundle_without_enable_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "pdf_noise_image_classifier.pt"
            image_path = Path(tmp) / "candidate.png"
            Image.new("RGB", (96, 32), color="white").save(image_path)

            model = DummyPdfNoiseImageNet(meta_dim=len(pdf_noise_image_model._META_KEYS), num_classes=2)
            bundle = {
                "labels": ["content", "noise"],
                "meta_dim": len(pdf_noise_image_model._META_KEYS),
                "image_size": 96,
                "state_dict": model.state_dict(),
                "ready_for_runtime": True,
            }
            torch.save(bundle, model_path)

            with (
                patch.object(pdf_noise_image_model, "_DEFAULT_MODEL_PATH", model_path),
                patch.object(pdf_noise_image_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.object(pdf_noise_image_model, "PdfNoiseImageNet", DummyPdfNoiseImageNet),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_PDF_NOISE_IMAGE_MODEL": "",
                        "PPTCONVERT_PDF_NOISE_IMAGE_MODEL": "",
                        "PPTCONVERT_TRUST_PDF_NOISE_IMAGE_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                prediction = predict_pdf_noise_image_distribution(
                    image_path,
                    x0=24,
                    y0=8,
                    x1=560,
                    y1=72,
                    page_width=596,
                    page_height=842,
                    page_number=1,
                )

        self.assertIsInstance(prediction, PdfNoiseImagePrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_label, "noise")
        self.assertGreater(prediction.best_confidence, 0.8)
