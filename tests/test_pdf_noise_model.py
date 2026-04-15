import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

import core.pdf_noise_model as pdf_noise_model
from core.pdf_noise_model import (
    PdfNoisePrediction,
    build_pdf_noise_feature_record,
    predict_pdf_noise_distribution,
)


class DummyPdfNoiseModel:
    def predict_proba(self, rows):
        row = rows[0]
        text = row["text"]
        if "资料领取" in text or "微信" in text:
            return [[0.02, 0.98]]
        return [[0.95, 0.05]]


class DummyDecisionPdfNoiseModel:
    class _DecisionVector(list):
        ndim = 1

    def decision_function(self, rows):
        row = rows[0]
        text = row["text"]
        if "资料领取" in text or "微信" in text:
            return self._DecisionVector([3.8])
        return self._DecisionVector([-2.9])


class PdfNoiseModelTest(unittest.TestCase):
    def tearDown(self):
        pdf_noise_model._CACHE_KEY = None
        pdf_noise_model._CACHE_BUNDLE = None

    def test_build_pdf_noise_feature_record_includes_text_and_meta(self):
        record = build_pdf_noise_feature_record(
            "各种考试资料领取请加微信",
            x0=40,
            y0=18,
            x1=240,
            y1=32,
            page_width=600,
            page_height=800,
            page_number=2,
            line_index_in_block=1,
            line_count_in_block=3,
        )

        self.assertIn("[TEXT]", record["text"])
        self.assertEqual(record["meta"]["page_number"], 2.0)
        self.assertEqual(record["meta"]["contains_ad_keyword"], 1.0)
        self.assertEqual(record["meta"]["top_band"], 1.0)

    def test_predict_pdf_noise_distribution_loads_ready_default_bundle_without_enable_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "pdf_noise_text_classifier.pkl"
            bundle = {
                "labels": ["content", "noise"],
                "model": DummyPdfNoiseModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with (
                patch.object(pdf_noise_model, "_DEFAULT_MODEL_PATH", model_path),
                patch.object(pdf_noise_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_PICKLED_PDF_NOISE_MODEL": "",
                        "PPTCONVERT_PDF_NOISE_MODEL": "",
                        "PPTCONVERT_TRUST_PDF_NOISE_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                prediction = predict_pdf_noise_distribution(
                    "各种考试资料领取请加微信",
                    x0=36,
                    y0=20,
                    x1=240,
                    y1=36,
                    page_width=600,
                    page_height=800,
                )

        self.assertIsInstance(prediction, PdfNoisePrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_label, "noise")
        self.assertGreater(prediction.best_confidence, 0.9)

    def test_predict_pdf_noise_distribution_supports_binary_decision_function_bundle(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "pdf_noise_text_classifier.pkl"
            bundle = {
                "labels": ["content", "noise"],
                "model": DummyDecisionPdfNoiseModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_PDF_NOISE_MODEL": "1",
                    "PPTCONVERT_TRUST_PDF_NOISE_MODEL": "1",
                },
                clear=False,
            ):
                prediction = predict_pdf_noise_distribution(
                    "各种考试资料领取请加微信",
                    x0=36,
                    y0=20,
                    x1=240,
                    y1=36,
                    page_width=600,
                    page_height=800,
                    model_path=model_path,
                )

        self.assertIsInstance(prediction, PdfNoisePrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_label, "noise")
        self.assertGreater(prediction.best_confidence, 0.9)
