import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

import core.learned_subject_model as learned_subject_model
from core.learned_subject_model import (
    LearnedSubjectPrediction,
    build_subject_feature_record,
    predict_subject_distribution,
)


class DummySubjectModel:
    def predict_proba(self, rows):
        row = rows[0]
        text = row["text"]
        if "根据上述定义" in text:
            return [[0.05, 0.85, 0.10]]
        return [[0.70, 0.10, 0.20]]


class DummyDecisionSubjectModel:
    class _DecisionMatrix(list):
        ndim = 2

    def decision_function(self, rows):
        row = rows[0]
        text = row["text"]
        if "根据上述定义" in text:
            return self._DecisionMatrix([[-0.8, 2.6, 0.1]])
        return self._DecisionMatrix([[2.2, -1.0, 0.2]])


class LearnedSubjectModelTest(unittest.TestCase):
    def tearDown(self):
        learned_subject_model._CACHE_KEY = None
        learned_subject_model._CACHE_BUNDLE = None

    def test_build_subject_feature_record_includes_text_and_meta(self):
        record = build_subject_feature_record(
            stem="根据上述定义，下列符合定义的是",
            options=["甲", "乙", "丙", "丁"],
            material_text="",
            image_count=4,
            material_header="",
        )

        self.assertIn("[STEM]", record["text"])
        self.assertEqual(record["meta"]["option_count"], 4.0)
        self.assertEqual(record["meta"]["image_count"], 4.0)

    def test_predict_subject_distribution_loads_pickled_bundle(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummySubjectModel(),
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "1",
                    "PPTCONVERT_TRUST_SUBJECT_MODEL": "1",
                },
                clear=False,
            ):
                prediction = predict_subject_distribution(
                    stem="根据上述定义，下列符合定义的是",
                    options=["甲", "乙", "丙", "丁"],
                    model_path=model_path,
                )

        self.assertIsInstance(prediction, LearnedSubjectPrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_kind, "reasoning")
        self.assertGreater(prediction.best_confidence, 0.8)

    def test_predict_subject_distribution_loads_ready_default_bundle_without_enable_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummyDecisionSubjectModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with (
                patch.object(learned_subject_model, "_DEFAULT_MODEL_PATH", model_path),
                patch.object(learned_subject_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "",
                        "PPTCONVERT_SUBJECT_MODEL": "",
                        "PPTCONVERT_TRUST_SUBJECT_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                prediction = predict_subject_distribution(
                    stem="根据上述定义，下列符合定义的是",
                    options=["甲", "乙", "丙", "丁"],
                )

        self.assertIsInstance(prediction, LearnedSubjectPrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_kind, "reasoning")
        self.assertGreater(prediction.best_confidence, 0.8)

    def test_predict_subject_distribution_can_disable_default_bundle_with_env_flag(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummyDecisionSubjectModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with (
                patch.object(learned_subject_model, "_DEFAULT_MODEL_PATH", model_path),
                patch.object(learned_subject_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "0",
                        "PPTCONVERT_SUBJECT_MODEL": "",
                        "PPTCONVERT_TRUST_SUBJECT_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                prediction = predict_subject_distribution(
                    stem="根据上述定义，下列符合定义的是",
                    options=["甲", "乙", "丙", "丁"],
                )

        self.assertIsNone(prediction)

    def test_predict_subject_distribution_skips_not_ready_bundle(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummySubjectModel(),
                "ready_for_runtime": False,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "1",
                    "PPTCONVERT_TRUST_SUBJECT_MODEL": "1",
                },
                clear=False,
            ):
                prediction = predict_subject_distribution(
                    stem="根据上述定义，下列符合定义的是",
                    options=["甲", "乙", "丙", "丁"],
                    model_path=model_path,
                )

        self.assertIsNone(prediction)

    def test_predict_subject_distribution_supports_decision_function_bundle_via_env_model_path(self):
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummyDecisionSubjectModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "1",
                    "PPTCONVERT_TRUST_SUBJECT_MODEL": "1",
                    "PPTCONVERT_SUBJECT_MODEL": str(model_path),
                },
                clear=False,
            ):
                prediction = predict_subject_distribution(
                    stem="根据上述定义，下列符合定义的是",
                    options=["甲", "乙", "丙", "丁"],
                )

        self.assertIsInstance(prediction, LearnedSubjectPrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_kind, "reasoning")
        self.assertGreater(prediction.best_confidence, 0.8)
