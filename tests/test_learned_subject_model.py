import pickle
import tempfile
import unittest
from pathlib import Path

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


class LearnedSubjectModelTest(unittest.TestCase):
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

            prediction = predict_subject_distribution(
                stem="根据上述定义，下列符合定义的是",
                options=["甲", "乙", "丙", "丁"],
                model_path=model_path,
            )

        self.assertIsInstance(prediction, LearnedSubjectPrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_kind, "reasoning")
        self.assertGreater(prediction.best_confidence, 0.8)

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

            prediction = predict_subject_distribution(
                stem="根据上述定义，下列符合定义的是",
                options=["甲", "乙", "丙", "丁"],
                model_path=model_path,
            )

        self.assertIsNone(prediction)
