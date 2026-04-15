import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

import core.repair_action_model as repair_action_model
from core.repair_action_model import RepairActionPrediction, clear_repair_model_cache, predict_repair_action_distribution
from domain.models import MaterialSet, OptionNode, QuestionNode, Section


class DummyRepairFamilyModel:
    class _DecisionMatrix(list):
        ndim = 2

    def decision_function(self, rows):
        text = rows[0]["text"]
        if "同比增长8.4%" in text:
            return self._DecisionMatrix([[-0.5, 0.1, 2.7]])
        return self._DecisionMatrix([[0.1, 2.4, -0.6]])


class RepairActionModelTest(unittest.TestCase):
    def tearDown(self):
        clear_repair_model_cache()

    def test_predict_repair_action_distribution_loads_family_bundle_via_env_model_path(self):
        section = Section(kind="data", title="资料分析")
        material = MaterialSet(material_id="m1", header="材料一", body="")
        question = QuestionNode(
            source_number="101",
            stem="2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。根据上述资料，下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_action_family_classifier.pkl"
            bundle = {
                "labels": ["asset_repair", "boundary_repair", "material_repair"],
                "label_field": "action_family",
                "model": DummyRepairFamilyModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_REPAIR_MODEL": "1",
                    "PPTCONVERT_TRUST_REPAIR_MODEL": "1",
                    "PPTCONVERT_REPAIR_FAMILY_MODEL": str(model_path),
                },
                clear=False,
            ):
                prediction = predict_repair_action_distribution(
                    section=section,
                    material=material,
                    question=question,
                    label_field="action_family",
                )

        self.assertIsInstance(prediction, RepairActionPrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_label, "material_repair")
        self.assertGreater(prediction.best_confidence, 0.8)

    def test_predict_repair_action_distribution_loads_ready_default_family_bundle_without_enable_flag(self):
        section = Section(kind="unknown", title="待确认")
        question = QuestionNode(
            source_number="1",
            stem="下列关于法律常识的说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_action_family_classifier.pkl"
            bundle = {
                "labels": ["asset_repair", "boundary_repair", "material_repair"],
                "label_field": "action_family",
                "model": DummyRepairFamilyModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with (
                patch.object(repair_action_model, "_DEFAULT_FAMILY_MODEL_PATH", model_path),
                patch.object(repair_action_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_PICKLED_REPAIR_MODEL": "",
                        "PPTCONVERT_REPAIR_FAMILY_MODEL": "",
                        "PPTCONVERT_TRUST_REPAIR_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                prediction = predict_repair_action_distribution(
                    section=section,
                    material=None,
                    question=question,
                    label_field="action_family",
                )

        self.assertIsInstance(prediction, RepairActionPrediction)
        assert prediction is not None
        self.assertEqual(prediction.best_label, "boundary_repair")
        self.assertGreater(prediction.best_confidence, 0.8)

    def test_predict_repair_action_distribution_skips_not_ready_bundle(self):
        section = Section(kind="data", title="资料分析")
        material = MaterialSet(material_id="m1", header="材料一", body="")
        question = QuestionNode(
            source_number="101",
            stem="2023年全市规模以上工业增加值同比增长8.4%，根据上述资料，下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_action_family_classifier.pkl"
            bundle = {
                "labels": ["asset_repair", "boundary_repair", "material_repair"],
                "label_field": "action_family",
                "model": DummyRepairFamilyModel(),
                "ready_for_runtime": False,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_REPAIR_MODEL": "1",
                    "PPTCONVERT_TRUST_REPAIR_MODEL": "1",
                    "PPTCONVERT_REPAIR_FAMILY_MODEL": str(model_path),
                },
                clear=False,
            ):
                prediction = predict_repair_action_distribution(
                    section=section,
                    material=material,
                    question=question,
                    label_field="action_family",
                )

        self.assertIsNone(prediction)


if __name__ == "__main__":
    unittest.main()
