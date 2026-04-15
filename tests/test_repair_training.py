import json
import tempfile
import unittest
from pathlib import Path

from scripts.train_repair_action_model import _rebalance_training_split, _sample_from_row, train_repair_action_model


class RepairActionTrainingTest(unittest.TestCase):
    def test_sample_from_row_uses_state_record(self):
        row = {
            "action": "split_embedded_next_question",
            "state_record": {
                "text": "[CURRENT] 72. 如果甲成立 73. 某商品现价是多少",
                "meta": {"option_count": 4.0, "prev_present": 1.0},
            },
        }

        sample = _sample_from_row(row)

        self.assertIn("[CURRENT]", sample["text"])
        self.assertEqual(sample["meta"]["option_count"], 4.0)

    def test_train_repair_action_model_on_tiny_dataset(self):
        try:
            import sklearn  # noqa: F401
        except ModuleNotFoundError:
            self.skipTest("未安装 scikit-learn")

        rows = [
            {
                "source_pdf": "pdf-a",
                "action": "split_embedded_next_question",
                "state_record": {"text": "72 题干里混入 73 题", "meta": {"option_count": 4.0}},
            },
            {
                "source_pdf": "pdf-b",
                "action": "split_embedded_next_question",
                "state_record": {"text": "当前题干包含下一题号", "meta": {"option_count": 4.0}},
            },
            {
                "source_pdf": "pdf-c",
                "action": "move_spilled_option_back",
                "state_record": {"text": "上一题 D 选项串进当前题干", "meta": {"option_count": 4.0}},
            },
            {
                "source_pdf": "pdf-d",
                "action": "move_spilled_option_back",
                "state_record": {"text": "D 选项被挂到下一题开头", "meta": {"option_count": 4.0}},
            },
        ]

        with tempfile.TemporaryDirectory() as tmp:
            dataset_path = Path(tmp) / "repair_actions.jsonl"
            model_path = Path(tmp) / "repair_action_classifier.pkl"
            with dataset_path.open("w", encoding="utf-8") as fh:
                for row in rows:
                    fh.write(json.dumps(row, ensure_ascii=False) + "\n")

            summary = train_repair_action_model(dataset_path, model_path)

            self.assertTrue(model_path.exists())
            self.assertTrue(Path(summary["metrics_path"]).exists())
            self.assertEqual(summary["label_count"], 2)
            self.assertIn(summary["split_strategy"], {"grouped_by_pdf", "grouped_by_pdf_partial_labels"})

    def test_train_repair_action_model_supports_family_labels(self):
        try:
            import sklearn  # noqa: F401
        except ModuleNotFoundError:
            self.skipTest("未安装 scikit-learn")

        rows = [
            {
                "source_pdf": "pdf-a",
                "action": "split_embedded_next_question",
                "action_family": "boundary_repair",
                "state_record": {"text": "当前题干包含下一题号", "meta": {"option_count": 4.0}},
            },
            {
                "source_pdf": "pdf-b",
                "action": "move_spilled_option_back",
                "action_family": "boundary_repair",
                "state_record": {"text": "D 选项被挂到下一题开头", "meta": {"option_count": 4.0}},
            },
            {
                "source_pdf": "pdf-c",
                "action": "move_data_intro_back_to_material",
                "action_family": "material_repair",
                "state_record": {"text": "资料说明被挂进首题题干", "meta": {"has_material": 1.0}},
            },
            {
                "source_pdf": "pdf-d",
                "action": "move_data_assets_to_material",
                "action_family": "material_repair",
                "state_record": {"text": "图表被挂到首题题干图片区", "meta": {"stem_asset_count": 2.0}},
            },
        ]

        with tempfile.TemporaryDirectory() as tmp:
            dataset_path = Path(tmp) / "repair_actions.jsonl"
            model_path = Path(tmp) / "repair_action_family_classifier.pkl"
            with dataset_path.open("w", encoding="utf-8") as fh:
                for row in rows:
                    fh.write(json.dumps(row, ensure_ascii=False) + "\n")

            summary = train_repair_action_model(dataset_path, model_path, label_field="action_family")

            self.assertTrue(model_path.exists())
            self.assertEqual(summary["label_field"], "action_family")
            self.assertEqual(summary["label_count"], 2)

    def test_rebalance_training_split_upsamples_rare_labels(self):
        samples = [{"text": "a"}, {"text": "b"}, {"text": "c"}]
        labels = ["boundary_repair", "boundary_repair", "asset_repair"]

        balanced_samples, balanced_labels, before, after = _rebalance_training_split(samples, labels)

        self.assertEqual(before["boundary_repair"], 2)
        self.assertEqual(before["asset_repair"], 1)
        self.assertGreater(after["asset_repair"], before["asset_repair"])
        self.assertEqual(len(balanced_samples), len(balanced_labels))


if __name__ == "__main__":
    unittest.main()
