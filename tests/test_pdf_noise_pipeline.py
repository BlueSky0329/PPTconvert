import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from scripts.build_pdf_noise_text_dataset import build_dataset
from scripts.train_pdf_noise_text_model import train_pdf_noise_model


class PdfNoisePipelineTest(unittest.TestCase):
    def test_build_dataset_collects_noise_and_content_rows(self):
        records = [
            {
                "text": "各种考试资料购买，请加微信：行测资料库",
                "page_number": 1,
                "page_width": 600.0,
                "page_height": 800.0,
                "x0": 40.0,
                "y0": 18.0,
                "x1": 260.0,
                "y1": 32.0,
                "line_index_in_block": 0,
                "line_count_in_block": 1,
            },
            {
                "text": "根据上述材料，下列说法正确的是",
                "page_number": 1,
                "page_width": 600.0,
                "page_height": 800.0,
                "x0": 72.0,
                "y0": 80.0,
                "x1": 420.0,
                "y1": 102.0,
                "line_index_in_block": 0,
                "line_count_in_block": 1,
            },
        ]

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pdf_path = root / "sample.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            catalog_path = root / "catalog.json"
            output_path = root / "pdf_noise_text.jsonl"
            catalog_path.write_text(
                json.dumps({"pdfs": [{"path": "sample.pdf", "form": "set_paper"}]}, ensure_ascii=False),
                encoding="utf-8",
            )

            with patch("scripts.build_pdf_noise_text_dataset.ROOT", root), patch(
                "scripts.build_pdf_noise_text_dataset.iter_pdf_text_line_records",
                return_value=records,
            ):
                summary = build_dataset(catalog_path, output_path)

            rows = [json.loads(line) for line in output_path.read_text(encoding="utf-8").splitlines() if line.strip()]

        self.assertEqual(summary["row_count"], 2)
        self.assertEqual(summary["label_distribution"]["noise"], 1)
        self.assertEqual(summary["label_distribution"]["content"], 1)
        self.assertIn("feature_record", rows[0])
        self.assertEqual(rows[0]["source_pdf"], "sample.pdf")

    def test_train_pdf_noise_model_on_tiny_dataset(self):
        with tempfile.TemporaryDirectory() as tmp:
            dataset_path = Path(tmp) / "pdf_noise_text.jsonl"
            model_path = Path(tmp) / "pdf_noise_text_classifier.pkl"
            rows = [
                {
                    "source_pdf": "pdf-a",
                    "label": "noise",
                    "feature_record": {
                        "text": "[ZONE] top\n[TEXT] 各种考试资料购买 请加微信",
                        "meta": {"top_band": 1.0, "contains_ad_keyword": 1.0, "text_length": 14.0},
                    },
                },
                {
                    "source_pdf": "pdf-a",
                    "label": "content",
                    "feature_record": {
                        "text": "[ZONE] middle\n[TEXT] 根据上述材料 下列说法正确的是",
                        "meta": {"top_band": 0.0, "contains_ad_keyword": 0.0, "text_length": 16.0},
                    },
                },
                {
                    "source_pdf": "pdf-b",
                    "label": "noise",
                    "feature_record": {
                        "text": "[ZONE] bottom\n[TEXT] 第 1 页 共 20 页",
                        "meta": {"bottom_band": 1.0, "contains_page_keyword": 1.0, "text_length": 10.0},
                    },
                },
                {
                    "source_pdf": "pdf-b",
                    "label": "content",
                    "feature_record": {
                        "text": "[ZONE] middle\n[TEXT] 甲 乙 两地相距 240 千米",
                        "meta": {"bottom_band": 0.0, "contains_page_keyword": 0.0, "text_length": 14.0},
                    },
                },
                {
                    "source_pdf": "pdf-c",
                    "label": "noise",
                    "feature_record": {
                        "text": "[ZONE] top\n[TEXT] 行测 资料 领取",
                        "meta": {"top_band": 1.0, "contains_header_keyword": 1.0, "text_length": 8.0},
                    },
                },
                {
                    "source_pdf": "pdf-c",
                    "label": "content",
                    "feature_record": {
                        "text": "[ZONE] middle\n[TEXT] 下列关于宪法的说法正确的是",
                        "meta": {"top_band": 0.0, "contains_header_keyword": 0.0, "text_length": 15.0},
                    },
                },
            ]
            dataset_path.write_text(
                "\n".join(json.dumps(row, ensure_ascii=False) for row in rows),
                encoding="utf-8",
            )

            summary = train_pdf_noise_model(dataset_path, model_path)

            self.assertTrue(model_path.exists())
            self.assertGreaterEqual(summary["macro_f1"], 0.0)
            self.assertIn("ready_for_runtime", summary)


if __name__ == "__main__":
    unittest.main()
