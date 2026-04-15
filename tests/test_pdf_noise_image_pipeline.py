import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from PIL import Image

from scripts.build_pdf_noise_image_dataset import build_dataset
from scripts.train_pdf_noise_image_model import train_pdf_noise_image_model


class PdfNoiseImagePipelineTest(unittest.TestCase):
    def test_build_dataset_collects_image_rows(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pdf_path = root / "sample.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            catalog_path = root / "catalog.json"
            output_path = root / "pdf_noise_images.jsonl"
            image_path = root / "sample.png"
            Image.new("RGB", (64, 64), color="white").save(image_path)
            catalog_path.write_text(
                json.dumps({"pdfs": [{"path": "sample.pdf", "form": "set_paper"}]}, ensure_ascii=False),
                encoding="utf-8",
            )
            fake_rows = [
                {
                    "source_pdf": "sample.pdf",
                    "label": "noise",
                    "label_source": "rule_decorative_image",
                    "image_path": str(image_path),
                    "page_number": 1,
                    "bbox": [10.0, 10.0, 74.0, 74.0],
                    "page_size": [596.0, 842.0],
                    "image_size": [64, 64],
                    "meta_record": {"top_band": 1.0, "banner_like": 0.0},
                },
                {
                    "source_pdf": "sample.pdf",
                    "label": "content",
                    "label_source": "default_content_image",
                    "image_path": str(image_path),
                    "page_number": 1,
                    "bbox": [120.0, 220.0, 320.0, 420.0],
                    "page_size": [596.0, 842.0],
                    "image_size": [64, 64],
                    "meta_record": {"top_band": 0.0, "banner_like": 0.0},
                },
            ]

            with patch("scripts.build_pdf_noise_image_dataset.ROOT", root), patch(
                "scripts.build_pdf_noise_image_dataset.require_fitz",
                return_value=None,
            ), patch(
                "scripts.build_pdf_noise_image_dataset._iter_image_records_from_pdf",
                return_value=fake_rows,
            ):
                summary = build_dataset(catalog_path, output_path)

            rows = [json.loads(line) for line in output_path.read_text(encoding="utf-8").splitlines() if line.strip()]

        self.assertEqual(summary["row_count"], 2)
        self.assertEqual(summary["label_distribution"]["noise"], 1)
        self.assertEqual(summary["label_distribution"]["content"], 1)
        self.assertEqual(rows[0]["source_pdf"], "sample.pdf")

    def test_build_dataset_rewrites_image_paths_to_final_assets_dir(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pdf_path = root / "sample.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            catalog_path = root / "catalog.json"
            output_path = root / "pdf_noise_images.jsonl"
            catalog_path.write_text(
                json.dumps({"pdfs": [{"path": "sample.pdf", "form": "set_paper"}]}, ensure_ascii=False),
                encoding="utf-8",
            )

            def fake_iter(_pdf_path, *, source_pdf, assets_dir):
                staged_image = assets_dir / "stage.png"
                staged_image.write_bytes(b"fake")
                return [
                    {
                        "source_pdf": source_pdf,
                        "label": "noise",
                        "label_source": "rule_decorative_image",
                        "image_path": str(staged_image),
                        "page_number": 1,
                        "bbox": [10.0, 10.0, 74.0, 74.0],
                        "page_size": [596.0, 842.0],
                        "image_size": [64, 64],
                        "meta_record": {"top_band": 1.0, "banner_like": 0.0},
                    }
                ]

            with patch("scripts.build_pdf_noise_image_dataset.ROOT", root), patch(
                "scripts.build_pdf_noise_image_dataset.require_fitz",
                return_value=None,
            ), patch(
                "scripts.build_pdf_noise_image_dataset._iter_image_records_from_pdf",
                side_effect=fake_iter,
            ):
                build_dataset(catalog_path, output_path)

            rows = [json.loads(line) for line in output_path.read_text(encoding="utf-8").splitlines() if line.strip()]

        self.assertEqual(
            Path(rows[0]["image_path"]).as_posix(),
            (output_path.parent / "pdf_noise_images_assets" / "stage.png").as_posix(),
        )

    def test_build_dataset_preserves_existing_outputs_when_startup_check_fails(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            catalog_path = root / "catalog.json"
            output_path = root / "pdf_noise_images.jsonl"
            assets_dir = root / "pdf_noise_images_assets"
            assets_dir.mkdir(parents=True, exist_ok=True)
            old_asset = assets_dir / "keep.txt"
            old_asset.write_text("old-asset", encoding="utf-8")
            output_path.write_text("old-output\n", encoding="utf-8")
            summary_path = output_path.with_suffix(".summary.json")
            summary_path.write_text('{"old": true}', encoding="utf-8")
            catalog_path.write_text(json.dumps({"pdfs": []}, ensure_ascii=False), encoding="utf-8")

            with self.assertRaisesRegex(RuntimeError, "PyMuPDF"):
                with patch("scripts.build_pdf_noise_image_dataset.require_fitz", side_effect=RuntimeError("需要安装 PyMuPDF")):
                    build_dataset(catalog_path, output_path)

            self.assertEqual(output_path.read_text(encoding="utf-8"), "old-output\n")
            self.assertEqual(summary_path.read_text(encoding="utf-8"), '{"old": true}')
            self.assertEqual(old_asset.read_text(encoding="utf-8"), "old-asset")

    def test_train_pdf_noise_image_model_on_tiny_dataset(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            dataset_path = root / "pdf_noise_images.jsonl"
            model_path = root / "pdf_noise_image_classifier.pt"

            def make_image(name: str, color: str):
                path = root / name
                Image.new("RGB", (64, 64), color=color).save(path)
                return str(path)

            rows = [
                {
                    "source_pdf": "pdf-a",
                    "label": "noise",
                    "image_path": make_image("noise-a.png", "white"),
                    "page_number": 1,
                    "bbox": [12.0, 10.0, 90.0, 58.0],
                    "page_size": [596.0, 842.0],
                    "image_size": [64, 64],
                    "meta_record": {
                        "page_number": 1.0,
                        "page_top_ratio": 0.01,
                        "page_bottom_ratio": 0.06,
                        "width_ratio": 0.13,
                        "height_ratio": 0.06,
                        "image_aspect_ratio": 1.0,
                        "center_offset_ratio": 0.32,
                        "top_band": 1.0,
                        "bottom_band": 0.0,
                        "left_edge": 1.0,
                        "right_edge": 0.0,
                        "corner_like": 1.0,
                        "banner_like": 0.0,
                        "small_image": 1.0,
                    },
                },
                {
                    "source_pdf": "pdf-a",
                    "label": "content",
                    "image_path": make_image("content-a.png", "black"),
                    "page_number": 1,
                    "bbox": [140.0, 220.0, 420.0, 520.0],
                    "page_size": [596.0, 842.0],
                    "image_size": [64, 64],
                    "meta_record": {
                        "page_number": 1.0,
                        "page_top_ratio": 0.26,
                        "page_bottom_ratio": 0.62,
                        "width_ratio": 0.47,
                        "height_ratio": 0.36,
                        "image_aspect_ratio": 1.0,
                        "center_offset_ratio": 0.0,
                        "top_band": 0.0,
                        "bottom_band": 0.0,
                        "left_edge": 0.0,
                        "right_edge": 0.0,
                        "corner_like": 0.0,
                        "banner_like": 0.0,
                        "small_image": 0.0,
                    },
                },
                {
                    "source_pdf": "pdf-b",
                    "label": "noise",
                    "image_path": make_image("noise-b.png", "white"),
                    "page_number": 1,
                    "bbox": [500.0, 12.0, 560.0, 72.0],
                    "page_size": [596.0, 842.0],
                    "image_size": [64, 64],
                    "meta_record": {
                        "page_number": 1.0,
                        "page_top_ratio": 0.01,
                        "page_bottom_ratio": 0.08,
                        "width_ratio": 0.10,
                        "height_ratio": 0.07,
                        "image_aspect_ratio": 1.0,
                        "center_offset_ratio": 0.39,
                        "top_band": 1.0,
                        "bottom_band": 0.0,
                        "left_edge": 0.0,
                        "right_edge": 1.0,
                        "corner_like": 1.0,
                        "banner_like": 0.0,
                        "small_image": 1.0,
                    },
                },
                {
                    "source_pdf": "pdf-b",
                    "label": "content",
                    "image_path": make_image("content-b.png", "black"),
                    "page_number": 1,
                    "bbox": [160.0, 240.0, 460.0, 580.0],
                    "page_size": [596.0, 842.0],
                    "image_size": [64, 64],
                    "meta_record": {
                        "page_number": 1.0,
                        "page_top_ratio": 0.28,
                        "page_bottom_ratio": 0.69,
                        "width_ratio": 0.50,
                        "height_ratio": 0.40,
                        "image_aspect_ratio": 1.0,
                        "center_offset_ratio": 0.02,
                        "top_band": 0.0,
                        "bottom_band": 0.0,
                        "left_edge": 0.0,
                        "right_edge": 0.0,
                        "corner_like": 0.0,
                        "banner_like": 0.0,
                        "small_image": 0.0,
                    },
                },
            ]
            dataset_path.write_text(
                "\n".join(json.dumps(row, ensure_ascii=False) for row in rows),
                encoding="utf-8",
            )

            summary = train_pdf_noise_image_model(
                dataset_path,
                model_path,
                image_size=48,
                epochs=1,
                batch_size=2,
                learning_rate=1e-3,
            )

            self.assertTrue(model_path.exists())
            self.assertGreaterEqual(summary["macro_f1"], 0.0)
            self.assertIn("ready_for_runtime", summary)


if __name__ == "__main__":
    unittest.main()
