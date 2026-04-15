import unittest
from pathlib import Path
import tempfile
from unittest.mock import patch

from PIL import Image, ImageDraw

from core.pdf_exam_extract import (
    _is_background_like_image_block,
    _is_decorative_image_block,
    _is_decorative_text_line,
    _is_learned_noise_image_block,
    _is_noise_text_line,
    _is_watermark_like_image_block,
    _should_drop_text_line_record,
    _merge_page_image_blocks,
    _order_page_blocks,
    segments_to_lines,
)


class _DummyRect:
    def __init__(self, width: float, height: float = 800):
        self.width = width
        self.height = height


class _DummyPage:
    def __init__(self, width: float, height: float = 800):
        self.rect = _DummyRect(width, height)
        self._image_infos = []

    def get_image_info(self, xrefs=False):
        return list(self._image_infos)


def _text_block(label: str, x0: float, y0: float, x1: float, y1: float, text: str | None = None) -> dict:
    content = text if text is not None else label
    return {
        "type": 0,
        "bbox": (x0, y0, x1, y1),
        "label": label,
        "lines": [
            {
                "bbox": (x0, y0, x1, y1),
                "spans": [{"text": content}],
            }
        ],
    }


def _empty_text_block(label: str, x0: float, y0: float, x1: float, y1: float) -> dict:
    return {
        "type": 0,
        "bbox": (x0, y0, x1, y1),
        "label": label,
        "lines": [],
    }


def _image_block(label: str, x0: float, y0: float, x1: float, y1: float) -> dict:
    return {
        "type": 1,
        "bbox": (x0, y0, x1, y1),
        "label": label,
    }


class PdfExamExtractTest(unittest.TestCase):
    def test_decorative_header_logo_image_is_filtered(self):
        page = _DummyPage(width=600, height=800)
        block = _image_block("logo", 20, 24, 140, 52)

        self.assertTrue(_is_decorative_image_block(page, block))

    def test_tiny_inline_image_is_filtered(self):
        page = _DummyPage(width=600, height=800)
        block = _image_block("inline-frag", 156, 146, 170, 157)

        self.assertTrue(_is_decorative_image_block(page, block))

    def test_large_content_image_is_not_filtered_by_size_only(self):
        page = _DummyPage(width=600, height=800)
        block = _image_block("figure", 24, 108, 576, 742)

        self.assertFalse(_is_decorative_image_block(page, block))

    def test_centered_footer_page_number_is_filtered(self):
        page = _DummyPage(width=600, height=800)
        line = {
            "bbox": (294.0, 748.0, 306.0, 760.0),
            "spans": [{"text": "7"}],
        }

        self.assertTrue(_is_decorative_text_line(page, line, "7"))

    def test_scan_ad_text_is_filtered(self):
        self.assertTrue(_is_noise_text_line("各种考试资料购买，请加微信：行测资料库"))

    def test_scan_ad_vertical_chars_are_filtered(self):
        page = _DummyPage(width=600, height=800)
        line = {
            "bbox": (456.3, 22.8, 466.8, 33.3),
            "spans": [{"text": "扫"}],
        }

        self.assertTrue(_is_decorative_text_line(page, line, "扫"))

    def test_learned_noise_model_can_filter_suspicious_top_line(self):
        record = {
            "text": "资料领取请加v信资料库",
            "page_number": 1,
            "page_width": 600.0,
            "page_height": 800.0,
            "x0": 40.0,
            "y0": 16.0,
            "x1": 260.0,
            "y1": 30.0,
            "line_index_in_block": 0,
            "line_count_in_block": 1,
        }

        with patch(
            "core.pdf_exam_extract.predict_pdf_noise_distribution",
            return_value=type(
                "Prediction",
                (),
                {"best_label": "noise", "best_confidence": 0.97},
            )(),
        ):
            self.assertTrue(_should_drop_text_line_record(record))

    def test_middle_body_line_is_not_filtered_by_learned_noise_model_candidate_gate(self):
        record = {
            "text": "根据上述材料，下列说法正确的是",
            "page_number": 1,
            "page_width": 600.0,
            "page_height": 800.0,
            "x0": 72.0,
            "y0": 220.0,
            "x1": 420.0,
            "y1": 242.0,
            "line_index_in_block": 0,
            "line_count_in_block": 1,
        }

        with patch(
            "core.pdf_exam_extract.predict_pdf_noise_distribution",
            side_effect=AssertionError("正文中段不该触发噪声模型"),
        ):
            self.assertFalse(_should_drop_text_line_record(record))

    def test_top_banner_and_corner_qr_images_are_filtered(self):
        page = _DummyPage(width=596, height=842)
        top_banner = _image_block("banner", 11.9, 0.0, 583.4, 68.5)
        qr = _image_block("qr", 477.6, 8.0, 531.5, 61.9)
        logo = _image_block("logo", 64.4, 9.6, 112.4, 57.6)

        self.assertTrue(_is_decorative_image_block(page, top_banner))
        self.assertTrue(_is_decorative_image_block(page, qr))
        self.assertTrue(_is_decorative_image_block(page, logo))

    def test_learned_noise_model_can_filter_suspicious_top_image(self):
        page = _DummyPage(width=596, height=842)
        block = _image_block("candidate", 120.0, 24.0, 430.0, 120.0)

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "candidate.png"
            Image.new("RGB", (96, 32), color="white").save(image_path)
            with patch(
                "core.pdf_exam_extract.predict_pdf_noise_image_distribution",
                return_value=type(
                    "Prediction",
                    (),
                    {"best_label": "noise", "best_confidence": 0.98},
                )(),
            ):
                self.assertTrue(_is_learned_noise_image_block(page, block, str(image_path), page_number=1))

    def test_blank_full_page_background_image_is_filtered(self):
        page = _DummyPage(width=596, height=842)
        block = _image_block("background", 0.0, 90.0, 596.0, 782.0)

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "background.png"
            Image.new("RGB", (595, 692), color="black").save(image_path)
            self.assertTrue(_is_background_like_image_block(page, block, str(image_path), page_number=1))

    def test_high_contrast_full_page_figure_is_not_filtered_as_background(self):
        page = _DummyPage(width=596, height=842)
        block = _image_block("figure", 12.0, 94.0, 584.0, 770.0)

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "figure.png"
            image = Image.new("RGB", (256, 256), color="white")
            for x in range(0, 256, 16):
                for y in range(0, 256, 16):
                    if (x // 16 + y // 16) % 2 == 0:
                        for dx in range(16):
                            for dy in range(16):
                                image.putpixel((x + dx, y + dy), (0, 0, 0))
            image.save(image_path)
            self.assertFalse(_is_background_like_image_block(page, block, str(image_path), page_number=1))

    def test_faint_full_page_watermark_image_is_filtered(self):
        page = _DummyPage(width=596, height=842)
        block = _image_block("watermark", 8.0, 92.0, 588.0, 782.0)

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "watermark.png"
            image = Image.new("RGB", (256, 256), color=(248, 248, 248))
            draw = ImageDraw.Draw(image)
            for offset in range(-80, 280, 28):
                draw.line((offset, 0, offset + 120, 256), fill=(212, 212, 212), width=6)
            image.save(image_path)
            self.assertTrue(_is_watermark_like_image_block(page, block, str(image_path), page_number=1))

    def test_page_scan_like_full_page_image_is_not_filtered_as_watermark(self):
        page = _DummyPage(width=596, height=842)
        block = _image_block("scan", 0.0, 90.0, 596.0, 782.0)

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "scan.png"
            image = Image.new("RGB", (256, 256), color="white")
            draw = ImageDraw.Draw(image)
            for y in range(24, 232, 18):
                draw.rectangle((18, y, 228, y + 3), fill=(0, 0, 0))
            image.save(image_path)
            self.assertFalse(_is_watermark_like_image_block(page, block, str(image_path), page_number=1))

    def test_middle_content_image_does_not_trigger_image_noise_model(self):
        page = _DummyPage(width=596, height=842)
        block = _image_block("figure", 88.0, 220.0, 512.0, 600.0)

        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "figure.png"
            Image.new("RGB", (160, 160), color="white").save(image_path)
            with patch(
                "core.pdf_exam_extract.predict_pdf_noise_image_distribution",
                side_effect=AssertionError("正文主图不该触发图片噪声模型"),
            ):
                self.assertFalse(_is_learned_noise_image_block(page, block, str(image_path), page_number=1))

    def test_segments_to_lines_strips_scan_ad_sequence(self):
        segments = [
            ("题干第一行", None),
            ("扫", None),
            ("码", None),
            ("关", None),
            ("注", None),
            ("各种考试资料购买，请加微信：行测资料库", None),
            ("题干第二行", None),
        ]

        self.assertEqual(
            segments_to_lines(segments),
            [("题干第一行", None), ("题干第二行", None)],
        )

    def test_order_page_blocks_prefers_left_then_right_for_two_column_pages(self):
        page = _DummyPage(width=600, height=500)
        blocks = [
            _text_block("header", 40, 10, 560, 40),
            _text_block("left-1", 40, 60, 250, 95),
            _text_block("right-1", 340, 62, 560, 98),
            _text_block("left-2", 42, 110, 252, 145),
            _text_block("right-2", 342, 112, 562, 148),
            _text_block("left-3", 44, 160, 254, 195),
            _text_block("right-3", 344, 162, 564, 198),
            _text_block("footer", 40, 500, 560, 540),
        ]

        ordered = _order_page_blocks(page, blocks)

        self.assertEqual(
            [block["label"] for block in ordered],
            ["header", "left-1", "left-2", "left-3", "right-1", "right-2", "right-3", "footer"],
        )

    def test_order_page_blocks_keeps_default_order_for_single_column_pages(self):
        page = _DummyPage(width=600)
        blocks = [
            _text_block("line-1", 40, 20, 540, 50),
            _text_block("line-2", 40, 70, 540, 100),
            _text_block("line-3", 40, 120, 540, 150),
        ]

        ordered = _order_page_blocks(page, blocks)

        self.assertEqual(
            [block["label"] for block in ordered],
            ["line-1", "line-2", "line-3"],
        )

    def test_order_page_blocks_ignores_empty_right_side_blocks_when_detecting_columns(self):
        page = _DummyPage(width=600)
        blocks = [
            _text_block("q36", 40, 60, 560, 95),
            _empty_text_block("ghost-right-1", 392, 62, 405, 90),
            _text_block("q37", 40, 110, 560, 145),
            _empty_text_block("ghost-right-2", 523, 112, 537, 140),
            _text_block("q38", 40, 160, 560, 195),
        ]

        ordered = _order_page_blocks(page, blocks)

        self.assertEqual(
            [block["label"] for block in ordered],
            ["q36", "ghost-right-1", "q37", "ghost-right-2", "q38"],
        )

    def test_order_page_blocks_keeps_y_order_when_right_band_has_little_vertical_overlap(self):
        page = _DummyPage(width=600, height=800)
        blocks = [
            _text_block("q80-1", 40, 60, 560, 95),
            _text_block("q80-2", 40, 100, 560, 135),
            _text_block("prompt", 40, 150, 250, 185),
            _text_block("option-a", 40, 190, 250, 225),
            _text_block("formula-left", 40, 560, 170, 590),
            _text_block("formula-right", 360, 565, 560, 595),
            _text_block("tail-right", 430, 640, 560, 668, text="总检测天数"),
        ]

        ordered = _order_page_blocks(page, blocks)

        self.assertEqual(
            [block["label"] for block in ordered],
            ["q80-1", "q80-2", "prompt", "option-a", "formula-left", "formula-right", "tail-right"],
        )

    def test_merge_page_image_blocks_adds_missing_images_from_page_info(self):
        page = _DummyPage(width=600)
        page._image_infos = [
            {"xref": 12, "bbox": (100.0, 120.0, 240.0, 220.0)},
        ]
        blocks = [
            _text_block("text", 40, 20, 540, 60),
        ]

        merged = _merge_page_image_blocks(page, blocks)

        self.assertEqual(len(merged), 2)
        self.assertEqual(merged[1]["type"], 1)
        self.assertEqual(merged[1]["xref"], 12)
        self.assertEqual(merged[1]["bbox"], (100.0, 120.0, 240.0, 220.0))

    def test_merge_page_image_blocks_skips_duplicate_bbox_images(self):
        page = _DummyPage(width=600)
        page._image_infos = [
            {"xref": 12, "bbox": (100.0, 120.0, 240.0, 220.0)},
        ]
        blocks = [
            _text_block("text", 40, 20, 540, 60),
            {**_image_block("img", 100.0, 120.0, 240.0, 220.0), "xref": 12, "ext": "png"},
        ]

        merged = _merge_page_image_blocks(page, blocks)

        self.assertEqual(len(merged), 2)


if __name__ == "__main__":
    unittest.main()
