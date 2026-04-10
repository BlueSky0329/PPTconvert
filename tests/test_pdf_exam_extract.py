import unittest

from core.pdf_exam_extract import _merge_page_image_blocks, _order_page_blocks


class _DummyRect:
    def __init__(self, width: float):
        self.width = width


class _DummyPage:
    def __init__(self, width: float):
        self.rect = _DummyRect(width)
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
    def test_order_page_blocks_prefers_left_then_right_for_two_column_pages(self):
        page = _DummyPage(width=600)
        blocks = [
            _text_block("header", 40, 10, 560, 40),
            _text_block("left-1", 40, 60, 250, 95),
            _text_block("right-1", 340, 62, 560, 98),
            _text_block("left-2", 42, 110, 252, 145),
            _text_block("right-2", 342, 112, 562, 148),
            _text_block("footer", 40, 500, 560, 540),
        ]

        ordered = _order_page_blocks(page, blocks)

        self.assertEqual(
            [block["label"] for block in ordered],
            ["header", "left-1", "left-2", "right-1", "right-2", "footer"],
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
