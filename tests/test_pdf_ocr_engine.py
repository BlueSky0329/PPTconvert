import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import core.pdf_ocr_engine as pdf_ocr_engine
from core.pdf_ocr_engine import (
    OCRLine,
    _cache_path,
    _fingerprint_pdf,
    _parse_rapidocr_result,
    clear_ocr_cache,
    is_ocr_available,
    is_ocr_dependency_available,
    ocr_pdf_page,
    synthesize_text_segments,
)


class OCRLineRoundtripTest(unittest.TestCase):
    def test_to_dict_from_dict_roundtrip(self):
        original = OCRLine(
            text="试题一",
            bbox=(12.5, 30.0, 180.0, 48.0),
            confidence=0.94,
            page_number=3,
        )
        restored = OCRLine.from_dict(original.to_dict())
        self.assertEqual(restored, original)

    def test_from_dict_tolerates_bad_values(self):
        restored = OCRLine.from_dict(
            {
                "text": None,
                "bbox": None,
                "confidence": "bad",
                "page_number": "bad",
            }
        )
        self.assertEqual(restored.text, "")
        self.assertEqual(restored.bbox, (0.0, 0.0, 0.0, 0.0))
        self.assertEqual(restored.confidence, 0.0)
        self.assertEqual(restored.page_number, 0)


class IsOCRAvailableTest(unittest.TestCase):
    def tearDown(self):
        pdf_ocr_engine.reset_engine_cache()

    def test_returns_bool_without_crash(self):
        self.assertIsInstance(is_ocr_available(), bool)

    def test_env_disable_forces_unavailable(self):
        with patch.dict("os.environ", {pdf_ocr_engine._DISABLE_ENV: "1"}, clear=False):
            pdf_ocr_engine.reset_engine_cache()
            self.assertFalse(is_ocr_available())
            self.assertFalse(is_ocr_dependency_available())

    def test_dependency_check_does_not_load_engine(self):
        with (
            patch.object(
                pdf_ocr_engine,
                "_load_engine",
                side_effect=AssertionError("dependency check should stay lightweight"),
            ),
            patch(
                "core.pdf_ocr_engine.importlib.util.find_spec",
                return_value=object(),
            ),
        ):
            self.assertTrue(is_ocr_dependency_available())


class SynthesizeTextSegmentsTest(unittest.TestCase):
    def test_empty_input_returns_empty(self):
        self.assertEqual(synthesize_text_segments({}), [])

    def test_drops_low_confidence_lines(self):
        pages = {
            1: [
                OCRLine("高置信度", (10.0, 10.0, 60.0, 24.0), 0.9, 1),
                OCRLine("低置信度", (10.0, 30.0, 60.0, 44.0), 0.05, 1),
            ]
        }
        segments = synthesize_text_segments(pages, min_confidence=0.5)
        self.assertEqual(segments, [("高置信度", None)])

    def test_joins_same_line_and_splits_across_lines(self):
        pages = {
            1: [
                OCRLine("左侧", (10.0, 10.0, 40.0, 24.0), 0.9, 1),
                OCRLine("右侧", (50.0, 10.0, 80.0, 24.0), 0.9, 1),
                OCRLine("下一行", (10.0, 40.0, 80.0, 54.0), 0.9, 1),
            ]
        }
        segments = synthesize_text_segments(pages)
        self.assertEqual(
            segments,
            [("左侧 右侧", None), ("下一行", None)],
        )

    def test_orders_by_page_number(self):
        pages = {
            2: [OCRLine("第二页", (10.0, 10.0, 60.0, 24.0), 0.9, 2)],
            1: [OCRLine("第一页", (10.0, 10.0, 60.0, 24.0), 0.9, 1)],
        }
        segments = synthesize_text_segments(pages)
        self.assertEqual(segments, [("第一页", None), ("第二页", None)])


class ParseRapidOCRResultTest(unittest.TestCase):
    def test_parses_well_formed_entries(self):
        raw = [
            [
                [[10.0, 20.0], [120.0, 20.0], [120.0, 44.0], [10.0, 44.0]],
                "题干一",
                0.92,
            ]
        ]
        lines = _parse_rapidocr_result(raw, page_number=4)
        self.assertEqual(len(lines), 1)
        line = lines[0]
        self.assertEqual(line.text, "题干一")
        self.assertEqual(line.bbox, (10.0, 20.0, 120.0, 44.0))
        self.assertAlmostEqual(line.confidence, 0.92)
        self.assertEqual(line.page_number, 4)

    def test_skips_empty_and_malformed_entries(self):
        raw = [
            None,
            [[], "", 0.9],
            [[[1.0, 1.0]], "只有一个点", 0.9],
            [
                [[0.0, 0.0], [10.0, 0.0], [10.0, 10.0], [0.0, 10.0]],
                "   ",
                0.9,
            ],
        ]
        self.assertEqual(_parse_rapidocr_result(raw, page_number=1), [])

    def test_empty_result_returns_empty_list(self):
        self.assertEqual(_parse_rapidocr_result([], page_number=1), [])
        self.assertEqual(_parse_rapidocr_result(None, page_number=1), [])


class PageCacheTest(unittest.TestCase):
    def tearDown(self):
        pdf_ocr_engine.reset_engine_cache()

    def test_ocr_pdf_page_missing_file_returns_empty(self):
        self.assertEqual(ocr_pdf_page("/tmp/does-not-exist.pdf", 1), [])

    def test_ocr_pdf_page_without_engine_returns_empty(self):
        with patch.object(pdf_ocr_engine, "_load_engine", return_value=None):
            with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as handle:
                handle.write(b"fake")
                fake_pdf = Path(handle.name)
            try:
                self.assertEqual(ocr_pdf_page(fake_pdf, 1), [])
            finally:
                fake_pdf.unlink(missing_ok=True)

    def test_ocr_pdf_page_uses_cache_when_available(self):
        fake_engine = object()
        cached_lines = [
            OCRLine("缓存命中", (0.0, 0.0, 10.0, 10.0), 0.99, 1),
        ]
        with tempfile.TemporaryDirectory() as tmp_cache:
            cache_dir = Path(tmp_cache)
            with (
                patch.object(pdf_ocr_engine, "_CACHE_DIR", cache_dir),
                patch.object(pdf_ocr_engine, "_load_engine", return_value=fake_engine),
                patch.object(pdf_ocr_engine, "_render_page_image") as render_mock,
            ):
                with tempfile.NamedTemporaryFile(suffix=".pdf", delete=False) as handle:
                    handle.write(b"fake pdf bytes")
                    fake_pdf = Path(handle.name)
                try:
                    fingerprint = _fingerprint_pdf(fake_pdf)
                    cache_file = cache_dir / (
                        f"{fingerprint}_p1_dpi200_{pdf_ocr_engine._ENGINE_FINGERPRINT}.json"
                    )
                    cache_file.write_text(
                        json.dumps(
                            {
                                "engine_fingerprint": pdf_ocr_engine._ENGINE_FINGERPRINT,
                                "page_number": 1,
                                "lines": [line.to_dict() for line in cached_lines],
                            }
                        ),
                        encoding="utf-8",
                    )
                    result = ocr_pdf_page(fake_pdf, 1)
                    self.assertEqual(result, cached_lines)
                    render_mock.assert_not_called()
                finally:
                    fake_pdf.unlink(missing_ok=True)

    def test_clear_ocr_cache_no_directory_returns_zero(self):
        with tempfile.TemporaryDirectory() as tmp:
            missing_dir = Path(tmp) / "nope"
            with patch.object(pdf_ocr_engine, "_CACHE_DIR", missing_dir):
                self.assertEqual(clear_ocr_cache(), 0)

    def test_clear_ocr_cache_all_removes_files(self):
        with tempfile.TemporaryDirectory() as tmp:
            cache_dir = Path(tmp)
            (cache_dir / "abc_p1_dpi200_foo.json").write_text("{}", encoding="utf-8")
            (cache_dir / "def_p2_dpi200_foo.json").write_text("{}", encoding="utf-8")
            with patch.object(pdf_ocr_engine, "_CACHE_DIR", cache_dir):
                self.assertEqual(clear_ocr_cache(), 2)
            self.assertEqual(list(cache_dir.glob("*.json")), [])


class CachePathHelpersTest(unittest.TestCase):
    def test_cache_path_is_deterministic(self):
        path_a = _cache_path("abc123", 7, 220)
        path_b = _cache_path("abc123", 7, 220)
        self.assertEqual(path_a, path_b)
        self.assertIn("abc123", path_a.name)
        self.assertIn("p7", path_a.name)
        self.assertIn("dpi220", path_a.name)

    def test_fingerprint_uses_nanosecond_mtime(self):
        class FakeStat:
            st_size = 1234

            def __init__(self, mtime_ns: int):
                self.st_mtime_ns = mtime_ns

        class FakePdfPath:
            def __init__(self):
                self._stats = [
                    FakeStat(1_700_000_000_000_000_001),
                    FakeStat(1_700_000_000_000_000_999),
                ]

            def stat(self):
                return self._stats.pop(0)

            def resolve(self):
                return "C:/sample.pdf"

        fake_pdf = FakePdfPath()
        fingerprint_a = _fingerprint_pdf(fake_pdf)  # type: ignore[arg-type]
        fingerprint_b = _fingerprint_pdf(fake_pdf)  # type: ignore[arg-type]
        self.assertNotEqual(fingerprint_a, fingerprint_b)


if __name__ == "__main__":
    unittest.main()
