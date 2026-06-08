# -*- coding: utf-8 -*-
import os
import tempfile
import unittest


class OcrIngestGatingTest(unittest.TestCase):
    def setUp(self):
        try:
            import fitz  # noqa: F401
        except Exception:
            self.skipTest("PyMuPDF 不可用")

    def _make_pdf(self, with_text):
        import fitz

        directory = tempfile.mkdtemp()
        path = os.path.join(directory, "sample.pdf")
        doc = fitz.open()
        for _ in range(3):
            page = doc.new_page()
            if with_text:
                page.insert_text(
                    (72, 72),
                    "This is a page that carries a real text layer used to verify "
                    "that scanned-PDF detection does not misfire on born-digital files. " * 2,
                )
        doc.save(path)
        doc.close()
        return path

    def test_text_pdf_routes_native(self):
        from ingest.pdf.ocr_ingest import document_scanned_ratio, looks_scanned

        path = self._make_pdf(with_text=True)
        ratio, pages = document_scanned_ratio(path)
        self.assertEqual(pages, 3)
        self.assertEqual(ratio, 0.0)
        self.assertFalse(looks_scanned(path))

    def test_textless_pdf_routes_ocr(self):
        from ingest.pdf.ocr_ingest import document_scanned_ratio, looks_scanned

        path = self._make_pdf(with_text=False)
        ratio, pages = document_scanned_ratio(path)
        self.assertEqual(pages, 3)
        self.assertEqual(ratio, 1.0)
        self.assertTrue(looks_scanned(path))

    def test_unreadable_pdf_does_not_raise(self):
        from ingest.pdf.ocr_ingest import document_scanned_ratio, looks_scanned

        directory = tempfile.mkdtemp()
        path = os.path.join(directory, "broken.pdf")
        with open(path, "wb") as handle:
            handle.write(b"not a real pdf")
        self.assertEqual(document_scanned_ratio(path), (0.0, 0))
        self.assertFalse(looks_scanned(path))


if __name__ == "__main__":
    unittest.main()
