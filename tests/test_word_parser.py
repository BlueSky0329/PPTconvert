import os
import tempfile
import unittest

from docx import Document

from core.word_parser import (
    WordParser,
    _match_question_start,
    _normalize_answer,
    _parse_options_from_line,
)


class WordParserHelpersTest(unittest.TestCase):
    def test_match_question_start_supports_multiple_formats(self):
        self.assertEqual(
            _match_question_start("\uff082020\u00b7\u4e0a\u6d77\uff09\u9898\u5e72"),
            ("\uff082020\u00b7\u4e0a\u6d77\uff09", "\u9898\u5e72"),
        )
        self.assertEqual(
            _match_question_start("1. simple stem"),
            (None, "simple stem"),
        )
        self.assertEqual(
            _match_question_start("2\u3001another stem"),
            (None, "another stem"),
        )
        self.assertEqual(
            _match_question_start("\u7b2c3\u9898 \u4e2d\u6587\u9898\u5e72"),
            (None, "\u4e2d\u6587\u9898\u5e72"),
        )
        self.assertIsNone(_match_question_start("2024. not a supported plain question prefix"))

    def test_parse_options_from_line_supports_multiple_markers(self):
        options = _parse_options_from_line(
            "A. alpha B. beta C. gamma E. epsilon"
        )
        self.assertEqual([option.letter for option in options], ["A", "B", "C", "E"])
        self.assertEqual(options[-1].text, "epsilon")

    def test_normalize_answer_deduplicates_letters(self):
        self.assertEqual(_normalize_answer("A, C / A"), "AC")


class WordParserIntegrationTest(unittest.TestCase):
    def test_parse_docx_with_numeric_and_chinese_question_prefixes(self):
        document = Document()
        document.add_paragraph("1. First stem")
        document.add_paragraph("A. Alpha")
        document.add_paragraph("B. Beta")
        document.add_paragraph("\u7b2c2\u9898 Second stem")
        document.add_paragraph("A. One")
        document.add_paragraph("B. Two")
        document.add_paragraph("\uff082020\u00b7\u4e0a\u6d77\uff09Third stem")
        document.add_paragraph("A. Three")
        document.add_paragraph("B. Four")

        tmp_path = None
        parser = WordParser()
        try:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name
            document.save(tmp_path)

            questions = parser.parse(tmp_path)
            self.assertEqual(len(questions), 3)
            self.assertEqual(questions[0].stem, "First stem")
            self.assertEqual(questions[1].stem, "Second stem")
            self.assertEqual(questions[2].stem, "Third stem")
            self.assertEqual(questions[2].source_label, "\uff082020\u00b7\u4e0a\u6d77\uff09")
            self.assertEqual([opt.letter for opt in questions[0].options], ["A", "B"])
        finally:
            parser.cleanup()
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)
