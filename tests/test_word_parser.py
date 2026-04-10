import os
import tempfile
import unittest

from docx import Document

from core.word_parser import (
    WordParser,
    _extract_question_number,
    _material_header_from_text,
    _match_question_start,
    _normalize_answer,
    _parse_options_from_line,
    _section_kind_from_text,
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

    def test_material_and_section_detection(self):
        self.assertEqual(_section_kind_from_text("2026年·天津·资料分析"), "data")
        self.assertEqual(_section_kind_from_text("四. 数量关系:"), "quant")
        self.assertEqual(_material_header_from_text("材料一"), "材料一")
        self.assertIsNone(_material_header_from_text("题干正文"))
        self.assertEqual(_extract_question_number("12. 题干"), "12")


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

    def test_data_material_is_prefixed_to_each_question(self):
        document = Document()
        document.add_paragraph("2026年·天津·资料分析")
        document.add_paragraph("材料一")
        document.add_paragraph("这是材料第一段。")
        document.add_paragraph("这是材料第二段。")
        document.add_paragraph("111. 第一题")
        document.add_paragraph("A. Alpha")
        document.add_paragraph("B. Beta")
        document.add_paragraph("112. 第二题")
        document.add_paragraph("A. Gamma")
        document.add_paragraph("B. Delta")

        tmp_path = None
        parser = WordParser()
        try:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name
            document.save(tmp_path)

            questions = parser.parse(tmp_path)
            self.assertEqual(len(questions), 2)
            self.assertEqual(questions[0].stem, "第一题")
            self.assertEqual(questions[1].stem, "第二题")
            self.assertEqual(questions[0].source_question_number, "111")
            self.assertEqual(questions[1].source_question_number, "112")
            self.assertEqual(questions[0].material_header, "材料一")
            self.assertEqual(questions[0].material_text, "这是材料第一段。\n这是材料第二段。")
        finally:
            parser.cleanup()
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)

    def test_data_material_without_question_numbers_can_still_parse(self):
        document = Document()
        document.add_paragraph("2026年·天津·资料分析")
        document.add_paragraph("材料一")
        document.add_paragraph("这是材料正文。")
        document.add_paragraph("第一道题题干")
        document.add_paragraph("A. Alpha")
        document.add_paragraph("B. Beta")
        document.add_paragraph("第二道题题干")
        document.add_paragraph("A. Gamma")
        document.add_paragraph("B. Delta")

        tmp_path = None
        parser = WordParser()
        try:
            with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as tmp:
                tmp_path = tmp.name
            document.save(tmp_path)

            questions = parser.parse(tmp_path)
            self.assertEqual(len(questions), 2)
            self.assertEqual(questions[0].stem, "第一道题题干")
            self.assertEqual(questions[1].stem, "第二道题题干")
            self.assertEqual(questions[0].material_header, "材料一")
            self.assertEqual(questions[0].material_text, "这是材料正文。")
            self.assertEqual([opt.letter for opt in questions[0].options], ["A", "B"])
        finally:
            parser.cleanup()
            if tmp_path and os.path.exists(tmp_path):
                os.unlink(tmp_path)
