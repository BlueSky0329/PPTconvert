# -*- coding: utf-8 -*-
import unittest

from core.pdf_exam_parse import (
    _detect_subject_section_kind,
    _is_boilerplate_line,
    _is_subject_section_title,
)


class CountSuffixSectionTitleTest(unittest.TestCase):
    """老试卷的「科目名(共N 题，参考时限N 分钟)」式篇题应被识别为篇题，
    而不是被当成 boilerplate 跳过（曾导致整份 0 题）。"""

    def test_count_suffix_header_recognized_as_title(self):
        line = "数量关系(共15 题，参考时限15 分钟)"
        self.assertTrue(_is_subject_section_title(line, "quant"))
        self.assertEqual(_detect_subject_section_kind(line), "quant")

    def test_count_suffix_header_is_not_boilerplate(self):
        self.assertFalse(_is_boilerplate_line("数量关系(共15 题，参考时限15 分钟)"))
        self.assertFalse(_is_boilerplate_line("言语理解与表达(共40 题，参考时限35 分钟)"))

    def test_pure_descriptor_without_subject_stays_boilerplate(self):
        # 没有科目名的纯说明行仍应按 boilerplate 跳过
        self.assertTrue(_is_boilerplate_line("(共15 题，参考时限15 分钟)"))
        self.assertIsNone(_detect_subject_section_kind("(共15 题，参考时限15 分钟)"))

    def test_count_suffix_across_subjects(self):
        for line, kind in [
            ("言语理解与表达(共40 题，参考时限35 分钟)", "verbal"),
            ("判断推理(共35 题，参考时限35 分钟)", "reasoning"),
            ("常识判断(共20 题，参考时限10 分钟)", "common_sense"),
            ("资料分析(共20 题，参考时限25 分钟)", "data"),
        ]:
            self.assertEqual(_detect_subject_section_kind(line), kind, line)


if __name__ == "__main__":
    unittest.main()
