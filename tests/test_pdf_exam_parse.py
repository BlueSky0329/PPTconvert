import unittest

from core.pdf_exam_models import ExamQuestion, MaterialUnit, RichLine
from core.pdf_exam_parse import (
    _option_cluster_end,
    _split_into_material_units,
    parse_line_items,
    parse_material_block,
)


class TestPdfExamParse(unittest.TestCase):
    def test_parse_all_six_subject_sections(self):
        items = [
            ("一. 政治理论：", None),
            ("政治题干", None),
            ("1.", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("二. 常识判断：", None),
            ("常识题干", None),
            ("11.", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("三. 言语理解与表达：", None),
            ("言语题干", None),
            ("21.", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("四. 数量关系：", None),
            ("数量题干", None),
            ("66.", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("五. 判断推理：", None),
            ("推理题干", None),
            ("76.", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("六. 资料分析：", None),
            ("材料一", None),
            ("材料正文", None),
            ("111.", None),
            ("资料题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.politics_sections), 1)
        self.assertEqual(len(exam.common_sense_sections), 1)
        self.assertEqual(len(exam.verbal_sections), 1)
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.reasoning_sections), 1)
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(exam.politics_sections[0].questions[0].source_number, "1")
        self.assertEqual(exam.reasoning_sections[0].questions[0].source_number, "76")
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "111")

    def test_stem_not_merged_with_options(self):
        items = [
            ("2025年销量多少：", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        self.assertIsNone(_option_cluster_end(items, 0, len(items)))
        self.assertIsNotNone(_option_cluster_end(items, 1, len(items)))

    def test_material_intro_stem(self):
        items = [
            ("2026年·天津·资料分析", None),
            ("材料一", None),
            ("这是材料段落。", None),
            ("2025年销量多少：", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        u = parse_material_block(items, 1, len(items))
        assert u is not None
        self.assertEqual(len(u.intro_lines), 1)
        self.assertEqual(len(u.questions), 1)

    def test_parse_sections(self):
        items = [
            ("2026年·天津·资料分析", None),
            ("材料一", None),
            ("材料正文", None),
            ("题干？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("2026年·天津·数量关系", None),
            ("第一题？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="both")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.quant_sections), 1)

    def test_two_line_section_title(self):
        items = [
            ("2026年·天津·", None),
            ("资料分析", None),
            ("材料一", None),
            ("材料", None),
            ("问？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="both")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertIn("资料分析", exam.data_sections[0].title)

    def test_fullwidth_year_title(self):
        items = [
            ("２０２６年·天津·资料分析", None),
            ("材料一", None),
            ("材", None),
            ("题？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="data")
        self.assertEqual(len(exam.data_sections), 1)

    def test_part_section_title(self):
        items = [
            ("第二部分 资料分析", None),
            ("材料一", None),
            ("材", None),
            ("题？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="data")
        self.assertEqual(len(exam.data_sections), 1)

    def test_fullwidth_option_letters(self):
        items = [
            ("2026年·天津·资料分析", None),
            ("材料一", None),
            ("材料正文", None),
            ("问？", None),
            ("Ａ．1\tＢ．2\tＣ．3\tＤ．4", None),
        ]
        exam = parse_line_items(items, mode="data")
        self.assertEqual(len(exam.data_sections[0].materials[0].questions), 1)

    def test_no_material_row_fallback(self):
        """无「材料一」行时整段按一篇解析。"""
        items = [
            ("2026年·天津·资料分析", None),
            ("某段材料", None),
            ("题？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="data")
        self.assertEqual(len(exam.data_sections[0].materials), 1)
        self.assertGreater(len(exam.data_sections[0].materials[0].questions), 0)

    def test_outline_style_section_titles(self):
        """四. 数量关系 / 六. 资料分析 大纲式篇题。"""
        items = [
            ("四. 数量关系：", None),
            ("在这部分试题中，每道题呈现一段表述数字关系的文字。", None),
            ("问？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("六. 资料分析：", None),
            ("所给出的图、表、文字或综合性资料均有若干个问题要你回答。", None),
            ("材", None),
            ("题？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="both")
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.data_sections), 1)
        self.assertGreater(len(exam.quant_sections[0].questions), 0)
        self.assertGreater(len(exam.data_sections[0].materials[0].questions), 0)

    def test_parse_without_titles_falls_back_to_quant_by_heuristic(self):
        items = [
            ("甲、乙两地相距240千米，客车和货车同时出发，相向而行，几小时后相遇？", None),
            ("66.", None),
            ("A．4\tB．5\tC．6\tD．8", None),
            ("某商品按8折出售后利润率为20%，其成本是多少？", None),
            ("67.", None),
            ("A．80\tB．96\tC．100\tD．120", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.quant_sections[0].questions), 2)

    def test_parse_without_titles_can_force_data_subject(self):
        items = [
            ("材料一", None),
            ("2024年某市工业增加值同比增长8.3%，服务业增加值同比增长6.1%。", None),
            ("111.", None),
            ("根据上述材料，下列说法正确的是：", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="all", document_subject_hint="data")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.data_sections[0].materials), 1)
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "111")

    def test_partial_missing_section_title_can_split_to_reasoning(self):
        items = [
            ("四. 数量关系：", None),
            ("甲、乙两队合修一段公路，若甲单独修需要12天，乙单独修需要18天，两队合修几天完成？", None),
            ("66.", None),
            ("A．6\tB．7\tC．8\tD．9", None),
            ("如果所有甲都是乙，且有些乙是丙，那么下列哪项一定为真？", None),
            ("76.", None),
            ("A．有些甲是丙\tB．有些丙是甲\tC．有些乙不是丙\tD．所有甲都是乙", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.quant_sections[0].questions), 1)
        self.assertEqual(len(exam.reasoning_sections), 1)
        self.assertEqual(len(exam.reasoning_sections[0].questions), 1)

    def test_objective_section_can_split_out_embedded_data_material(self):
        items = [
            ("四. 数量关系：", None),
            ("甲、乙两地相距240千米，两车相向而行几小时后相遇？", None),
            ("66.", None),
            ("A．4\tB．5\tC．6\tD．8", None),
            ("材料一", None),
            ("2024年某市工业增加值同比增长8.3%，服务业增加值同比增长6.1%。", None),
            ("111.", None),
            ("根据上述材料，下列说法正确的是：", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.quant_sections[0].questions), 1)
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "111")

    def test_other_section_stops_quant_block(self):
        items = [
            ("四. 数量关系：", None),
            ("第一题题干", None),
            ("66.", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("五. 判断推理：", None),
            ("判断题干", None),
            ("76.", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("六. 资料分析：", None),
            ("材料一", None),
            ("材料正文", None),
            ("111.", None),
            ("资料题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="both")
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.quant_sections[0].questions), 1)
        self.assertEqual(len(exam.data_sections), 1)

    def test_question_number_line_removed_from_stem(self):
        items = [
            ("四. 数量关系：", None),
            ("第一题题干", None),
            ("66.", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="quant")
        q = exam.quant_sections[0].questions[0]
        stem_texts = ["".join(text for text, _ in rl.parts) for rl in q.stem_lines]
        self.assertEqual(stem_texts, ["第一题题干"])
        self.assertEqual(q.source_number, "66")

    def test_objective_section_boilerplate_continuation_not_merged_into_first_question(self):
        items = [
            ("三. 言语理解与表达：", None),
            ("本部分包括表达与理解两方面的内容。请根据题目要求,在四个选项中选出一个最", None),
            ("恰当的答案。", None),
            ("36.", None),
            ("第一题题干", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="verbal")
        question = exam.verbal_sections[0].questions[0]
        stem_texts = ["".join(text for text, _ in rl.parts) for rl in question.stem_lines]
        self.assertEqual(stem_texts, ["第一题题干"])
        self.assertEqual(question.source_number, "36")

    def test_material_split_keeps_new_intro(self):
        items = [("六. 资料分析：", None)]
        for group_index, base in enumerate((111, 116, 121, 126), 1):
            items.append((f"第{group_index}组材料说明", None))
            for offset in range(5):
                qno = base + offset
                items.append((f"{qno}.", None))
                items.append((f"第{qno}题题干", None))
                items.append(("A．1\tB．2\tC．3\tD．4", None))
        exam = parse_line_items(items, mode="data")
        self.assertEqual(len(exam.data_sections[0].materials), 4)
        self.assertGreater(len(exam.data_sections[0].materials[1].intro_lines), 0)
        self.assertEqual(exam.data_sections[0].materials[1].questions[0].source_number, "116")

    def test_image_options_are_preserved(self):
        items = [
            ("四. 数量关系：", None),
            ("看图选择", None),
            ("66.", None),
            ("A.", None),
            ("", "a.png"),
            ("B.", None),
            ("", "b.png"),
            ("C.", None),
            ("", "c.png"),
            ("D.", None),
            ("", "d.png"),
        ]
        exam = parse_line_items(items, mode="quant")
        option_lines = exam.quant_sections[0].questions[0].option_lines
        self.assertEqual(len(option_lines), 8)
        self.assertEqual(option_lines[0].parts[0][0], "A．")
        self.assertEqual(option_lines[1].parts[0][1], "a.png")
        self.assertEqual(option_lines[-1].parts[0][1], "d.png")

    def test_option_cluster_supports_three_plus_one_lines(self):
        items = [
            ("四. 数量关系：", None),
            ("66. 第一题题干", None),
            ("A．1\tB．2\tC．3", None),
            ("D．4", None),
            ("67. 第二题题干", None),
            ("A．5\tB．6\tC．7\tD．8", None),
        ]
        exam = parse_line_items(items, mode="quant")
        self.assertEqual(len(exam.quant_sections[0].questions), 2)
        self.assertEqual(exam.quant_sections[0].questions[0].source_number, "66")
        self.assertEqual(exam.quant_sections[0].questions[1].source_number, "67")

    def test_d_option_continuation_not_leaked_to_next_stem(self):
        items = [
            ("四. 数量关系：", None),
            ("66.", None),
            ("第一题题干", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁", None),
            ("D项续行说明", None),
            ("67. 第二题题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions
        self.assertEqual(len(questions), 2)
        first_option_text = "".join(text for text, _img in questions[0].option_lines[-1].parts)
        self.assertIn("D项续行说明", first_option_text)
        second_stem = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertNotIn("D项续行说明", second_stem)
        self.assertEqual(questions[1].source_number, "67")

    def test_inline_question_transition_after_d_option(self):
        items = [
            ("四. 数量关系：", None),
            ("66.", None),
            ("第一题题干", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁 67. 第二题题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions
        self.assertEqual(len(questions), 2)
        first_option_text = "".join(text for text, _img in questions[0].option_lines[-1].parts)
        self.assertEqual(first_option_text, "D．丁")
        second_stem = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertIn("第二题题干", second_stem)
        self.assertEqual(questions[1].source_number, "67")

    def test_split_twenty_questions_four_by_five(self):
        """20 题拆成四组，每组 5 题。"""
        stub_q = ExamQuestion(stem_lines=[RichLine(parts=[("x", None)])], option_lines=[])
        u = MaterialUnit(header="材料一", intro_lines=[], questions=[stub_q] * 20)
        parts = _split_into_material_units(u)
        self.assertEqual(len(parts), 4)
        self.assertEqual(len(parts[0].questions), 5)
        self.assertEqual(len(parts[3].questions), 5)
        self.assertEqual(parts[0].header, "材料一")
        self.assertEqual(parts[3].header, "材料四")

    def test_material_split_recovers_intro_spilled_into_previous_d_option(self):
        items = [("六. 资料分析：", None), ("", "mat1.png")]
        for qno in range(111, 116):
            items.append((f"{qno}.", None))
            items.append((f"第{qno}题题干", None))
            if qno < 115:
                items.append(("A．1\tB．2\tC．3\tD．4", None))
            else:
                items.extend(
                    [
                        ("A．甲", None),
                        ("B．乙", None),
                        ("C．丙", None),
                        ("D．丁", None),
                        ("第二组材料说明第一行", None),
                        ("第二组材料说明第二行", None),
                        ("", "mat2a.png"),
                        ("", "mat2b.png"),
                    ]
                )
        for qno in range(116, 121):
            items.append((f"{qno}.", None))
            items.append((f"第{qno}题题干", None))
            items.append(("A．1\tB．2\tC．3\tD．4", None))

        exam = parse_line_items(items, mode="data")

        self.assertEqual(len(exam.data_sections[0].materials), 2)
        first_last = exam.data_sections[0].materials[0].questions[-1]
        first_option_tail = [
            "".join(text for text, _img in line.parts)
            for line in first_last.option_lines
        ]
        self.assertNotIn("第二组材料说明第一行", first_option_tail)
        second_material = exam.data_sections[0].materials[1]
        second_intro_texts = [
            "".join(text for text, _img in line.parts)
            for line in second_material.intro_lines
            if any((text or "").strip() for text, _img in line.parts)
        ]
        second_intro_imgs = [
            img
            for line in second_material.intro_lines
            for _text, img in line.parts
            if img
        ]
        self.assertIn("第二组材料说明第一行", second_intro_texts)
        self.assertEqual(second_intro_imgs, ["mat2a.png", "mat2b.png"])
        self.assertEqual(second_material.questions[0].source_number, "116")


if __name__ == "__main__":
    unittest.main()
