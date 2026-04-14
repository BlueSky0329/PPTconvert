import unittest

from core.pdf_exam_models import ExamQuestion, MaterialUnit, RichLine
from core.pdf_exam_parse import (
    _option_cluster_end,
    _preprocess_line_items,
    _split_into_material_units,
    parse_quant_block,
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

    def test_data_section_can_split_generic_material_headers(self):
        items = [
            ("2015 年江苏资料分析", None),
            ("材料", None),
            ("第一组材料正文", None),
            ("1.", None),
            ("第一题题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("材料", None),
            ("第二组材料正文", None),
            ("6.", None),
            ("第六题题干", None),
            ("A．5\tB．6\tC．7\tD．8", None),
        ]
        exam = parse_line_items(items, mode="data")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.data_sections[0].materials), 2)
        self.assertEqual(exam.data_sections[0].materials[0].header, "材料一")
        self.assertEqual(exam.data_sections[0].materials[1].header, "材料二")
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "1")
        self.assertEqual(exam.data_sections[0].materials[1].questions[0].source_number, "6")

    def test_parse_without_titles_can_detect_data_from_generic_material_headers(self):
        items = [
            ("材料", None),
            ("2024年某市工业增加值同比增长8.3%，服务业增加值同比增长6.1%。", None),
            ("111.", None),
            ("根据上述材料，下列说法正确的是：", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("材料", None),
            ("2024年某市服务业增加值占比进一步提高。", None),
            ("116.", None),
            ("根据上述材料，下列说法正确的是：", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.data_sections[0].materials), 2)
        self.assertEqual(exam.data_sections[0].materials[1].questions[0].source_number, "116")

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

    def test_explicit_common_sense_section_does_not_split_single_definition_like_question(self):
        items = [
            ("二. 常识判断：", None),
            ("1.", None),
            ("下列关于行政处罚的说法正确的是：", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("2.", None),
            ("某种行为是指行为人违反管理秩序并造成一定后果的行为。根据上述定义，下列属于该行为的是：", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.common_sense_sections), 1)
        self.assertEqual(len(exam.common_sense_sections[0].questions), 2)
        self.assertEqual(len(exam.reasoning_sections), 0)

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

    def test_objective_section_skips_count_and_start_prompt_lines(self):
        items = [
            ("一. 政治理论：", None),
            ("(共 15 题,参考时限 15 分钟)", None),
            ("根据题目要求,在四个选项中选出一个正确答案。", None),
            ("请开始答题:", None),
            ("1.2025 年政府工作报告提出要大力提振消费。下列说法正确的是( )。", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="politics")
        question = exam.politics_sections[0].questions[0]
        self.assertEqual(question.source_number, "1")
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertNotIn("参考时限", stem_text)
        self.assertNotIn("请开始答题", stem_text)
        self.assertIn("2025 年政府工作报告", stem_text)

    def test_year_leading_question_number_without_gap_is_recognized(self):
        items = [
            ("一. 政治理论：", None),
            ("12.2025 年我们要扎实推进重点领域改革。下列说法正确的是( )。", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="politics")
        question = exam.politics_sections[0].questions[0]
        self.assertEqual(question.source_number, "12")
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertIn("2025 年我们要扎实推进重点领域改革", stem_text)

    def test_question_number_accepts_quote_and_blank_prefix_stem(self):
        items = [
            ("三. 言语理解与表达：", None),
            ("46.                    ,严格环境执法是环境法治建设的重要内容。", None),
            ("A 法者,天下之准绳也", None),
            ("B 徒法不足以自行", None),
            ("C 立善防恶谓之礼,禁非立是谓之法", None),
            ("D 治国无其法则乱,守法而不变则衰", None),
        ]
        exam = parse_line_items(items, mode="verbal")
        question = exam.verbal_sections[0].questions[0]
        self.assertEqual(question.source_number, "46")
        first_option = "".join(text for text, _ in question.option_lines[0].parts)
        self.assertTrue(first_option.startswith("A．法者"))

        quote_exam = parse_line_items(
            [
                ("三. 言语理解与表达：", None),
                ("48.“网络开盒”行为是指不法分子通过非法手段获取特定人隐私信息。", None),
                ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ],
            mode="verbal",
        )
        self.assertEqual(quote_exam.verbal_sections[0].questions[0].source_number, "48")

    def test_question_number_accepts_numbered_intro_stem(self):
        items = [
            ("三. 言语理解与表达：", None),
            ("59.1但在数字化平台的设计理念上,有的地方仍是从单一部门的执法需求出发。", None),
            ("2对此,数字化平台建设应坚持整体化、系统化的思路。", None),
            ("A．451362\tB．512463\tC．461352\tD．541632", None),
        ]
        exam = parse_line_items(items, mode="verbal")
        question = exam.verbal_sections[0].questions[0]
        self.assertEqual(question.source_number, "59")
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertIn("1但在数字化平台", stem_text)

    def test_question_number_accepts_plain_number_plus_space_stem(self):
        items = [
            ("五. 判断推理：", None),
            ("74 把下面六个图形分为两类,使每一类图形都有各自的共同特征或规律,分类正确的一项是:", None),
            ("A．135,246\tB．123,456\tC．156,234\tD．124,356", None),
        ]
        exam = parse_line_items(items, mode="reasoning")
        question = exam.reasoning_sections[0].questions[0]
        self.assertEqual(question.source_number, "74")
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertTrue(stem_text.startswith("把下面六个图形分为两类"))

    def test_explicit_politics_section_keeps_common_sense_like_questions(self):
        items = [
            ("一. 政治理论：", None),
            ("1.", None),
            ("习近平总书记指出,要坚定不移推进高水平对外开放。下列说法正确的是：", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("2.", None),
            ("下列关于全面依法治国的说法不正确的是：", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("3.", None),
            ("下列关于周边外交的说法正确的是：", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.politics_sections), 1)
        self.assertEqual(len(exam.politics_sections[0].questions), 3)
        self.assertEqual(len(exam.common_sense_sections), 0)

    def test_numeric_stem_fragment_is_not_treated_as_other_section_title(self):
        items = [
            ("四. 数量关系：", None),
            ("61.某社区有甲、乙、丙三个快递驿站。甲的营业收入比乙多20%,丙的营业收入是甲、乙收入之和的4", None),
            ("11。乙的总成本比营业收入少2", None),
            ("A．高10 个百分点以上\tB．高不到10 个百分点\tC．低10 个百分点以上\tD．低不到10 个百分点", None),
            ("62.第二题题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual(len(exam.quant_sections[0].questions), 2)
        self.assertEqual([q.source_number for q in exam.quant_sections[0].questions], ["61", "62"])

    def test_reasoning_subsection_titles_do_not_cut_off_reasoning_questions(self):
        items = [
            ("五. 判断推理：", None),
            ("一、图形推理。请按每道题的答题要求作答。", None),
            ("71.从所给的四个选项中,选择最合适的一项。", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("二、定义判断。每道题先给出定义,然后列出四种情况。", None),
            ("选出一个最符合或最不符合该定义的答案。注意:假设这个定义是正确的,不容置疑的。", None),
            ("76.信息交合法,又可以称为“要素标的发明法”。", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="all")
        self.assertEqual(len(exam.reasoning_sections), 1)
        self.assertEqual([q.source_number for q in exam.reasoning_sections[0].questions], ["71", "76"])

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

    def test_image_options_before_blank_markers_are_rebalanced(self):
        items = [
            ("四. 数量关系：", None),
            ("65. 第一题题干", None),
            ("", "a.png"),
            ("", "b.png"),
            ("", "c.png"),
            ("", "d.png"),
            ("A.", None),
            ("B.", None),
            ("C.", None),
            ("D.", None),
        ]
        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]

        self.assertEqual([part[1] for line in question.stem_lines for part in line.parts if part[1]], [])
        self.assertEqual(len(question.option_lines), 8)
        self.assertEqual(question.option_lines[0].parts[0][1], "a.png")
        self.assertEqual(question.option_lines[1].parts[0][0], "A．")
        self.assertEqual(question.option_lines[-2].parts[0][1], "d.png")
        self.assertEqual(question.option_lines[-1].parts[0][0], "D．")

    def test_preprocess_strips_scan_ad_noise_lines(self):
        items = [
            ("题干第一行", None),
            ("扫", None),
            ("码", None),
            ("关", None),
            ("注", None),
            ("各种考试资料购买，请加微信：行测资料库", None),
            ("题干第二行", None),
        ]

        self.assertEqual(
            _preprocess_line_items(items),
            [("题干第一行", None), ("题干第二行", None)],
        )

    def test_blank_option_markers_without_text_or_images_are_ignored(self):
        items = [
            ("四. 数量关系：", None),
            ("32. 某题题干", None),
            ("A.", None),
            ("B.", None),
            ("C.", None),
            ("D.", None),
            ("", "figure.png"),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
        ]

        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]

        self.assertEqual(question.source_number, "32")
        option_texts = [
            "".join(text for text, _img in line.parts)
            for line in question.option_lines
            if any(text for text, _img in line.parts)
        ]
        self.assertEqual([text[-1] for text in option_texts[-4:]], ["1", "2", "3", "4"])
        self.assertFalse(any(text.endswith("．") or text.endswith(".") for text in option_texts[:-4]))

    def test_noise_between_stem_and_options_does_not_pollute_question(self):
        items = [
            ("四. 数量关系：", None),
            ("32. 某图示题题干第一行", None),
            ("题干第二行", None),
            ("", "figure.png"),
            ("扫", None),
            ("码", None),
            ("关", None),
            ("注", None),
            ("各种考试资料购买，请加微信：行测资料库", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
        ]

        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]
        stem_text = "".join(text for line in question.stem_lines for text, _img in line.parts)
        option_text = "".join(text for line in question.option_lines for text, _img in line.parts)

        self.assertNotIn("扫码", stem_text)
        self.assertNotIn("微信", stem_text)
        self.assertNotIn("行测资料库", option_text)

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

    def test_prompt_line_after_d_option_is_not_swallowed_into_previous_question(self):
        items = [
            ("五. 判断推理：", None),
            ("79.", None),
            ("前题题干", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁", None),
            ("根据上述定义,下列最能体现现实主义的是:", None),
            ("A．满面尘灰烟火色", None),
            ("B．飞流直下三千尺", None),
            ("C．我见青山多妩媚", None),
            ("D．遥望齐州九点烟", None),
        ]
        exam = parse_line_items(items, mode="reasoning")
        questions = exam.reasoning_sections[0].questions
        self.assertEqual([q.source_number for q in questions], ["79", ""])
        second_stem = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertIn("根据上述定义,下列最能体现现实主义的是", second_stem)

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

    def test_numeric_leading_stem_after_previous_d_option_starts_new_question(self):
        items = [
            ("六. 资料分析：", None),
            ("材料", None),
            ("某市居民收入情况如下表。", None),
            ("9.", None),
            ("上一题题干", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁", None),
            ("10. 2014 年全国城镇居民人数占全国居民人数的比重是:", None),
            ("A．36.4%", None),
            ("B．42.1%", None),
            ("C．52.7%", None),
            ("D．69.9%", None),
        ]
        exam = parse_line_items(items, mode="data")
        questions = exam.data_sections[0].materials[0].questions
        self.assertEqual(len(questions), 2)
        self.assertEqual(questions[0].source_number, "9")
        self.assertEqual(questions[1].source_number, "10")
        second_stem = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertIn("2014 年全国城镇居民人数", second_stem)

    def test_bare_single_option_lines_do_not_split_class_labels_in_quant_stem(self):
        items = [
            ("四. 数量关系：", None),
            ("67.", None),
            ("某培训学校有A、B、C 三个班,共有149 名学生,学生可以走读或者寄宿。", None),
            ("A 班总人数不超过60 人,B 班走读的学生人数占该班总人数的19", None),
            ("12,B 班总人数与C 班相差1 人。则该培训学校B 班与C 班共", None),
            ("", "stem.png"),
            ("5,当甲45,C 班走读的学生有多少名学生?", None),
            ("A.92           B.91          C.90           D.89", None),
        ]
        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]
        self.assertEqual(question.source_number, "67")
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertIn("A、B、C 三个班", stem_text)
        self.assertIn("A 班总人数不超过60 人", stem_text)
        self.assertEqual(len(question.option_lines), 2)

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

    def test_preprocess_splits_numeric_sequence_question_lines(self):
        items = [
            ("1.7,8,9,11,17,41,(", None),
            (")", None),
            ("A.8", None),
            ("6", None),
            ("B.123", None),
            ("C.161", None),
            ("D.192", None),
            ("2.-2,5,0,7,4,(", None),
            (")", None),
            ("A.8", None),
            ("B.9", None),
            ("C.12", None),
            ("D.17", None),
        ]

        processed = _preprocess_line_items(items)

        self.assertEqual(processed[0], ("1.", None))
        self.assertEqual(processed[1], ("7,8,9,11,17,41,(", None))
        self.assertIn(("A. 86", None), processed)
        self.assertIn(("2.", None), processed)

    def test_preprocess_merges_numeric_continuation_fragment(self):
        items = [
            ("57.乡村综艺虽植根乡土,主要受众却是都市人群。", None),
            ("在《种地吧》中,10 名年轻人拿起锄头、镰刀下地干活,用近", None),
            ("200.", None),
            ("天在142 亩土地上播种、灌溉、施肥。", None),
            ("这段文字主要介绍了乡村综艺的:", None),
        ]

        processed = _preprocess_line_items(items)

        texts = [text for text, _ in processed if text]
        self.assertTrue(any("用近200天在142 亩土地上播种、灌溉、施肥。" in text for text in texts))
        self.assertNotIn(("200.", None), processed)

    def test_preprocess_does_not_split_quantity_measure_text_into_question(self):
        items = [
            ("54.在一望无际的大海上,舰艇编队是如何辨别方位?", None),
            ("30 余颗卫星组成的“星座”持续播发加密定位信号。", None),
            ("下列说法与原文不符的一项是:", None),
        ]

        processed = _preprocess_line_items(items)
        texts = [text for text, _ in processed if text]

        self.assertIn("30 余颗卫星组成的“星座”持续播发加密定位信号。", texts)
        self.assertNotIn("30.", texts)

    def test_parse_quant_block_keeps_image_only_questions(self):
        items = _preprocess_line_items(
            [
                ("1.从所给的四个选项中选择最合适的一个填入问号处。", None),
                ("", "q1.png"),
                ("2.从所给的四个选项中选择最合适的一个填入问号处。", None),
                ("", "q2.png"),
                ("3.定义判断题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ]
        )

        questions = parse_quant_block(items, 0, len(items))

        self.assertEqual([question.source_number for question in questions], ["1", "2", "3"])
        self.assertEqual(len(questions[0].option_lines), 0)
        self.assertEqual(len(questions[1].option_lines), 0)
        self.assertTrue(any(img == "q1.png" for line in questions[0].stem_lines for _text, img in line.parts))
        self.assertTrue(any(img == "q2.png" for line in questions[1].stem_lines for _text, img in line.parts))

    def test_sparse_material_headers_do_not_force_objective_book_into_data(self):
        items = []
        for number in range(1, 9):
            items.extend(
                [
                    (f"{number}.根据上述定义,下列属于该概念的是:", None),
                    ("A.甲", None),
                    ("B.乙", None),
                    ("C.丙", None),
                    ("D.丁", None),
                ]
            )
        items.extend(
            [
                ("材料一:", None),
                ("9.以下哪项如果为真,最能削弱上述观点?", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ]
        )

        exam = parse_line_items(items, mode="all")

        self.assertEqual(len(exam.data_sections), 0)
        self.assertGreaterEqual(sum(len(section.questions) for section in exam.iter_objective_sections()), 9)

    def test_parse_without_titles_can_anchor_graphic_reasoning_from_question_batch(self):
        items = [
            ("1.从所给的四个选项中选择最合适的一个填入问号处,使之呈现一定的规律性:", None),
            ("", "q1.png"),
            ("2.从所给的四个选项中选择最合适的一个填入问号处,使之呈现一定的规律性:", None),
            ("", "q2.png"),
            ("3.把下面的六个图形分为两类,使每一类图形都有各自的共同特征或规律,分类正确的一项是:", None),
            ("", "q3.png"),
            ("A.135,246", None),
            ("B.136,245", None),
            ("C.123,456", None),
            ("D.145,236", None),
            ("4.左图给定的是纸盒的外表面,右边哪项能由它折叠而成?", None),
            ("", "q4.png"),
        ]

        exam = parse_line_items(items, mode="all")

        self.assertEqual(len(exam.reasoning_sections), 1)
        self.assertEqual([q.source_number for q in exam.reasoning_sections[0].questions], ["1", "2", "3", "4"])

    def test_filename_single_subject_book_anchors_reasoning_strategy(self):
        items = []
        for number in range(1, 5):
            items.extend(
                [
                    (f"{number}.根据上述定义,下列符合定义的是:", None),
                    ("A.甲", None),
                    ("B.乙", None),
                    ("C.丙", None),
                    ("D.丁", None),
                ]
            )
        items.extend(
            [
                ("材料一", None),
                ("5.以下哪项如果为真,最能削弱上述观点?", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ]
        )

        exam = parse_line_items(
            items,
            mode="all",
            source_name="行测——判断推理（2000题）.pdf",
        )

        self.assertEqual(len(exam.data_sections), 0)
        self.assertEqual(len(exam.reasoning_sections), 1)
        self.assertEqual(
            [question.source_number for question in exam.reasoning_sections[0].questions],
            ["1", "2", "3", "4", "5"],
        )

    def test_filename_set_paper_prefers_grouped_multi_section_strategy(self):
        items = [
            ("1.这段文字主要说明的是:", None),
            ("A.基层治理更强调精细治理", None),
            ("B.财政投入最重要", None),
            ("C.市场机制应完全主导", None),
            ("D.居民参与会降低效率", None),
            ("2.依次填入画横线部分最恰当的一项是:", None),
            ("A.精细", None),
            ("B.精确", None),
            ("C.精致", None),
            ("D.精准", None),
            ("61.某商品按8折出售后利润率为20%,其成本是多少?", None),
            ("A.80", None),
            ("B.96", None),
            ("C.100", None),
            ("D.120", None),
            ("62.甲、乙两车分别以每小时60千米和80千米的速度相向而行,几小时后相遇?", None),
            ("A.4", None),
            ("B.5", None),
            ("C.6", None),
            ("D.8", None),
        ]

        exam = parse_line_items(
            items,
            mode="all",
            source_name="模拟卷十一.pdf",
        )

        self.assertEqual(len(exam.verbal_sections), 1)
        self.assertEqual(len(exam.quant_sections), 1)
        self.assertEqual([q.source_number for q in exam.verbal_sections[0].questions], ["1", "2"])
        self.assertEqual([q.source_number for q in exam.quant_sections[0].questions], ["61", "62"])

    def test_filename_set_paper_uses_standard_ranges_and_material_tail(self):
        items = [
            ("1.关于推进中国式现代化的相关表述,正确的是:", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("16.根据我国民法典相关规定,下列说法正确的是:", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("26.这段文字主要说明的是:", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("61.某商品按8折出售后利润率为20%,其成本是多少?", None),
            ("A.80", None),
            ("B.96", None),
            ("C.100", None),
            ("D.120", None),
            ("71.如果甲成立,那么乙成立。以下哪项最能削弱上述论证?", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("材料一", None),
            ("2024年某市固定资产投资额同比增长8.3%。", None),
            ("101.", None),
            ("根据上述材料,下列说法正确的是:", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        exam = parse_line_items(
            items,
            mode="all",
            source_name="模拟卷十一.pdf",
        )

        self.assertEqual([q.source_number for q in exam.politics_sections[0].questions], ["1"])
        self.assertEqual([q.source_number for q in exam.common_sense_sections[0].questions], ["16"])
        self.assertEqual([q.source_number for q in exam.verbal_sections[0].questions], ["26"])
        self.assertEqual([q.source_number for q in exam.quant_sections[0].questions], ["61"])
        self.assertEqual([q.source_number for q in exam.reasoning_sections[0].questions], ["71"])
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "101")


if __name__ == "__main__":
    unittest.main()
