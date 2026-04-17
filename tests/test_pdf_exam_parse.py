import unittest

from core.pdf_exam_models import ExamQuestion, MaterialUnit, RichLine
from core.pdf_exam_parse import (
    _parse_options_line,
    _repartition_strict_five_question_materials,
    _repair_local_question_number_anomalies,
    _realign_adjacent_data_material_intros,
    _option_cluster_end,
    _preprocess_line_items,
    _split_into_material_units,
    parse_quant_block,
    parse_line_items,
    parse_material_block,
)


class TestPdfExamParse(unittest.TestCase):
    def test_parse_options_line_ignores_embedded_pdf_markers(self):
        parsed = _parse_options_line(
            "A.农贸市场售卖的富硒土豆，深受消费者欢迎 "
            "B.超市中热卖的药酒，能治疗老年人风湿性关节炎 "
            "C.用于补充婴幼儿维生素D、促进钙吸收的咀嚼片 "
            "D.李某从某美容整形医院购买的减肥茶，喝后腹泻"
        )
        assert parsed is not None
        self.assertEqual(
            [(option.letter, option.text) for option in parsed],
            [
                ("A", "农贸市场售卖的富硒土豆，深受消费者欢迎"),
                ("B", "超市中热卖的药酒，能治疗老年人风湿性关节炎"),
                ("C", "用于补充婴幼儿维生素D、促进钙吸收的咀嚼片"),
                ("D", "李某从某美容整形医院购买的减肥茶，喝后腹泻"),
            ],
        )

    def test_parse_options_line_keeps_roman_numeral_combination_choices(self):
        parsed = _parse_options_line("A、I、II B、I、III C、II、III D、I、II、III")
        assert parsed is not None
        self.assertEqual(
            [(option.letter, option.text) for option in parsed],
            [
                ("A", "I、II"),
                ("B", "I、III"),
                ("C", "II、III"),
                ("D", "I、II、III"),
            ],
        )

    def test_parse_options_line_recovers_trailing_cd_pair(self):
        parsed = _parse_options_line("C.未有这三朝的陶瓷器物存世D.史料记载语焉不详")
        assert parsed is not None
        self.assertEqual(
            [(option.letter, option.text) for option in parsed],
            [
                ("C", "未有这三朝的陶瓷器物存世"),
                ("D", "史料记载语焉不详"),
            ],
        )

    def test_parse_options_line_keeps_embedded_vitamin_d_inside_single_option(self):
        parsed = _parse_options_line("C.用于补充婴幼儿维生素D、促进钙吸收的咀嚼片")
        assert parsed is not None
        self.assertEqual(
            [(option.letter, option.text) for option in parsed],
            [("C", "用于补充婴幼儿维生素D、促进钙吸收的咀嚼片")],
        )

    def test_parse_options_line_ignores_spaced_embedded_markers(self):
        parsed = _parse_options_line(
            "A.农贸市场售卖的富硒土豆，深受消费者欢迎 "
            "B.超市中热卖的药酒，能治疗老年人风湿性关节炎 "
            "C.用于补充婴幼儿维生素 D、促进钙吸收的咀嚼片 "
            "D.李某从某美容整形医院购买的减肥茶，喝后腹泻"
        )
        assert parsed is not None
        self.assertEqual(
            [(option.letter, option.text) for option in parsed],
            [
                ("A", "农贸市场售卖的富硒土豆，深受消费者欢迎"),
                ("B", "超市中热卖的药酒，能治疗老年人风湿性关节炎"),
                ("C", "用于补充婴幼儿维生素 D、促进钙吸收的咀嚼片"),
                ("D", "李某从某美容整形医院购买的减肥茶，喝后腹泻"),
            ],
        )

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
        exam = parse_line_items(items, mode="all", document_subject_hint="data")
        self.assertEqual(len(exam.data_sections), 1)

    def test_part_section_title(self):
        items = [
            ("第二部分 资料分析", None),
            ("材料一", None),
            ("材", None),
            ("题？", None),
            ("A．1\tB．2\tC．3\tD．4", None),
        ]
        exam = parse_line_items(items, mode="all", document_subject_hint="data")
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

    def test_data_section_can_split_parenthesized_material_markers(self):
        items = [
            ("（一", None),
            ("）", None),
            ("第一组材料正文", None),
            ("1.", None),
            ("第一题题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("(二)", None),
            ("第二组材料正文", None),
            ("6.", None),
            ("第六题题干", None),
            ("A．5\tB．6\tC．7\tD．8", None),
        ]
        exam = parse_line_items(items, mode="all", document_subject_hint="data")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.data_sections[0].materials), 2)
        self.assertEqual(exam.data_sections[0].materials[0].header, "材料一")
        self.assertEqual(exam.data_sections[0].materials[1].header, "材料二")
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "1")
        self.assertEqual(exam.data_sections[0].materials[1].questions[0].source_number, "6")

    def test_parse_without_titles_ignores_internal_other_outline_title(self):
        items = [
            ("(一)", None),
            ("第一组材料正文", None),
            ("1.", None),
            ("第一题题干", None),
            ("A．1\tB．2\tC．3\tD．4", None),
            ("(二)", None),
            ("三、债券市场对外开放情况", None),
            ("第二组材料正文", None),
            ("6.", None),
            ("第六题题干", None),
            ("A．5\tB．6\tC．7\tD．8", None),
        ]
        exam = parse_line_items(items, mode="all", source_name="行测——资料分析（950题）.pdf")
        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.data_sections[0].materials), 2)
        self.assertEqual(exam.data_sections[0].materials[0].questions[0].source_number, "1")
        self.assertEqual(exam.data_sections[0].materials[1].questions[0].source_number, "6")

    def test_local_question_number_repair_preserves_large_number_jump(self):
        questions = [
            ExamQuestion(stem_lines=[RichLine(parts=[("前题", None)])], option_lines=[], source_number="76"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常起点", None)])], option_lines=[], source_number="147"),
            ExamQuestion(stem_lines=[RichLine(parts=[("中间题1", None)])], option_lines=[], source_number="148"),
            ExamQuestion(stem_lines=[RichLine(parts=[("中间题2", None)])], option_lines=[], source_number="149"),
            ExamQuestion(stem_lines=[RichLine(parts=[("中间题3", None)])], option_lines=[], source_number="150"),
            ExamQuestion(stem_lines=[RichLine(parts=[("中间题4", None)])], option_lines=[], source_number="151"),
            ExamQuestion(stem_lines=[RichLine(parts=[("中间题5", None)])], option_lines=[], source_number="152"),
            ExamQuestion(stem_lines=[RichLine(parts=[("中间题6", None)])], option_lines=[], source_number="153"),
            ExamQuestion(stem_lines=[RichLine(parts=[("后题", None)])], option_lines=[], source_number="589"),
        ]

        _repair_local_question_number_anomalies(questions)

        self.assertEqual(
            [question.source_number for question in questions],
            ["76", "147", "148", "149", "150", "151", "152", "153", "589"],
        )

    def test_local_question_number_repair_still_fixes_small_window(self):
        questions = [
            ExamQuestion(stem_lines=[RichLine(parts=[("前题", None)])], option_lines=[], source_number="12"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常1", None)])], option_lines=[], source_number="3"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常2", None)])], option_lines=[], source_number="4"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常3", None)])], option_lines=[], source_number="5"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常4", None)])], option_lines=[], source_number="6"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常5", None)])], option_lines=[], source_number="7"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常6", None)])], option_lines=[], source_number="8"),
            ExamQuestion(stem_lines=[RichLine(parts=[("异常7", None)])], option_lines=[], source_number="9"),
            ExamQuestion(stem_lines=[RichLine(parts=[("后题", None)])], option_lines=[], source_number="20"),
        ]

        _repair_local_question_number_anomalies(questions)

        self.assertEqual(
            [question.source_number for question in questions],
            ["12", "13", "14", "15", "16", "17", "18", "19", "20"],
        )

    def test_single_subject_data_book_repartitions_headerless_five_question_materials(self):
        items = [
            ("(四)", None),
            ("第一组材料正文", None),
            ("1.", None),
            ("第一题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("2.", None),
            ("第二题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("3.", None),
            ("第三题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("4.", None),
            ("第四题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("5.", None),
            ("第五题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("第二组材料正文", None),
            ("6.", None),
            ("第六题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("7.", None),
            ("第七题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("8.", None),
            ("第八题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("9.", None),
            ("第九题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
            ("10.", None),
            ("第十题题干", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
        ]

        exam = parse_line_items(
            items,
            mode="all",
            source_name="行测——资料分析（950题）.pdf",
        )

        self.assertEqual(len(exam.data_sections), 1)
        self.assertEqual(len(exam.data_sections[0].materials), 2)
        self.assertEqual(
            [question.source_number for question in exam.data_sections[0].materials[0].questions],
            ["1", "2", "3", "4", "5"],
        )
        self.assertEqual(
            [question.source_number for question in exam.data_sections[0].materials[1].questions],
            ["6", "7", "8", "9", "10"],
        )
        self.assertIn("回答1-5题", exam.data_sections[0].materials[0].header)
        self.assertIn("回答6-10题", exam.data_sections[0].materials[1].header)

    def test_repartition_strict_five_question_materials_splits_intro_blocks(self):
        questions = [
            ExamQuestion(
                stem_lines=[RichLine(parts=[(f"{index}题题干", None)])],
                option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                source_number=str(index),
            )
            for index in range(121, 131)
        ]
        materials = _repartition_strict_five_question_materials(
            [
                MaterialUnit(
                    header="材料一",
                    intro_lines=[
                        RichLine(parts=[("第一组材料正文", None)]),
                        RichLine(parts=[("", "intro1.png")]),
                        RichLine(parts=[("第二组材料正文", None)]),
                        RichLine(parts=[("", "intro2.png")]),
                    ],
                    questions=questions,
                )
            ]
        )

        self.assertEqual(len(materials), 2)
        self.assertEqual([question.source_number for question in materials[0].questions], ["121", "122", "123", "124", "125"])
        self.assertEqual([question.source_number for question in materials[1].questions], ["126", "127", "128", "129", "130"])
        self.assertIn("第一组材料正文", "".join(text for line in materials[0].intro_lines for text, _ in line.parts))
        self.assertIn("第二组材料正文", "".join(text for line in materials[1].intro_lines for text, _ in line.parts))

    def test_repartition_strict_five_question_materials_recovers_cross_unit_intro_options(self):
        leading_questions = [
            ExamQuestion(
                stem_lines=[RichLine(parts=[(f"{index}题题干", None)])],
                option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                source_number=str(index),
            )
            for index in range(746, 751)
        ]
        trailing_questions = [
            ExamQuestion(
                stem_lines=[RichLine(parts=[(f"{index}题题干", None)])],
                option_lines=[],
                source_number=str(index),
            )
            for index in (751, 752, 753)
        ]
        trailing_questions.append(
            ExamQuestion(
                stem_lines=[RichLine(parts=[("755题题干", None)])],
                option_lines=[RichLine(parts=[("A．甲\tB．乙\tC．丙\tD．丁", None)])],
                source_number="755",
            )
        )

        materials = _repartition_strict_five_question_materials(
            [
                MaterialUnit(
                    header="材料三",
                    intro_lines=[RichLine(parts=[("", "prev.png")])],
                    questions=leading_questions + trailing_questions,
                ),
                MaterialUnit(
                    header="材料一",
                    intro_lines=[
                        RichLine(parts=[("", "chart1.png")]),
                        RichLine(parts=[("第二组材料正文", None)]),
                        RichLine(parts=[("A．160 亿元", None)]),
                        RichLine(parts=[("B．171 亿元", None)]),
                        RichLine(parts=[("C．181 亿元", None)]),
                        RichLine(parts=[("D．190 亿元", None)]),
                        RichLine(parts=[("A．多约73 亿元", None)]),
                        RichLine(parts=[("B．少约73 亿元", None)]),
                        RichLine(parts=[("C．多约86 亿元", None)]),
                        RichLine(parts=[("D．少约86 亿元", None)]),
                        RichLine(parts=[("A．4.8%", None)]),
                        RichLine(parts=[("B．5.5%", None)]),
                        RichLine(parts=[("C．6.5%", None)]),
                        RichLine(parts=[("D．3.9%", None)]),
                    ],
                    questions=[
                        ExamQuestion(
                            stem_lines=[RichLine(parts=[("754题题干", None)])],
                            option_lines=[
                                RichLine(parts=[("A．多约76%", None)]),
                                RichLine(parts=[("B．少约76%", None)]),
                                RichLine(parts=[("C．多约120%", None)]),
                                RichLine(parts=[("D．少约120%", None)]),
                                RichLine(parts=[("尾部材料说明", None)]),
                            ],
                            source_number="754",
                        )
                    ],
                ),
            ]
        )

        self.assertEqual(len(materials), 2)
        recovered = materials[1]
        self.assertEqual(
            [question.source_number for question in recovered.questions],
            ["751", "752", "753", "754", "755"],
        )
        intro_text = "".join(text for line in recovered.intro_lines for text, _ in line.parts)
        self.assertIn("第二组材料正文", intro_text)
        self.assertIn("尾部材料说明", intro_text)
        self.assertNotIn("160 亿元", intro_text)
        self.assertEqual(len(recovered.questions[0].option_lines), 4)
        self.assertEqual(len(recovered.questions[1].option_lines), 4)
        self.assertEqual(len(recovered.questions[2].option_lines), 4)
        self.assertEqual(len(recovered.questions[3].option_lines), 4)

    def test_repartition_strict_five_question_materials_moves_last_question_spill_to_next_group(self):
        questions = [
            ExamQuestion(
                stem_lines=[RichLine(parts=[(f"{index}题题干", None)])],
                option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                source_number=str(index),
            )
            for index in range(341, 351)
        ]
        questions[4].option_lines.append(RichLine(parts=[("", "material346.png")]))

        materials = _repartition_strict_five_question_materials(
            [
                MaterialUnit(
                    header="材料一",
                    intro_lines=[RichLine(parts=[("前一组材料", None)])],
                    questions=questions,
                )
            ]
        )

        self.assertEqual(len(materials), 2)
        self.assertEqual(
            [question.source_number for question in materials[0].questions],
            ["341", "342", "343", "344", "345"],
        )
        self.assertEqual(
            [question.source_number for question in materials[1].questions],
            ["346", "347", "348", "349", "350"],
        )
        self.assertEqual(
            [line.parts for line in materials[1].intro_lines],
            [[("", "material346.png")]],
        )

    def test_realign_adjacent_data_material_intros_moves_suffix_to_next_material(self):
        previous = MaterialUnit(
            header="材料（回答111-115题）",
            intro_lines=[
                RichLine(parts=[("“十三五”期间，我国平均每公顷森林蓄积量达到(", None)]),
                RichLine(parts=[("2020 年，受新冠肺炎疫情影响，我国民航全行业完成旅客运输量41777.82 万人次。", None)]),
                RichLine(parts=[("国内航线完成旅客运输量40821.30 万人次。", None)]),
                RichLine(parts=[("", "chart.png")]),
            ],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("“十三五”期间，我国平均每公顷森林蓄积量达到多少立方米？", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(111, 116)
            ],
        )
        current = MaterialUnit(
            header="材料（回答116-120题）",
            intro_lines=[],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("相比2019 年、2020 年我国民航全行业完成旅客运输中，国内航线完成旅客运输总量占比约：", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(116, 121)
            ],
        )

        materials = _realign_adjacent_data_material_intros([previous, current])

        self.assertEqual(len(materials[0].intro_lines), 1)
        self.assertEqual(len(materials[1].intro_lines), 3)
        moved_text = "".join(text for line in materials[1].intro_lines for text, _ in line.parts)
        self.assertIn("民航全行业完成旅客运输量", moved_text)

    def test_realign_adjacent_data_material_intros_can_move_whole_intro(self):
        previous = MaterialUnit(
            header="材料（回答161-165题）",
            intro_lines=[
                RichLine(parts=[("受疫情影响，2020 年全国社会消费品零售总额391981 亿元。", None)]),
                RichLine(parts=[("G 省消费市场受新冠肺炎疫情冲击更为明显。", None)]),
                RichLine(parts=[("", "sales.png")]),
            ],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2020 年上半年，我国农产品进口额中欧洲国家或地区约占：", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(161, 166)
            ],
        )
        current = MaterialUnit(
            header="材料（回答166-170题）",
            intro_lines=[],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2016—2020 年，G 省社会消费品零售总额年均增长率约为：", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(166, 171)
            ],
        )

        materials = _realign_adjacent_data_material_intros([previous, current])

        self.assertEqual(materials[0].intro_lines, [])
        self.assertEqual(len(materials[1].intro_lines), 3)
        moved_text = "".join(text for line in materials[1].intro_lines for text, _ in line.parts)
        self.assertIn("社会消费品零售总额", moved_text)

    def test_realign_adjacent_data_material_intros_keeps_generic_prompt_with_images_with_current_material(self):
        previous = MaterialUnit(
            header="材料（回答556-560题）",
            intro_lines=[
                RichLine(parts=[("根据以下资料，回答问题。", None)]),
                RichLine(parts=[("", "chart1.png")]),
                RichLine(parts=[("", "chart2.png")]),
            ],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2018 年我国平均每家海洋主题公园全年游客规模约为：", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(556, 561)
            ],
        )
        current = MaterialUnit(
            header="材料（回答561-565题）",
            intro_lines=[],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2018 年中国进出口贸易总额为4.62 万亿美元，其中集成电路进出口贸易额占比：", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(561, 566)
            ],
        )

        materials = _realign_adjacent_data_material_intros([previous, current])

        self.assertEqual(len(materials[0].intro_lines), 3)
        self.assertEqual(materials[1].intro_lines, [])

    def test_realign_adjacent_data_material_intros_does_not_move_image_only_intro_forward(self):
        previous = MaterialUnit(
            header="材料（回答331-335题）",
            intro_lines=[RichLine(parts=[("", "chart.png")])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("根据所给资料，下列说法正确的是：", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(331, 336)
            ],
        )
        current = MaterialUnit(
            header="材料（回答336-340题）",
            intro_lines=[],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("下一组资料分析题", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(336, 341)
            ],
        )

        materials = _realign_adjacent_data_material_intros([previous, current])

        self.assertEqual(len(materials[0].intro_lines), 1)
        self.assertEqual(materials[1].intro_lines, [])

    def test_realign_adjacent_data_material_intros_can_move_generic_prompt_text_back_to_previous_material(self):
        previous = MaterialUnit(
            header="材料（回答341-345题）",
            intro_lines=[],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2018 年1-7 月份，全国房地产开发投资约为多少亿元？", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(341, 346)
            ],
        )
        current = MaterialUnit(
            header="材料（回答346-350题）",
            intro_lines=[RichLine(parts=[("根据以下资料，回答问题。", None)])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2019 年全国实现 GDP 约多少万亿元？", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(346, 351)
            ],
        )

        materials = _realign_adjacent_data_material_intros([previous, current])

        self.assertEqual(len(materials[0].intro_lines), 1)
        self.assertEqual(materials[1].intro_lines, [])

    def test_realign_adjacent_data_material_intros_keeps_image_only_intro_with_current_material(self):
        previous = MaterialUnit(
            header="材料（回答341-345题）",
            intro_lines=[],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2018 年1-7 月份，全国房地产开发投资约为多少亿元？", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(341, 346)
            ],
        )
        current = MaterialUnit(
            header="材料（回答346-350题）",
            intro_lines=[RichLine(parts=[("", "chart.png")])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("2019 年全国实现 GDP 约多少万亿元？", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(index),
                )
                for index in range(346, 351)
            ],
        )

        materials = _realign_adjacent_data_material_intros([previous, current])

        self.assertEqual(materials[0].intro_lines, [])
        self.assertEqual(len(materials[1].intro_lines), 1)

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

    def test_question_number_supports_four_digit_sequence(self):
        items = [
            ("五. 判断推理：", None),
            ("999.", None),
            ("第999题题干", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("1000.", None),
            ("第1000题题干", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
            ("1001.", None),
            ("第1001题题干", None),
            ("A．甲\tB．乙\tC．丙\tD．丁", None),
        ]
        exam = parse_line_items(items, mode="reasoning")
        questions = exam.reasoning_sections[0].questions
        self.assertEqual([question.source_number for question in questions], ["999", "1000", "1001"])

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

    def test_parse_politics_question_recovers_cross_page_options_around_footer(self):
        items = [
            ("一. 政治理论：", None),
            ("1.(2025·联考)关于某项政策的说法,下列正确的有几项?", None),
            ("1第一项表述", None),
            ("2第二项表述", None),
            ("A.1 项", None),
            ("B.2 项", None),
            ("24", None),
            ("C.3 项", None),
            ("D.4 项", None),
            ("2.(2025·联考)下一题题干", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        exam = parse_line_items(items, mode="politics")
        questions = exam.politics_sections[0].questions

        self.assertEqual([question.source_number for question in questions], ["1", "2"])
        self.assertEqual(len(questions[0].option_lines), 4)
        self.assertIn("2第二项表述", "".join(text for line in questions[0].stem_lines for text, _ in line.parts))
        self.assertEqual(
            [line.parts[0][0] for line in questions[0].option_lines],
            ["A．1 项", "B．2 项", "C．3 项", "D．4 项"],
        )

    def test_parse_politics_question_accepts_inline_d_option_with_comma(self):
        items = [
            ("一. 政治理论：", None),
            ("3.(2025·联考)下列说法正确的是:A.甲 B.乙 C.丙 D,丁", None),
        ]

        exam = parse_line_items(items, mode="politics")
        question = exam.politics_sections[0].questions[0]

        self.assertEqual(question.source_number, "3")
        self.assertEqual(
            [line.parts[0][0] for line in question.option_lines],
            ["A．甲", "B．乙", "C．丙", "D．丁"],
        )

    def test_parse_politics_question_promotes_embedded_single_option_lines(self):
        items = [
            ("一. 政治理论：", None),
            ("3.(2025·联考)下列说法正确的是:", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D,丁", None),
        ]

        exam = parse_line_items(items, mode="politics")
        question = exam.politics_sections[0].questions[0]

        self.assertEqual(question.source_number, "3")
        self.assertEqual(
            [line.parts[0][0] for line in question.option_lines],
            ["A．甲", "B．乙", "C．丙", "D．丁"],
        )

    def test_parse_politics_question_repairs_page_footer_split_question(self):
        items = [
            ("一. 政治理论：", None),
            ("1.(2025·联考)下列说法正确的有几项?", None),
            ("1第一项表述", None),
            ("2第二项表述", None),
            ("3第三项表述", None),
            ("作", None),
            ("47", None),
            ("4第四项表述", None),
            ("A.12", None),
            ("B.34", None),
            ("C.14", None),
            ("D.23", None),
            ("2.(2025·联考)下一题题干", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        exam = parse_line_items(items, mode="politics")
        questions = exam.politics_sections[0].questions

        self.assertEqual([question.source_number for question in questions], ["1", "2"])
        self.assertIn("4第四项表述", "".join(text for line in questions[0].stem_lines for text, _ in line.parts))
        self.assertEqual(
            [line.parts[0][0] for line in questions[0].option_lines],
            ["A．12", "B．34", "C．14", "D．23"],
        )

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

    def test_preprocess_strips_answer_overview_block_before_next_heading(self):
        items = [
            ("16.(2018·吉林)某题题干", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("答案速览", None),
            ("1-5", None),
            ("DBACC", None),
            ("16", None),
            ("B", None),
            ("6", None),
            ("二、进阶提升", None),
            ("1.(2025·联考)下一题题干", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
        ]

        self.assertEqual(
            _preprocess_line_items(items),
            [
                ("16.", None),
                ("(2018·吉林)某题题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
                ("二、进阶提升", None),
                ("1.", None),
                ("(2025·联考)下一题题干", None),
                ("A.1", None),
                ("B.2", None),
                ("C.3", None),
                ("D.4", None),
            ],
        )

    def test_preprocess_strips_probable_mid_question_page_footers(self):
        items = [
            ("1第一项表述", None),
            ("2第二项表述", None),
            ("39", None),
            ("3第三项表述", None),
            ("A.1 项", None),
            ("B.2 项", None),
            ("96", None),
            ("C.3 项", None),
            ("D.4 项", None),
            ("38", None),
            ("26.(2021·广东)下一题题干", None),
        ]

        self.assertEqual(
            _preprocess_line_items(items),
            [
                ("1第一项表述", None),
                ("2第二项表述", None),
                ("3第三项表述", None),
                ("A.1 项", None),
                ("B.2 项", None),
                ("C.3 项", None),
                ("D.4 项", None),
                ("26.", None),
                ("(2021·广东)下一题题干", None),
            ],
        )

    def test_preprocess_strips_page_footer_before_reset_question(self):
        items = [
            ("三、高难突破", None),
            ("84", None),
            ("1.", None),
            ("(2024·山东)下一题题干", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        self.assertEqual(
            _preprocess_line_items(items),
            [
                ("三、高难突破", None),
                ("1.", None),
                ("(2024·山东)下一题题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ],
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

    def test_blank_option_markers_are_preserved_when_they_are_the_only_options(self):
        items = [
            ("四. 数量关系：", None),
            ("200. 某数量题题干", None),
            ("A.", None),
            ("B.", None),
            ("C.", None),
            ("D.", None),
            ("201. 下一题题干", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
        ]

        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions

        self.assertEqual([q.source_number for q in questions], ["200", "201"])
        self.assertEqual(
            [
                "".join(text for text, _img in line.parts)
                for line in questions[0].option_lines
            ],
            ["A．", "B．", "C．", "D．"],
        )

    def test_preprocess_ignores_page_footer_between_small_question_and_options(self):
        items = [
            ("三、高难突破", None),
            ("1.(2025·山东)关于发展理念，下列说法正确的是：", None),
            ("87", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        self.assertEqual(
            _preprocess_line_items(items),
            [
                ("三、高难突破", None),
                ("1.", None),
                ("(2025·山东)关于发展理念,下列说法正确的是:", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ],
        )

    def test_split_number_only_marker_keeps_following_numeric_sequence_question(self):
        items = [
            ("四. 数量关系：", None),
            ("199. 前一题", None),
            ("A.9000", None),
            ("B.8000", None),
            ("C.7000", None),
            ("D.6000", None),
            ("200.", None),
            ("两个大人带四个孩子去坐只有六个位置的圆型旋转木马,那么两个大人不相邻的概率为:", None),
            ("A.", None),
            ("B.", None),
            ("C.", None),
            ("D.", None),
            ("201.", None),
            ("3,35,99,195,(", None),
            (")", None),
            ("A.272", None),
            ("B.306", None),
            ("C.323", None),
            ("D.340", None),
            ("202.", None),
            ("23,24,22,25,21,26,(", None),
            ("),(", None),
            (")", None),
            ("A.19 28", None),
            ("B.20 27", None),
            ("C.26 21", None),
            ("D.32 39", None),
        ]

        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions

        self.assertEqual([q.source_number for q in questions], ["199", "200", "201", "202"])

    def test_split_number_only_marker_keeps_visual_question_after_option_block(self):
        items = [
            ("四. 数量关系：", None),
            ("121. 前一题", None),
            ("A.8", None),
            ("B.12", None),
            ("C.14", None),
            ("D.16", None),
            ("122.", None),
            ("如图,长为4,宽为2 的长方形ABCD 的顶点A 处有昆虫P、Q。", None),
            ("问在整个运动过程中,昆虫Q 的移动时间范围是:", None),
            ("", "figure.png"),
            ("123.", None),
            ("某社区积极为某受灾地区捐款捐物。", None),
            ("A.15%", None),
            ("B.20%", None),
            ("C.21%", None),
            ("D.25%", None),
        ]

        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions

        self.assertEqual([q.source_number for q in questions], ["121", "122", "123"])
        self.assertTrue(any(part[1] == "figure.png" for part in questions[1].stem_lines[-1].parts))

    def test_number_only_question_with_leading_image_payload_is_kept(self):
        items = [
            ("四. 数量关系：", None),
            ("199. 前一题", None),
            ("A.11", None),
            ("B.12", None),
            ("C.13", None),
            ("D.14", None),
            ("", "figure.png"),
            ("200.", None),
            ("201. 下一题题干", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
        ]

        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions

        self.assertEqual([q.source_number for q in questions], ["199", "200", "201"])
        self.assertEqual(questions[1].stem_lines[0].parts[0][1], "figure.png")
        self.assertEqual(questions[1].option_lines, [])

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

    def test_prefixed_entity_labels_in_quant_stem_do_not_start_option_cluster(self):
        items = [
            ("四. 数量关系：", None),
            ("7.", None),
            ("A、B 两地直线距离为320 千米,甲车从A 地、乙车从B 地同时出发相向而行。已知甲车的速度是乙", None),
            ("车的3 倍,乙车与甲车相遇后,立即右转弯90 度并保持原速度继续行驶。那么甲车到达B 地时,其与乙车", None),
            ("之间的距离:", None),
            ("A.小于80 千米", None),
            ("B.80~90 千米", None),
            ("C.90~100 千米", None),
            ("D.大于100 千米", None),
        ]
        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertEqual(question.source_number, "7")
        self.assertIn("A、B 两地直线距离为320 千米", stem_text)
        self.assertIn("乙车与甲车相遇后", stem_text)
        self.assertEqual(len(question.option_lines), 4)

    def test_prefixed_entity_list_at_line_start_is_kept_in_quant_stem(self):
        items = [
            ("四. 数量关系：", None),
            ("216.", None),
            ("A、B、C 三个社区需要建设若干个5G 基站,三个社区可供选择的建设基站地点分别有2 个、4 个、", None),
            ("5 个,现从A、B、C 三个社区分别选取1、2、3 个地点随机分配给甲、乙、丙三个施工队进行建设,要求每", None),
            ("A.720 种", None),
            ("B.480 种", None),
            ("C.360 种", None),
            ("D.120 种", None),
        ]
        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertEqual(question.source_number, "216")
        self.assertIn("A、B、C 三个社区需要建设若干个5G 基站", stem_text)
        self.assertIn("随机分配给甲、乙、丙", stem_text)
        self.assertEqual(len(question.option_lines), 4)

    def test_prefixed_entity_list_mid_stem_is_not_parsed_as_options(self):
        items = [
            ("四. 数量关系：", None),
            ("208.", None),
            ("某街道服务中心的80 名职工通过相互投票选出6 名年度优秀职工,每人都只投一票,最终A、B、", None),
            ("C、D、E、F 这6 人当选。已知A 票数最多,共获得20 张选票;B、C 两人的票数相同,并列第2;", None),
            ("D、E 两人票数也相同,并列第3;F 获得10 张选票,排在第4。那么B、C 获得的选票最多为(", None),
            (")张。", None),
            ("A.11", None),
            ("B.12", None),
            ("C.13", None),
            ("D.14", None),
        ]
        exam = parse_line_items(items, mode="quant")
        question = exam.quant_sections[0].questions[0]
        stem_text = "".join(text for line in question.stem_lines for text, _ in line.parts)
        self.assertIn("C、D、E、F 这6 人当选", stem_text)
        self.assertIn("B、C 获得的选票最多为", stem_text)
        self.assertEqual(len(question.option_lines), 4)

    def test_parse_quant_block_repairs_shifted_image_only_reasoning_sequence(self):
        items = [
            ("五. 判断推理：", None),
            ("1330.", None),
            ("前题题干", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁", None),
            ("", "q1331.png"),
            ("1331.", None),
            ("", "q1332.png"),
            ("1332.", None),
            ("1333.", None),
            ("", "q1334.png"),
            ("1334.", None),
            ("1335.", None),
            ("1336.（1）金星（2）火星", None),
            ("A．3-2-5-4-1", None),
            ("B．4-2-1-5-3", None),
            ("C．4-1-2-3-5", None),
            ("D．2-1-5-3-4", None),
        ]
        exam = parse_line_items(items, mode="reasoning")
        questions = exam.reasoning_sections[0].questions
        self.assertEqual(
            [q.source_number for q in questions],
            ["1330", "1331", "1332", "1333", "1334", "1335", "1336"],
        )
        self.assertEqual(questions[0].option_lines[-1].parts[0][1], None)
        self.assertEqual(questions[1].stem_lines[0].parts[0][1], "q1331.png")
        self.assertEqual(questions[2].stem_lines[0].parts[0][1], "q1332.png")
        self.assertEqual(questions[3].stem_lines, [])
        self.assertEqual(questions[4].stem_lines[0].parts[0][1], "q1334.png")
        self.assertEqual(questions[5].stem_lines, [])

    def test_parse_quant_block_restores_prompt_only_visual_question_before_text_question(self):
        items = [
            ("五. 判断推理：", None),
            ("1382.从所给的四个选项中,选择最合适的一个填入问号处,使之呈现一定的规律性:", None),
            ("", "q1382.png"),
            ("1383.从所给的四个选项中,选择最合适的一个填入问号处,使之呈现一定的规律性:", None),
            ("1384.下图为给定的多面体及其外表面展开图,问字母A、B、C、D 和数字1、2、3、4 代表的棱的对应", None),
            ("关系为:", None),
            ("", "q1384a.png"),
            ("", "q1384b.png"),
            ("A.1-C,2-A,3-B,4-D", None),
            ("B.1-A,2-C,3-B,4-D", None),
            ("C.1-A,2-C,3-D,4-B", None),
            ("D.1-C,2-A,3-D,4-B", None),
        ]
        exam = parse_line_items(items, mode="reasoning")
        questions = exam.reasoning_sections[0].questions
        self.assertEqual([q.source_number for q in questions], ["1382", "1383", "1384"])
        q1383_text = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertIn("从所给的四个选项中", q1383_text)
        self.assertEqual(questions[1].option_lines, [])
        q1384_text = "".join(text for line in questions[2].stem_lines for text, _img in line.parts)
        self.assertTrue(q1384_text.startswith("下图为给定的多面体"))
        self.assertNotIn("从所给的四个选项中", q1384_text)

    def test_parse_quant_block_recovers_trailing_image_questions_before_next_text_question(self):
        items = [
            ("四. 数量关系：", None),
            ("352.4,5,(),14,22,27", None),
            ("A.8", None),
            ("B.9", None),
            ("C.10", None),
            ("D.11", None),
            ("", "q353.png"),
            ("353.", None),
            ("", "q354.png"),
            ("354.", None),
            ("355.3,10,21,36,()", None),
            ("A.55", None),
            ("B.56", None),
            ("C.58", None),
            ("D.62", None),
        ]
        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions
        self.assertEqual([q.source_number for q in questions], ["352", "353", "354", "355"])
        self.assertEqual(len(questions[0].option_lines), 4)
        self.assertEqual(questions[1].stem_lines[0].parts[0][1], "q353.png")
        self.assertEqual(questions[1].option_lines, [])
        self.assertEqual(questions[2].stem_lines[0].parts[0][1], "q354.png")
        self.assertEqual(questions[2].option_lines, [])
        q355_text = "".join(text for line in questions[3].stem_lines for text, _img in line.parts)
        self.assertEqual(q355_text, "3,10,21,36,()")
        self.assertEqual(len(questions[3].option_lines), 4)

    def test_parse_quant_block_restores_dropped_tens_digit_sequence_with_multiline_stem_image_question(self):
        items = [
            ("四. 数量关系：", None),
            ("12.", None),
            ("上一题题干", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
            ("3.", None),
            ("某班级对70 多名学生进行数学和英语科目摸底测验。", None),
            ("那两科均及格的学生有多少人?", None),
            ("A.31", None),
            ("B.37", None),
            ("C.41", None),
            ("D.44", None),
            ("4.", None),
            ("收割一片稻田,可选择甲、乙、丙3 台农机。", None),
            ("那么丙的收割速度在以下哪个范围内?", None),
            ("A.小于6 亩/小时", None),
            ("B.6~7 亩/小时", None),
            ("C.7~8 亩/小时", None),
            ("D.大于8 亩/小时", None),
            ("5.", None),
            ("公司研发部门共5 名员工,年龄各不相同。", None),
            ("年龄排名第三的员工最大可能为多少岁?", None),
            ("A.33", None),
            ("B.34", None),
            ("C.35", None),
            ("D.36", None),
            ("6.", None),
            ("一只闹钟的秒针顶点距离表盘圆心4 厘米。", None),
            ("小王烧开一壶水的时间内,秒针顶点累计移动了", None),
            ("厘米。那么这一时间段内,分针顶点与表盘圆心的连线扫过的扇形面积为多", None),
            ("少平方厘米?", None),
            ("", "clock.png"),
            ("7.", None),
            ("A、B 两地直线距离为320 千米。", None),
            ("那么甲车到达B 地时,其与乙车之间的距离:", None),
            ("A.小于80 千米", None),
            ("B.80~90 千米", None),
            ("C.90~100 千米", None),
            ("D.大于100 千米", None),
            ("8.", None),
            ("甲以技术入股加入某互联网初创企业。", None),
            ("那么第一次投资前公司的估算价值是第二次投资的多少倍?", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
            ("9.", None),
            ("有30 个2 克的砝码和8 个5 克的砝码。", None),
            ("有多少个不能用这些砝码称量出来?", None),
            ("A.0", None),
            ("B.1", None),
            ("C.2", None),
            ("D.3", None),
            ("20.", None),
            ("12 个人排成1 列纵队。", None),
            ("那么有多少种重新编队的方法?", None),
            ("A.16", None),
            ("B.18", None),
            ("C.20", None),
            ("D.24", None),
        ]
        exam = parse_line_items(items, mode="quant")
        questions = exam.quant_sections[0].questions
        self.assertEqual(
            [q.source_number for q in questions],
            ["12", "13", "14", "15", "16", "17", "18", "19", "20"],
        )
        q16_text = "".join(text for line in questions[4].stem_lines for text, _img in line.parts)
        self.assertIn("一只闹钟的秒针顶点距离表盘圆心4 厘米", q16_text)
        self.assertTrue(any(part[1] == "clock.png" for line in questions[4].stem_lines for part in line.parts))

    def test_parse_quant_block_moves_next_stem_out_of_previous_option_tail(self):
        items = [
            ("三. 言语理解：", None),
            ("865.上一题题干", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁866.人脸识别系统深度学习的数据越多", None),
            ("机器就能够自主学习到伪造痕迹。", None),
            ("依次填入划横线部分最恰当的一项是:", None),
            ("A.弊端无计可施", None),
            ("B.瑕疵无所遁形", None),
            ("C.错误插翅难逃", None),
            ("D.缺陷束手无策", None),
        ]
        exam = parse_line_items(items, mode="verbal")
        questions = exam.verbal_sections[0].questions
        self.assertEqual([q.source_number for q in questions], ["865", "866"])
        self.assertEqual(
            ["".join(text for text, _img in line.parts) for line in questions[0].option_lines],
            ["A．甲", "B．乙", "C．丙", "D．丁"],
        )
        q866_text = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertIn("人脸识别系统深度学习的数据越多", q866_text)
        self.assertIn("依次填入划横线部分最恰当的一项是", q866_text)
        self.assertEqual(len(questions[1].option_lines), 4)

    def test_parse_quant_block_moves_quote_prefixed_stem_out_of_previous_option_tail(self):
        items = [
            ("三. 言语理解：", None),
            ("1632.将以上6个句子重新排列,语序正确的一项是:", None),
            ("A.163245", None),
            ("B.134526", None),
            ("C.416325", None),
            ("D.425316", None),
            ("1633.“", None),
            ("。”人民代表大会制度之所以具有强大生命力和显著优越性,关键在于深深植根于", None),
            ("人民之中。填入画横线部分最恰当的一项是:", None),
            ("A.人视水见形,视民知治不", None),
            ("B.为政之要,以顺民心为本", None),
            ("C.善为政者,弊则补之,决则塞之", None),
            ("D.苟利于民,不必法古;苟周于事,不必循俗", None),
        ]
        exam = parse_line_items(items, mode="verbal")
        questions = exam.verbal_sections[0].questions
        self.assertEqual([q.source_number for q in questions], ["1632", "1633"])
        q1633_text = "".join(text for line in questions[1].stem_lines for text, _img in line.parts)
        self.assertIn("人民代表大会制度之所以具有强大生命力", q1633_text)
        self.assertIn("填入画横线部分最恰当的一项是", q1633_text)
        self.assertEqual(len(questions[1].option_lines), 4)

    def test_parse_quant_block_ignores_numeric_continuation_after_option_line(self):
        items = [
            ("三. 言语理解：", None),
            ("1344.", None),
            ("下列选项中,有语病的一项是:", None),
            ("A.“枪声就是命令。”驻扎在附近的八路军听到枪声后,立即赶来营救,有", None),
            ("24.", None),
            ("名官兵牺牲在这里。", None),
            ("B.第二项", None),
            ("C.第三项", None),
            ("D.第四项", None),
        ]
        exam = parse_line_items(items, mode="verbal")
        question = exam.verbal_sections[0].questions[0]
        stem_text = "".join(text for line in question.stem_lines for text, _img in line.parts)
        self.assertEqual(question.source_number, "1344")
        self.assertIn("有语病的一项是", stem_text)
        self.assertIn("24.", stem_text)
        self.assertIn("名官兵牺牲在这里。", stem_text)
        self.assertEqual(len(exam.verbal_sections[0].questions), 1)


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

    def test_preprocess_keeps_numeric_sequence_after_question_number_line(self):
        items = [
            ("77. 92.46,84.42,76.38,68.34,(", None),
            (")", None),
            ("A.50.25", None),
            ("B.53.26", None),
            ("C.55.17", None),
            ("D.56.30", None),
        ]

        processed = _preprocess_line_items(items)
        texts = [text for text, _ in processed if text]

        self.assertIn("77.", texts)
        self.assertIn("92.46,84.42,76.38,68.34,(", texts)
        self.assertNotIn("92.", texts)

    def test_preprocess_splits_spaced_numeric_sequence_question_line(self):
        items = [
            ("352.4 ,5,(),14,22,27", None),
            ("A.8", None),
            ("B.9", None),
            ("C.10", None),
            ("D.11", None),
        ]

        processed = _preprocess_line_items(items)

        self.assertEqual(processed[0], ("352.", None))
        self.assertEqual(processed[1], ("4 ,5,(),14,22,27", None))

    def test_parse_quant_block_merges_split_question_with_wrong_following_marker(self):
        items = _preprocess_line_items(
            [
                ("36.上一题题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
                ("37.研究人员开发了一种可以从指尖上的汗水中获取能量的新设备。只需按一下手指,就能额外产生", None),
                ("30.毫焦耳的能量。这意味着可自我维持的可穿戴电子产品更实用。以下哪项如果为真,最能支持上述论述?", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
                ("38.下一题题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ]
        )

        questions = parse_quant_block(items, 0, len(items))
        stem = "\n".join(
            "".join(text for text, _img in line.parts)
            for line in questions[1].stem_lines
        )

        self.assertEqual([question.source_number for question in questions], ["36", "37", "38"])
        self.assertIn("毫焦耳的能量", stem)

    def test_parse_quant_block_merges_wrong_small_marker_into_next_question(self):
        items = _preprocess_line_items(
            [
                ("610.前一题题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
                ("3", None),
                ("在目前留存的海洋石刻遗产中,比较集中的是第三大类即海洋宗教文化石刻。", None),
                ("611.下面这段文字,最适合填入文中的哪个位置?", None),
                ("A.1", None),
                ("B.2", None),
                ("C.3", None),
                ("D.4", None),
                ("612.下一题题干", None),
                ("A.甲", None),
                ("B.乙", None),
                ("C.丙", None),
                ("D.丁", None),
            ]
        )

        questions = parse_quant_block(items, 0, len(items))
        stem = "\n".join(
            "".join(text for text, _img in line.parts)
            for line in questions[1].stem_lines
        )

        self.assertEqual([question.source_number for question in questions], ["610", "611", "612"])
        self.assertIn("海洋石刻", stem)

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

    def test_filename_set_paper_supports_shorter_set_sequence_without_fixed_total(self):
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
            ("31.某商品按8折出售后利润率为20%,其成本是多少?", None),
            ("A.80", None),
            ("B.96", None),
            ("C.100", None),
            ("D.120", None),
            ("36.如果甲成立,那么乙成立。以下哪项最能削弱上述论证?", None),
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
        self.assertEqual([q.source_number for q in exam.quant_sections[0].questions], ["31"])
        self.assertEqual([q.source_number for q in exam.reasoning_sections[0].questions], ["36"])


    def test_cross_page_material_merges_truncated_intro(self):
        """材料正文在跨页处被截断（句末为逗号），自动拆组时应合并。"""
        from core.pdf_exam_parse import _merge_cross_page_material_units

        unit_a = MaterialUnit(
            header="材料一",
            intro_lines=[RichLine(parts=[("2024年全国经济运行总体平稳，", None)])],
            questions=[
                ExamQuestion(stem_lines=[RichLine(parts=[("题干A", None)])], option_lines=[
                    RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])
                ], source_number="111"),
            ],
        )
        unit_b = MaterialUnit(
            header="材料二",
            intro_lines=[RichLine(parts=[("GDP同比增长5.2%。", None)])],
            questions=[
                ExamQuestion(stem_lines=[RichLine(parts=[("题干B", None)])], option_lines=[
                    RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])
                ], source_number="112"),
            ],
        )

        merged = _merge_cross_page_material_units([unit_a, unit_b])
        self.assertEqual(len(merged), 1)
        self.assertEqual(len(merged[0].intro_lines), 2)
        self.assertEqual(len(merged[0].questions), 2)

    def test_cross_page_material_preserves_independent_materials(self):
        """材料正文以句号结尾时不应合并。"""
        from core.pdf_exam_parse import _merge_cross_page_material_units

        unit_a = MaterialUnit(
            header="材料一",
            intro_lines=[RichLine(parts=[("材料A正文。", None)])],
            questions=[
                ExamQuestion(stem_lines=[RichLine(parts=[("题干A", None)])], option_lines=[
                    RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])
                ], source_number="111"),
            ],
        )
        unit_b = MaterialUnit(
            header="材料二",
            intro_lines=[RichLine(parts=[("材料B正文。", None)])],
            questions=[
                ExamQuestion(stem_lines=[RichLine(parts=[("题干B", None)])], option_lines=[
                    RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])
                ], source_number="116"),
            ],
        )

        merged = _merge_cross_page_material_units([unit_a, unit_b])
        self.assertEqual(len(merged), 2)

    def test_cross_page_table_continuation_merges(self):
        """跨页表格（含表格续行字符）应合并。"""
        from core.pdf_exam_parse import _merge_cross_page_material_units

        unit_a = MaterialUnit(
            header="材料一",
            intro_lines=[
                RichLine(parts=[("年份│GDP│增速", None)]),
                RichLine(parts=[("2023│126│5.2%", None)]),
            ],
            questions=[
                ExamQuestion(stem_lines=[RichLine(parts=[("题干", None)])], option_lines=[
                    RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])
                ], source_number="111"),
            ],
        )
        unit_b = MaterialUnit(
            header="材料二",
            intro_lines=[
                RichLine(parts=[("2024│132│4.8%", None)]),
            ],
            questions=[
                ExamQuestion(stem_lines=[RichLine(parts=[("题干B", None)])], option_lines=[
                    RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])
                ], source_number="112"),
            ],
        )

        merged = _merge_cross_page_material_units([unit_a, unit_b])
        self.assertEqual(len(merged), 1)
        self.assertEqual(len(merged[0].intro_lines), 3)

    def test_cross_page_material_does_not_merge_distinct_question_ranges(self):
        """明确标了不同答题区间的材料组不应被跨页续接规则误并。"""
        from core.pdf_exam_parse import _merge_cross_page_material_units

        unit_a = MaterialUnit(
            header="材料一:根据材料,回答101—105题。",
            intro_lines=[RichLine(parts=[("2024年上半年工业企业利润同比下降，", None)])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("101.题干A", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number="101",
                )
            ],
        )
        unit_b = MaterialUnit(
            header="材料二:根据材料,回答106—110题。",
            intro_lines=[RichLine(parts=[("2022年12月国内市场手机出货量2786万部。", None)])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[("106.题干B", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number="106",
                )
            ],
        )

        merged = _merge_cross_page_material_units([unit_a, unit_b])
        self.assertEqual(len(merged), 2)
        self.assertEqual(merged[0].header, unit_a.header)
        self.assertEqual(merged[1].header, unit_b.header)

    def test_cross_page_material_does_not_merge_distinct_material_ordinals(self):
        from core.pdf_exam_parse import _merge_cross_page_material_units

        unit_a = MaterialUnit(
            header="材料一",
            intro_lines=[RichLine(parts=[("2024年上半年工业企业利润同比下降，", None)])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[(f"{101 + offset}.题干A", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(101 + offset),
                )
                for offset in range(5)
            ],
        )
        unit_b = MaterialUnit(
            header="材料二",
            intro_lines=[RichLine(parts=[("2022年12月国内市场手机出货量2786万部。", None)])],
            questions=[
                ExamQuestion(
                    stem_lines=[RichLine(parts=[(f"{106 + offset}.题干B", None)])],
                    option_lines=[RichLine(parts=[("A．1\tB．2\tC．3\tD．4", None)])],
                    source_number=str(106 + offset),
                )
                for offset in range(5)
            ],
        )

        merged = _merge_cross_page_material_units([unit_a, unit_b])
        self.assertEqual(len(merged), 2)

    def test_parse_material_body_corrects_mismatched_header_question_range(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("111.2023年3月末，5G移动电话用户数约比上年末增长：", None),
            ("A.10%", None),
            ("B.18%", None),
            ("C.25%", None),
            ("D.30%", None),
            ("112.根据上述资料，下列说法正确的是：", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("113.以下哪项最符合题意？", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
            ("114.表2中“？”处的数值为：", None),
            ("A.1", None),
            ("B.2", None),
            ("C.3", None),
            ("D.4", None),
            ("115.根据上述资料，下列能够推出的是：", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        unit = parse_material_body(
            items,
            0,
            len(items),
            "材料三：根据材料，回答110—115 题。",
        )

        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual(unit.header, "材料三：根据材料，回答111—115 题。")
        self.assertEqual([q.source_number for q in unit.questions], ["111", "112", "113", "114", "115"])

    def test_parse_material_body_preserves_image_only_question_before_next_marker(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("2021 年某市工业增加值同比增长8.3%。", None),
            ("24.", None),
            ("以下折线图中，最能反映增速变化趋势的是：", None),
            ("", "chart.png"),
            ("25.", None),
            ("根据上述材料，下列说法正确的是：", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料二")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual(len(unit.intro_lines), 1)
        self.assertEqual([question.source_number for question in unit.questions], ["24", "25"])
        self.assertTrue(any(part[1] == "chart.png" for part in unit.questions[0].stem_lines[1].parts))
        self.assertEqual(len(unit.questions[0].option_lines), 0)
        self.assertEqual(len(unit.questions[1].option_lines), 4)

    def test_parse_material_body_keeps_numbered_image_only_question_before_next_marker(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("97.", None),
            ("2020 年我国实缴职工的人均实缴住房公积金为:", None),
            ("A.1.50 万元", None),
            ("B.1.61 万元", None),
            ("C.1.71 万元", None),
            ("D.1.87 万元", None),
            ("98.", None),
            ("", "chart.png"),
            ("99.", None),
            ("2016-2020 年我国住房公积金实缴职工人数年增长超过4%的年份个数是:", None),
            ("A.2", None),
            ("B.3", None),
            ("C.4", None),
            ("D.5", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料三")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["97", "98", "99"])
        self.assertEqual(unit.questions[1].option_lines, [])
        self.assertTrue(any(part[1] == "chart.png" for part in unit.questions[1].stem_lines[0].parts))

    def test_parse_material_body_keeps_split_number_only_question_with_plain_stem(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("287.", None),
            ("按居民人均可支配收入数据估算,2019 年上半年农村与城镇居民的人数比例约为()。", None),
            ("A.1:2", None),
            ("B.2:3", None),
            ("C.3:4", None),
            ("D.4:5", None),
            ("288.", None),
            ("2017 年上半年,全国居民人均可支配收入中位数为元。", None),
            ("A.9234", None),
            ("B.11242", None),
            ("C.12186", None),
            ("D.13281", None),
            ("289.", None),
            ("2018 年上半年,全国居民人均财产净收入所占支配收入的比重为()。", None),
            ("A.5.7%", None),
            ("B.8.3%", None),
            ("C.8.6%", None),
            ("D.12.3%", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料一")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["287", "288", "289"])

    def test_parse_material_body_keeps_split_number_only_question_with_inline_options(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("492.", None),
            ("2017 年下半年,我国平均每月进口原油:", None),
            ("A.不到3300 万吨", None),
            ("B.在3300~3400 万吨之间", None),
            ("C.在3400~3500 万吨之间", None),
            ("D.超过3500 万吨", None),
            ("493.", None),
            ("2017 年9 月和10 月,我国原油进口金额分别为136 亿美元和121.4 亿美元。", None),
            ("10 月平均每吨原油的进口价格较9 月:A.上升了不到50 美元B.上升了50 美元以上C.下降了不到50 美元D.下降了50 美元以上", None),
            ("494.", None),
            ("2017 年4 月~2018 年4 月间,原油进口量同比增速超过煤炭进口量同比增速的月份有多少个?", None),
            ("A.3", None),
            ("B.5", None),
            ("C.8", None),
            ("D.10", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料一")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["492", "493", "494"])
        self.assertEqual(len(unit.questions[1].option_lines), 4)

    def test_parse_material_body_accepts_glued_option_lines_without_separator(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("327.", None),
            ("2013-2017 年，我国成年人人均期刊阅读量超过这五年平均水平的年份有()。", None),
            ("A2 个", None),
            ("B3 个", None),
            ("C4 个", None),
            ("D5 个", None),
            ("328.", None),
            ("2016 年我国成年人数字化阅读四个方式的接触率从高到低排列正确的是()。", None),
            ("A 网络在线阅读> 手机阅读> 电子阅读器阅读> 平板电脑阅读", None),
            ("B 手机阅读> 网络在线阅读> 电子阅读器阅读> 平板电脑阅读", None),
            ("C 网络在线阅读> 手机阅读> 平板电脑阅读> 电子阅读器阅读", None),
            ("D 手机阅读> 网络在线阅读> 平板电脑阅读> 电子阅读器阅读", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料一")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["327", "328"])
        self.assertEqual(len(unit.questions[0].option_lines), 4)

    def test_parse_material_body_keeps_split_blank_prompt_without_options(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("749.", None),
            ("2019 年全国艺术表演团体机构比2017 年多(", None),
            (")。", None),
            ("750.", None),
            ("根据上表，下列说法正确的是(", None),
            (")。", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料三")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["749", "750"])

    def test_parse_material_body_redistributes_orphan_option_clusters_to_previous_questions(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("749.", None),
            ("2019 年全国艺术表演团体机构比2017 年多(", None),
            (")。", None),
            ("750.", None),
            ("根据上表,下列说法正确的是(", None),
            (")。", None),
            ("C.“十二五”期间,全国艺术表演团体机构数实现翻一番", None),
            ("755.", None),
            ("从所给资料可知,下列说法中不正确的是(", None),
            (")。", None),
            ("A.基础研究经费同比增长最快", None),
            ("B.2020 年其他经费约1.53 亿元", None),
            ("C.2020 年各类企业研究与试验发展(", None),
            ("D.2020 年按研究与试验发展(", None),
            ("A.672 个", None),
            ("B.10740 个", None),
            ("C.13%", None),
            ("D.52%", None),
            ("A.2019 年,全国艺术表演团体的机构数、从业人员数、演出场次均同比增加", None),
            ("B.与2011 年相比,2019 年全国艺术表演团体从业人员数增速低于机构数", None),
            ("D.2017-2019 年,全国艺术表演团体演出观众人次年均下降3%", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料三")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["749", "750", "755"])
        self.assertEqual(len(unit.questions[0].option_lines), 4)
        self.assertEqual(len(unit.questions[1].option_lines), 4)
        self.assertEqual(len(unit.questions[2].option_lines), 4)
        second_stem = "".join(text for line in unit.questions[1].stem_lines for text, _ in line.parts)
        self.assertNotIn("十二五", second_stem)

    def test_parse_material_body_accepts_digit_leading_stem_after_chinese_comma_number(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("453、11 条特高压线路中，输送可再生能源电量占2016 年输送电量一半以上的线路有多少条？", None),
            ("A．7", None),
            ("B．6", None),
            ("C．5", None),
            ("D．4", None),
            ("454、11 条特高压线路中，2016 年可再生能源占输送电量比重为 0 的线路有多少条？", None),
            ("A．1", None),
            ("B．2", None),
            ("C．3", None),
            ("D．4", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料一")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["453", "454"])
        self.assertIn("11 条特高压线路中", "".join(part for part, _ in unit.questions[0].stem_lines[0].parts))
        self.assertIn("11 条特高压线路中", "".join(part for part, _ in unit.questions[1].stem_lines[0].parts))

    def test_parse_material_body_accepts_digit_leading_stem_without_gap_after_dot_number(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("142.131 家证券公司中，平均每家证券公司在2018 年第一季度实现营业收入约为：", None),
            ("A．659.4 亿元", None),
            ("B．5.0 亿元", None),
            ("C．669.5 亿元", None),
            ("D．6.0 亿元", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料二")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["142"])
        self.assertIn("131 家证券公司中", "".join(part for part, _ in unit.questions[0].stem_lines[0].parts))

    def test_parse_material_body_keeps_year_leading_first_question_after_intro(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("2017 年，我国体育产业总规模(总产出)达到21987.7 亿元，同比增长15.66%。", None),
            ("316、2016 年我国体育产业总规模达到多少亿元()。", None),
            ("A、14176.3", None),
            ("B、18295.6", None),
            ("C、19010.6", None),
            ("D、21036.4", None),
            ("317、2017 年体育服务业增加值占比应为多少?", None),
            ("A、55%", None),
            ("B、56.9%", None),
            ("C、58.2%", None),
            ("D、59.2%", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料二")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["316", "317"])
        intro_text = "".join(text for line in unit.intro_lines for text, _ in line.parts)
        self.assertIn("2017 年，我国体育产业总规模", intro_text)
        first_stem = "".join(text for line in unit.questions[0].stem_lines for text, _ in line.parts)
        self.assertIn("2016 年我国体育产业总规模达到多少亿元", first_stem)

    def test_parse_material_body_keeps_multiple_year_leading_questions(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("2017 年，S 市服务业小微样本企业总体实现营业收入105.39 亿元，同比增长3.1%。", None),
            ("601.2017 年，S 市服务业小微样本企业共有多少户?", None),
            ("A.不到3000 户", None),
            ("B.3000~4000 户之间", None),
            ("C.4001~5000 户之间", None),
            ("D.超过5000 户", None),
            ("602.2017 年，S 市服务业小微样本企业平均每万元资产实现营业收入比2015 年:", None),
            ("A.增长了不到5%", None),
            ("B.增长了5%以上", None),
            ("C.下降了不到5%", None),
            ("D.下降了5%以上", None),
            ("603.如S 市服务业小微样本企业数量为固定值，问2017 年S 市服务业小微样本企业户均比上年少缴纳营业税金及附加多少万元?", None),
            ("A.1.1", None),
            ("B.2.2", None),
            ("C.3.3", None),
            ("D.4.4", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料三")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["601", "602", "603"])
        intro_text = "".join(text for line in unit.intro_lines for text, _ in line.parts)
        self.assertIn("2017 年，S 市服务业小微样本企业总体实现营业收入", intro_text)

    def test_parse_material_body_keeps_digit_leading_stem_after_split_question_number_line(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("141.", None),
            ("2018 年第一季度，131 家证券公司资产管理业务净收入约为多少亿元?", None),
            ("A．120", None),
            ("B．210", None),
            ("C．275", None),
            ("D．315", None),
            ("142.", None),
            ("131 家证券公司中，平均每家证券公司在2018 年第一季度实现营业收入约为：", None),
            ("A．659.4 亿元", None),
            ("B．5.0 亿元", None),
            ("C．669.5 亿元", None),
            ("D．6.0 亿元", None),
            ("143.", None),
            ("2019 年第一季度，131 家证券公司总资产的同比增速约为：", None),
            ("A．2%", None),
            ("B．5%", None),
            ("C．8%", None),
            ("D．12%", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料二")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["141", "142", "143"])
        self.assertIn("131 家证券公司中", "".join(part for part, _ in unit.questions[1].stem_lines[0].parts))

    def test_parse_material_body_ignores_numeric_material_fragment_without_options(self):
        from core.pdf_exam_parse import _preprocess_line_items, parse_material_body

        items = [
            ("2020 年我国数字经济规模达到39.2 万亿元，保持9.7%的高位增长，占GDP 比重为38.6%，同比提升", None),
            ("2.4 个百分点。", None),
            ("产业数字化发展深入推进，2020 年我国服务业、工业、农业数字经济占行业增加值比重分别为40.7%、", None),
            ("21.0%和8.9%。", None),
            ("66.2020 年，数字经济规模超过5000 亿元的省份有多少个?", None),
            ("A．8", None),
            ("B．11", None),
            ("C．16", None),
            ("D．21", None),
        ]

        processed = _preprocess_line_items(items)
        unit = parse_material_body(processed, 0, len(processed), "材料四")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["66"])

    def test_parse_material_body_repairs_local_misread_question_number(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("843.", None),
            ("2014 年一季度该省主要农作物总收获面积约为多少万亩?", None),
            ("A．1250", None),
            ("B．1225", None),
            ("C．1150", None),
            ("D．1125", None),
            ("854.", None),
            ("下列选项中，2015 年一季度比上年同期收获面积增量最多的是哪类作物?", None),
            ("A．旱粮", None),
            ("B．薯类", None),
            ("C．油料作物", None),
            ("D．药材", None),
            ("845.", None),
            ("关于2015 年一季度 G 省主要农作物春收情况，下列说法错误的是：", None),
            ("A．甲", None),
            ("B．乙", None),
            ("C．丙", None),
            ("D．丁", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料三")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["843", "844", "845"])

    def test_preprocess_splits_glued_year_question_transition(self):
        from core.pdf_exam_parse import _preprocess_line_items

        items = [("4787.2014 年,100 美元的年平均价约为(", None)]
        processed = _preprocess_line_items(items)
        self.assertEqual(processed, [("477.", None), ("2014 年,100 美元的年平均价约为(", None)])

    def test_parse_material_body_ignores_note_enumeration_numbering(self):
        from core.pdf_exam_parse import parse_material_body

        items = [
            ("注:", None),
            ("1.", None),
            ("在线旅游市场包括在线机票、在线住宿和其他", None),
            ("2.", None),
            ("渗透率", None),
            ("296.", None),
            ("与2011 年相比，2016 年我国在线旅游市场交易规模增长了多少亿元?", None),
            ("A．3152.2", None),
            ("B．3180.4", None),
            ("C．3196.7", None),
            ("D．3220.5", None),
        ]

        unit = parse_material_body(items, 0, len(items), "材料二")
        self.assertIsNotNone(unit)
        assert unit is not None
        self.assertEqual([question.source_number for question in unit.questions], ["296"])

    def test_parse_quant_block_accepts_open_bracket_question_number_line(self):
        from core.pdf_exam_parse import parse_quant_block

        items = [
            ("885.(", None),
            (")对于零食相当于(", None),
            (")对于情绪", None),
            ("A．零售脾气", None),
            ("B．儿童病人", None),
            ("C．坚果喜悦", None),
            ("D．食品心情", None),
            ("886.(", None),
            (")对于芯片相当于金刚石对于(", None),
            (")", None),
            ("A．集成电路碳", None),
            ("B．硅探头", None),
            ("C．光刻机石墨", None),
            ("D．半导体金刚砂", None),
        ]

        questions = parse_quant_block(items, 0, len(items))
        self.assertEqual([question.source_number for question in questions], ["885", "886"])

    def test_parse_quant_block_accepts_quoted_option_without_separator(self):
        from core.pdf_exam_parse import parse_quant_block

        items = [
            ("1246.", None),
            ("《原野》:曹禺", None),
            ("A《日出》:老舍", None),
            ("B《寒夜》:巴金", None),
            ("C《月牙儿》:丁玲", None),
            ("D《四世同堂》:钱钟书", None),
        ]

        questions = parse_quant_block(items, 0, len(items))
        self.assertEqual(len(questions), 1)
        self.assertEqual(questions[0].source_number, "1246")
        self.assertEqual(len(questions[0].option_lines), 4)

    def test_parse_quant_block_repairs_spilled_d_option_from_next_question_stem(self):
        from core.pdf_exam_parse import parse_quant_block

        items = [
            ("327.田保姆:指在不改变土地承包关系的前提下,农户将耕、种、管、收等部分或全部作业环节委托给社会化组织完成。", None),
            ("下列不属于田保姆的是:", None),
            ("A.晚稻收割接近尾声,泥瓦匠老李仍然不慌不忙。", None),
            ("已经以每年5000 元的价格租给了邻居。", None),
            ("B.老刘全家商量后,就把田里的农活全部包给了他们。", None),
            ("村里的几个年轻人合伙买了大型收割机。", None),
            ("C.每年收获季,处置成堆的秸秆都让种植户非常头疼。", None),
            ("328.套秸秆打捆、粉碎机械,及时推出了秸秆回收加工服务项目。", None),
            ("D.某农业科技服务公司在各个村镇建立了所管地块的示范田。", None),
            ("当地不少农民变成了“甩手掌柜”。", None),
            ("去雇主化职业:指作为独立个体不受雇于任何雇主。", None),
            ("下列属于去雇主化职业的是:", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        questions = parse_quant_block(items, 0, len(items))
        self.assertEqual([question.source_number for question in questions], ["327", "328"])
        first_stem = "\n".join("".join(text for text, _ in line.parts) for line in questions[0].stem_lines)
        second_stem = "\n".join("".join(text for text, _ in line.parts) for line in questions[1].stem_lines)

        self.assertIn("D.某农业科技服务公司", first_stem)
        self.assertIn("当地不少农民变成了“甩手掌柜”", first_stem)
        self.assertTrue(second_stem.startswith("去雇主化职业:"))
        self.assertNotIn("套秸秆打捆", second_stem)

    def test_parse_quant_block_repairs_malformed_question_head_in_option_tail(self):
        from core.pdf_exam_parse import parse_quant_block

        items = [
            ("1016.下图所代表的物体之间存在一定的逻辑关系,与“图一:图二”逻辑关系最为相近的一项是()。", None),
            ("A.空调:冰箱", None),
            ("B.鲜花:植物", None),
            ("C.飞机:喇叭", None),
            ("D.时钟:时间", None),
            ("10417.与“5YA7HJ9ss”这组符号逻辑关系最为相近的一项是()。", None),
            ("A.6qqW4fgW8", None),
            ("B.OR45r32Ii", None),
            ("C.2fe6xc95W", None),
            ("D.1ee2RR6qQ", None),
            ("1018.每年4 月-5 月上旬,许多北方城市杨柳飞絮飘扬。", None),
            ("A.甲", None),
            ("B.乙", None),
            ("C.丙", None),
            ("D.丁", None),
        ]

        questions = parse_quant_block(items, 0, len(items))
        self.assertEqual([question.source_number for question in questions[:3]], ["1016", "1017", "1018"])
        self.assertEqual(len(questions[0].option_lines), 4)
        repaired_stem = "\n".join("".join(text for text, _ in line.parts) for line in questions[1].stem_lines)
        self.assertIn("5YA7HJ9ss", repaired_stem)
        self.assertEqual(len(questions[1].option_lines), 4)


if __name__ == "__main__":
    unittest.main()
