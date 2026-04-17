import os
from pathlib import Path
import tempfile
import unittest

from PIL import Image, ImageDraw
try:
    import fitz
except ImportError:  # pragma: no cover
    fitz = None

from core.pdf_exam_extract import ExtractedImageRegion
from core.pdf_exam_models import (
    DataAnalysisSection,
    ExamQuestion,
    MaterialUnit,
    ParsedExam,
    PoliticsSection,
    QuantSection,
    ReasoningSection,
    RichLine,
    VerbalSection,
)
from domain.models import AssetRef, OptionNode, QuestionNode
from domain.selectors import parse_question_ranges, select_project
from exporters.pptx_slides import iter_project_question_nodes, project_to_ppt_questions
from ingest.pdf.layout import PageTextLine
from ingest.pdf.project_builder import (
    _copy_asset,
    _split_option_fragments,
    _upgrade_placeholder_text_options_from_stem_asset,
    build_project_from_parsed_exam,
)


def _text_line(text: str) -> RichLine:
    return RichLine(parts=[(text, None)])


def _image_line(path: str) -> RichLine:
    return RichLine(parts=[("", path)])


class ExamProjectTest(unittest.TestCase):
    def test_split_option_fragments_ignores_embedded_markers_in_single_option_lines(self):
        self.assertEqual(
            _split_option_fragments("B．100°C..沸腾"),
            [("B", "100°C..沸腾")],
        )
        self.assertEqual(
            _split_option_fragments("B．A.TP2B.1 蛋白质发挥作用需要胆固醇"),
            [("B", "A.TP2B.1 蛋白质发挥作用需要胆固醇")],
        )
        self.assertEqual(
            _split_option_fragments("B．b3、5a、8n、p1、66"),
            [("B", "b3、5a、8n、p1、66")],
        )
        self.assertEqual(
            _split_option_fragments(
                "C．肉毒杆菌中的A、B、E、F 型会引起人中毒,C 和D 型主要针对畜禽类,对人没有作用;而G 型极为少见,目前还未见中毒报道"
            ),
            [
                (
                    "C",
                    "肉毒杆菌中的A、B、E、F 型会引起人中毒,C 和D 型主要针对畜禽类,对人没有作用;而G 型极为少见,目前还未见中毒报道",
                )
            ],
        )
        self.assertEqual(
            _split_option_fragments(
                "A．农贸市场售卖的富硒土豆，深受消费者欢迎 "
                "B．超市中热卖的药酒，能治疗老年人风湿性关节炎 "
                "C．用于补充婴幼儿维生素 D、促进钙吸收的咀嚼片 "
                "D．李某从某美容整形医院购买的减肥茶，喝后腹泻"
            ),
            [
                ("A", "农贸市场售卖的富硒土豆，深受消费者欢迎"),
                ("B", "超市中热卖的药酒，能治疗老年人风湿性关节炎"),
                ("C", "用于补充婴幼儿维生素 D、促进钙吸收的咀嚼片"),
                ("D", "李某从某美容整形医院购买的减肥茶，喝后腹泻"),
            ],
        )
        self.assertEqual(
            _split_option_fragments(
                "A．x>1..x2>1 B．100°C..沸腾 C．O3..臭氧 D．π..圆面积"
            ),
            [
                ("A", "x>1..x2>1"),
                ("B", "100°C..沸腾"),
                ("C", "O3..臭氧"),
                ("D", "π..圆面积"),
            ],
        )
        self.assertEqual(
            _split_option_fragments(
                "A．426351 B．325146 C．465132 D．354621以下是略有删节的公文部分内容,阅读之后回答1441—1445 题"
            ),
            [
                ("A", "426351"),
                ("B", "325146"),
                ("C", "465132"),
                ("D", "354621"),
            ],
        )
        self.assertEqual(
            _split_option_fragments("C．用于补充婴幼儿维生素D、促进钙吸收的咀嚼片"),
            [("C", "用于补充婴幼儿维生素D、促进钙吸收的咀嚼片")],
        )

    def test_copy_asset_reuses_files_already_in_target_dir(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            asset_path = Path(temp_dir) / "chart.png"
            asset_path.write_bytes(b"fake-image")

            copied_path = _copy_asset(str(asset_path), temp_dir, {})

            self.assertEqual(copied_path, str(asset_path))
            self.assertEqual(sorted(os.listdir(temp_dir)), ["chart.png"])

    def test_upgrade_placeholder_text_options_from_stem_asset(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "choice.png"
            image = Image.new("RGB", (500, 240), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((20, 15, 480, 70), outline="black", width=2)
            draw.text((40, 30), "根据图形判断", fill="black")
            for index, letter in enumerate("ABCD"):
                top = 95 + index * 32
                draw.rectangle((30, top, 470, top + 24), outline="black", width=2)
                draw.text((45, top + 4), f"{letter} 选项图", fill="black")
            image.save(image_path)
            image.close()

            question = QuestionNode(
                source_number="26",
                stem="根据图判断，下列说法正确的是：",
                stem_assets=[AssetRef(kind="stem_image", path=str(image_path), source_page=4)],
                options=[OptionNode(letter=letter, text="如图所示") for letter in "ABCD"],
            )

            repaired = _upgrade_placeholder_text_options_from_stem_asset(question)

            self.assertEqual(len(repaired.options), 4)
            self.assertTrue(all(option.image_path for option in repaired.options))
            self.assertEqual([option.text for option in repaired.options], ["", "", "", ""])
            self.assertTrue(repaired.stem_assets)

    def test_build_project_from_parsed_exam_ignores_following_passage_lines_after_d_option(self):
        exam = ParsedExam(
            verbal_sections=[
                VerbalSection(
                    title="言语理解与表达",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("下面6个句子的最佳顺序是：")],
                            option_lines=[
                                _text_line("A．426351"),
                                _text_line("B．325146"),
                                _text_line("C．465132"),
                                _text_line("D．354621"),
                                _text_line("以下是略有删节的公文部分内容,阅读之后回答1441—1445题"),
                                _text_line("(六)整体布局、协同联动、强化领域应用"),
                            ],
                            source_number="1440",
                        ),
                        ExamQuestion(
                            stem_lines=[_text_line("1441题题干")],
                            option_lines=[_text_line("A．甲"), _text_line("B．乙"), _text_line("C．丙"), _text_line("D．丁")],
                            source_number="1441",
                        ),
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        section = project.sections[0]
        question = section.questions[0]
        self.assertEqual(question.source_number, "1440")
        self.assertEqual([(option.letter, option.text) for option in question.options], [
            ("A", "426351"),
            ("B", "325146"),
            ("C", "465132"),
            ("D", "354621"),
        ])

    def test_build_project_from_parsed_exam_recovers_two_column_combination_options_from_stem(self):
        exam = ParsedExam(
            politics_sections=[
                PoliticsSection(
                    title="政治理论",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("习近平总书记关于网络强国的重要思想，是行动指南。下列属于网信工作的使命任务的是（ ）。"),
                                _text_line("①举旗帜聚民心"),
                                _text_line("②防风险保安全"),
                                _text_line("③强治理惠民生"),
                                _text_line("④讲政治助和谐"),
                                _text_line("⑤增动能促发展"),
                                _text_line("⑥谋合作图共赢"),
                                _text_line("A．12356                                A．12345"),
                                _text_line("C．23456                                B．12456"),
                            ],
                            option_lines=[],
                            source_number="5",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        question = project.sections[0].questions[0]
        self.assertEqual(question.source_number, "5")
        self.assertEqual([(option.letter, option.text) for option in question.options], [
            ("A", "12356"),
            ("B", "12345"),
            ("C", "23456"),
            ("D", "12456"),
        ])

    def test_build_project_from_parsed_exam_includes_general_subjects(self):
        exam = ParsedExam(
            politics_sections=[
                PoliticsSection(
                    title="一. 政治理论",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("1题题干")],
                            option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                            source_number="1",
                        )
                    ],
                )
            ],
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="66",
                        )
                    ],
                )
            ],
            reasoning_sections=[
                ReasoningSection(
                    title="五. 判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("76题题干")],
                            option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                            source_number="76",
                        )
                    ],
                )
            ],
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        self.assertEqual([section.kind for section in project.sections], ["politics", "quant", "reasoning"])
        self.assertEqual(project.question_count, 3)

    def test_build_project_merges_equivalent_objective_section_titles(self):
        exam = ParsedExam(
            reasoning_sections=[
                ReasoningSection(
                    title="判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("76题题干")],
                            option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                            source_number="76",
                        )
                    ],
                ),
                ReasoningSection(
                    title="五. 判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("77题题干")],
                            option_lines=[_text_line("A. 丙"), _text_line("B. 丁")],
                            source_number="77",
                        )
                    ],
                ),
            ],
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        self.assertEqual(len(project.sections), 1)
        self.assertEqual(project.sections[0].kind, "reasoning")
        self.assertEqual(project.sections[0].title, "五. 判断推理")
        self.assertEqual([question.source_number for question in project.sections[0].questions], ["76", "77"])

    def test_build_project_from_parsed_exam_preserves_material_regions_and_question_numbers(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="2026年·天津·资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("材料正文第一段")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[
                                        _text_line("A. 甲"),
                                        _text_line("B. 乙"),
                                    ],
                                    source_number="111",
                                )
                            ],
                        )
                    ],
                )
            ],
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[
                                _text_line("A. 1"),
                                _text_line("B. 2"),
                            ],
                            source_number="66",
                        )
                    ],
                )
            ],
        )
        layout = [
            PageTextLine(text="66题题干", page_number=1, x0=10, y0=20, x1=80, y1=32),
            PageTextLine(text="材料一", page_number=2, x0=10, y0=20, x1=40, y1=30),
            PageTextLine(text="材料正文第一段", page_number=2, x0=10, y0=32, x1=80, y1=45),
            PageTextLine(text="111题题干", page_number=2, x0=10, y0=60, x1=90, y1=72),
        ]

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            layout_lines=layout,
        )

        self.assertEqual(project.question_count, 2)
        self.assertEqual(project.sections[0].questions[0].source_number, "66")
        material = project.sections[1].material_sets[0]
        self.assertEqual(material.body_regions[0].page_number, 2)
        self.assertEqual(material.questions[0].source_number, "111")

    def test_select_project_filters_by_question_range(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="66",
                        ),
                        ExamQuestion(
                            stem_lines=[_text_line("67题题干")],
                            option_lines=[_text_line("A. 3"), _text_line("B. 4")],
                            source_number="67",
                        ),
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")
        filtered = select_project(project, subjects=["quant"], question_ranges=parse_question_ranges("67-67"))

        self.assertEqual(filtered.question_count, 1)
        self.assertEqual(filtered.sections[0].questions[0].source_number, "67")

    def test_project_to_ppt_questions_uses_material_images_for_data_questions(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("材料正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="111",
                                )
                            ],
                        )
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        questions = project_to_ppt_questions(
            project,
            material_image_map={"data-1-1": ["material.png"]},
        )

        self.assertEqual(len(questions), 1)
        self.assertEqual(questions[0].image_paths, ["material.png"])
        self.assertIsNone(questions[0].material_text)
        self.assertEqual(questions[0].source_question_number, "111")

    def test_project_to_ppt_questions_preserves_question_option_layout(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="66",
                        )
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")
        project.sections[0].questions[0].option_layout = "one_row"

        questions = project_to_ppt_questions(project)

        self.assertEqual(len(questions), 1)
        self.assertEqual(questions[0].option_layout, "one_row")

    def test_iter_project_question_nodes_matches_ppt_export_order(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="66",
                        )
                    ],
                )
            ],
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("材料正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="111",
                                )
                            ],
                        )
                    ],
                )
            ],
        )
        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        node_order = [
            (section.kind, getattr(material, "header", None), question.source_number)
            for section, material, question in iter_project_question_nodes(project)
        ]
        question_order = [
            question.source_question_number
            for question in project_to_ppt_questions(project)
        ]

        self.assertEqual(
            node_order,
            [
                ("quant", None, "66"),
                ("data", "材料一", "111"),
            ],
        )
        self.assertEqual(question_order, ["66", "111"])

    def test_build_project_splits_two_column_option_lines_into_four_options(self):
        exam = ParsedExam(
            verbal_sections=[
                VerbalSection(
                    title="第三部分 言语理解与表达",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("26题题干")],
                            option_lines=[
                                _text_line("A．甲\t\tB．乙"),
                                _text_line("C．丙\t\tD．丁"),
                            ],
                            source_number="26",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        question = project.sections[0].questions[0]
        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual([option.text for option in question.options], ["甲", "乙", "丙", "丁"])

    def test_build_project_assigns_images_before_option_labels_to_options(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="第四部分 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("65题题干")],
                            option_lines=[
                                _image_line("a.png"),
                                _image_line("b.png"),
                                _image_line("c.png"),
                                _image_line("d.png"),
                                _text_line("A."),
                                _text_line("B."),
                                _text_line("C."),
                                _text_line("D."),
                            ],
                            source_number="65",
                        )
                    ],
                )
            ]
        )
        image_regions = {
            "a.png": ExtractedImageRegion("a.png", 3, 0, 0, 10, 10),
            "b.png": ExtractedImageRegion("b.png", 3, 10, 0, 20, 10),
            "c.png": ExtractedImageRegion("c.png", 3, 20, 0, 30, 10),
            "d.png": ExtractedImageRegion("d.png", 3, 30, 0, 40, 10),
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            image_regions=image_regions,
        )

        question = project.sections[0].questions[0]
        self.assertEqual(len(question.stem_assets), 0)
        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual([option.image_path for option in question.options], ["a.png", "b.png", "c.png", "d.png"])
        self.assertEqual([option.source_page for option in question.options], [3, 3, 3, 3])

    def test_build_project_rebalances_trailing_image_labels_back_to_all_four_options(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="六. 资料分析:",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[
                                        _text_line("115题题干"),
                                        _image_line("a.png"),
                                    ],
                                    option_lines=[
                                        _text_line("A."),
                                        _image_line("b.png"),
                                        _text_line("B."),
                                        _image_line("c.png"),
                                        _text_line("C."),
                                        _image_line("d.png"),
                                        _text_line("D."),
                                    ],
                                    source_number="115",
                                )
                            ],
                        )
                    ],
                )
            ]
        )
        image_regions = {
            "a.png": ExtractedImageRegion("a.png", 28, 0, 0, 10, 10),
            "b.png": ExtractedImageRegion("b.png", 28, 10, 0, 20, 10),
            "c.png": ExtractedImageRegion("c.png", 28, 20, 0, 30, 10),
            "d.png": ExtractedImageRegion("d.png", 28, 30, 0, 40, 10),
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            image_regions=image_regions,
        )

        question = project.sections[0].material_sets[0].questions[0]
        self.assertEqual(len(question.stem_assets), 0)
        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual([option.image_path for option in question.options], ["a.png", "b.png", "c.png", "d.png"])
        self.assertEqual([option.source_page for option in question.options], [28, 28, 28, 28])

    def test_build_project_promotes_trailing_option_images_to_next_question_stem(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="第四部分 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("77题题干")],
                            option_lines=[
                                _text_line("A.50.25"),
                                _text_line("B.53.26"),
                                _text_line("C.55.17"),
                                _text_line("D.56.30"),
                                _image_line("q78_stem.png"),
                            ],
                            source_number="77",
                        ),
                        ExamQuestion(
                            stem_lines=[],
                            option_lines=[
                                _image_line("a.png"),
                                _text_line("A."),
                                _image_line("b.png"),
                                _text_line("B."),
                                _image_line("c.png"),
                                _text_line("C."),
                                _image_line("d.png"),
                                _text_line("D."),
                            ],
                            source_number="78",
                        ),
                        ExamQuestion(
                            stem_lines=[_text_line("79题题干")],
                            option_lines=[
                                _text_line("A.-1"),
                                _text_line("B.0"),
                                _text_line("C.1"),
                                _text_line("D.2"),
                                _image_line("q80_stem.png"),
                            ],
                            source_number="79",
                        ),
                        ExamQuestion(
                            stem_lines=[],
                            option_lines=[
                                _text_line("A.14"),
                                _text_line("B.17"),
                                _text_line("C.19"),
                                _text_line("D.21"),
                            ],
                            source_number="80",
                        ),
                    ],
                )
            ]
        )
        image_regions = {
            "q78_stem.png": ExtractedImageRegion("q78_stem.png", 11, 0, 0, 10, 10),
            "a.png": ExtractedImageRegion("a.png", 11, 10, 0, 20, 10),
            "b.png": ExtractedImageRegion("b.png", 11, 20, 0, 30, 10),
            "c.png": ExtractedImageRegion("c.png", 11, 30, 0, 40, 10),
            "d.png": ExtractedImageRegion("d.png", 11, 40, 0, 50, 10),
            "q80_stem.png": ExtractedImageRegion("q80_stem.png", 12, 0, 0, 10, 10),
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            image_regions=image_regions,
        )

        q77, q78, q79, q80 = project.sections[0].questions
        self.assertIsNone(q77.options[-1].image_path)
        self.assertEqual([asset.path for asset in q78.stem_assets], ["q78_stem.png"])
        self.assertEqual([option.image_path for option in q78.options], ["a.png", "b.png", "c.png", "d.png"])
        self.assertIsNone(q79.options[-1].image_path)
        self.assertEqual([asset.path for asset in q80.stem_assets], ["q80_stem.png"])
        self.assertEqual([option.text for option in q80.options], ["14", "17", "19", "21"])

    def test_build_project_splits_trailing_option_image_asset_into_four_choices(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            stem_path = Path(temp_dir) / "stem.png"
            stem_image = Image.new("RGB", (120, 60), "white")
            ImageDraw.Draw(stem_image).rectangle((12, 18, 100, 42), fill="black")
            stem_image.save(stem_path)

            options_path = Path(temp_dir) / "options.png"
            option_image = Image.new("RGB", (120, 200), "white")
            draw = ImageDraw.Draw(option_image)
            for index in range(4):
                top = 12 + index * 46
                draw.rectangle((14, top, 106, top + 20), fill="black")
            option_image.save(options_path)

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_image_line(str(stem_path)), _image_line(str(options_path))],
                                option_lines=[],
                                source_number="22",
                            )
                        ],
                    )
                ]
            )

            project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

            question = project.sections[0].questions[0]
            self.assertEqual([Path(asset.path).name for asset in question.stem_assets], ["stem.png"])
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))

    def test_build_project_splits_partial_option_images_across_two_assets(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            first_path = Path(temp_dir) / "first.png"
            first_image = Image.new("RGB", (260, 170), "white")
            draw = ImageDraw.Draw(first_image)
            draw.rectangle((18, 18, 242, 48), fill="black")
            draw.rectangle((18, 74, 118, 100), fill="black")
            draw.rectangle((18, 122, 118, 148), fill="black")
            first_image.save(first_path)

            second_path = Path(temp_dir) / "second.png"
            second_image = Image.new("RGB", (120, 110), "white")
            draw = ImageDraw.Draw(second_image)
            draw.rectangle((12, 18, 98, 42), fill="black")
            draw.rectangle((12, 64, 98, 88), fill="black")
            second_image.save(second_path)

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_image_line(str(first_path)), _image_line(str(second_path))],
                                option_lines=[],
                                source_number="145",
                            )
                        ],
                    )
                ]
            )

            project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

            question = project.sections[0].questions[0]
            self.assertEqual(len(question.stem_assets), 1)
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))

    def test_build_project_splits_grid_option_image_into_four_choices(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            stem_path = Path(temp_dir) / "stem.png"
            stem_image = Image.new("RGB", (120, 120), "white")
            ImageDraw.Draw(stem_image).ellipse((20, 20, 100, 100), outline="black", width=4)
            stem_image.save(stem_path)

            grid_path = Path(temp_dir) / "grid.png"
            grid_image = Image.new("RGB", (260, 180), "white")
            draw = ImageDraw.Draw(grid_image)
            draw.rectangle((20, 18, 100, 68), fill="black")
            draw.rectangle((150, 18, 230, 68), fill="black")
            draw.rectangle((20, 108, 100, 158), fill="black")
            draw.rectangle((150, 108, 230, 158), fill="black")
            grid_image.save(grid_path)

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_text_line("图形题题干"), _image_line(str(stem_path)), _image_line(str(grid_path))],
                                option_lines=[],
                                source_number="537",
                            )
                        ],
                    )
                ]
            )

            project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

            question = project.sections[0].questions[0]
            self.assertEqual([Path(asset.path).name for asset in question.stem_assets], ["stem.png"])
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))

    def test_build_project_extracts_inline_options_with_dash_marker(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="第四部分 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("50,25,12,13,0,-1,()。A-3 B.1 C.7 D.14")],
                            option_lines=[],
                            source_number="697",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        question = project.sections[0].questions[0]
        self.assertEqual(question.stem, "50,25,12,13,0,-1,()。")
        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual([option.text for option in question.options], ["3", "1", "7", "14"])

    def test_build_project_extracts_options_from_single_image_only_question(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "full.png"
            image = Image.new("RGB", (360, 120), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((20, 16, 220, 44), fill="black")
            draw.rectangle((18, 68, 78, 96), fill="black")
            draw.rectangle((104, 68, 164, 96), fill="black")
            draw.rectangle((190, 68, 250, 96), fill="black")
            draw.rectangle((276, 68, 336, 96), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_image_line(str(image_path))],
                                option_lines=[],
                                source_number="4",
                            )
                        ],
                    )
                ]
            )

            project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

            question = project.sections[0].questions[0]
            self.assertEqual([Path(asset.path).name for asset in question.stem_assets], ["full_stem.png"])
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))

    def test_build_project_keeps_stem_prefix_when_single_image_rows_are_tightly_packed(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "packed.png"
            image = Image.new("RGB", (360, 150), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((32, 18, 228, 30), fill="black")
            draw.rectangle((26, 36, 96, 54), fill="black")
            draw.rectangle((26, 62, 96, 80), fill="black")
            draw.rectangle((26, 88, 96, 106), fill="black")
            draw.rectangle((26, 114, 96, 132), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_image_line(str(image_path))],
                                option_lines=[],
                                source_number="145",
                            )
                        ],
                    )
                ]
            )

            project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

            question = project.sections[0].questions[0]
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertEqual(len(question.stem_assets), 1)
            self.assertTrue(question.stem_assets[0].path.endswith("_stem.png"))

    def test_build_project_keeps_stem_prefix_when_single_image_options_merge_into_one_band(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            image_path = Path(temp_dir) / "merged_band.png"
            image = Image.new("RGB", (220, 180), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((26, 10, 188, 28), fill="black")
            for index in range(4):
                top = 48 + index * 25
                draw.rectangle((8, top, 44, top + 18), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_image_line(str(image_path))],
                                option_lines=[],
                                source_number="353",
                            )
                        ],
                    )
                ]
            )

            project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

            question = project.sections[0].questions[0]
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertEqual(len(question.stem_assets), 1)
            self.assertTrue(question.stem_assets[0].path.endswith("_stem.png"))

    def test_build_project_reassigns_stem_assets_to_blank_options(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="第四部分 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("图形题题干"),
                                _image_line("c.png"),
                                _image_line("d.png"),
                            ],
                            option_lines=[
                                _text_line("A.10"),
                                _text_line("B.12"),
                                _text_line("C."),
                                _text_line("D."),
                            ],
                            source_number="137",
                        )
                    ],
                )
            ]
        )
        image_regions = {
            "c.png": ExtractedImageRegion("c.png", 17, 0, 0, 10, 10),
            "d.png": ExtractedImageRegion("d.png", 17, 10, 0, 20, 10),
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            image_regions=image_regions,
        )

        question = project.sections[0].questions[0]
        self.assertEqual(question.stem_assets, [])
        self.assertIsNone(question.options[0].image_path)
        self.assertIsNone(question.options[1].image_path)
        self.assertEqual([Path(question.options[2].image_path).name, Path(question.options[3].image_path).name], ["c.png", "d.png"])

    def test_build_project_shifts_option_images_forward_to_blank_slots(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="第四部分 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("带图片选项题干")],
                            option_lines=[
                                _text_line("A.1:2"),
                                _text_line("B.5:7"),
                                _image_line("c.png"),
                                _text_line("C."),
                                _image_line("d.png"),
                                _text_line("D."),
                            ],
                            source_number="67",
                        )
                    ],
                )
            ]
        )
        image_regions = {
            "c.png": ExtractedImageRegion("c.png", 10, 0, 0, 10, 10),
            "d.png": ExtractedImageRegion("d.png", 10, 10, 0, 20, 10),
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            image_regions=image_regions,
        )

        question = project.sections[0].questions[0]
        self.assertIsNone(question.options[1].image_path)
        self.assertEqual(Path(question.options[2].image_path).name, "c.png")
        self.assertEqual(Path(question.options[3].image_path).name, "d.png")

    @unittest.skipIf(fitz is None, "PyMuPDF not installed")
    def test_build_project_crops_blank_option_images_from_pdf(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            pdf_path = Path(temp_dir) / "blank_option_crops.pdf"
            doc = fitz.open()
            page = doc.new_page(width=600, height=800)
            page.draw_rect(fitz.Rect(220, 468, 252, 486), color=(0, 0, 0), fill=(0, 0, 0))
            page.draw_rect(fitz.Rect(350, 468, 382, 486), color=(0, 0, 0), fill=(0, 0, 0))
            doc.save(pdf_path)
            doc.close()

            exam = ParsedExam(
                quant_sections=[
                    QuantSection(
                        title="第四部分 数量关系",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_text_line("几何题干")],
                                option_lines=[
                                    _text_line("A．1"),
                                    _text_line("B．"),
                                    _text_line("C．"),
                                    _text_line("D．2"),
                                ],
                                source_number="305",
                            )
                        ],
                    )
                ]
            )
            layout_lines = [
                PageTextLine("几何题干", 1, 75.0, 440.0, 180.0, 450.0, 75.0, 440.0, 180.0, 450.0),
                PageTextLine("A.1", 1, 75.0, 473.0, 91.0, 484.0, 75.0, 473.0, 91.0, 484.0),
                PageTextLine("B.", 1, 201.0, 473.0, 211.0, 484.0, 201.0, 473.0, 211.0, 484.0),
                PageTextLine("C.", 1, 327.0, 473.0, 338.0, 484.0, 327.0, 473.0, 338.0, 484.0),
                PageTextLine("D.2", 1, 448.0, 473.0, 464.0, 484.0, 448.0, 473.0, 464.0, 484.0),
            ]

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path=str(pdf_path),
                layout_lines=layout_lines,
            )

            question = project.sections[0].questions[0]
            self.assertTrue(question.options[1].image_path)
            self.assertTrue(question.options[2].image_path)

    def test_build_project_normalizes_duplicate_option_letters_from_pdf_order_noise(self):
        exam = ParsedExam(
            politics_sections=[
                PoliticsSection(
                    title="第一部分 政治理论",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("5题题干")],
                            option_lines=[
                                _text_line("A．12356\t\tA．12345"),
                                _text_line("C．23456\t\tB．12456"),
                            ],
                            source_number="5",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        question = project.sections[0].questions[0]
        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual([option.text for option in question.options], ["12356", "12345", "23456", "12456"])

    def test_build_project_flattens_soft_wrapped_stem_lines(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("这是第一行"), _text_line("这是第二行")],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="66",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        self.assertEqual(project.sections[0].questions[0].stem, "这是第一行这是第二行")

    def test_build_project_preserves_numbered_stem_clauses_and_blank_prefix(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line(",严格环境执法是环境法治建设的重要内容。"),
                                _text_line("1持续推进重点领域整治"),
                                _text_line("2提高治理效能"),
                            ],
                            option_lines=[
                                _text_line("A. 甲"),
                                _text_line("B. 乙"),
                                _text_line("C. 丙"),
                                _text_line("D. 丁"),
                            ],
                            source_number="66",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        self.assertEqual(
            project.sections[0].questions[0].stem,
            "____,严格环境执法是环境法治建设的重要内容。\n1持续推进重点领域整治\n2提高治理效能",
        )

    def test_build_project_separates_prompt_after_enumerated_stem_clauses(self):
        exam = ParsedExam(
            verbal_sections=[
                VerbalSection(
                    title="第三部分 言语理解与表达",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("1第一句"),
                                _text_line("2第二句"),
                                _text_line("将以上2个句子重新排列,语序正确的一项是:"),
                            ],
                            option_lines=[
                                _text_line("A．12"),
                                _text_line("B．21"),
                                _text_line("C．13"),
                                _text_line("D．23"),
                            ],
                            source_number="59",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        self.assertEqual(
            project.sections[0].questions[0].stem,
            "1第一句\n2第二句\n将以上2个句子重新排列,语序正确的一项是:",
        )

    def test_build_project_repairs_fraction_like_stem_fragments(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="第四部分 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("甲的速度是乙的4"),
                                _text_line("5,当甲走了全程的1"),
                                _text_line("3时,甲、乙相距75米。"),
                            ],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="66",
                        ),
                        ExamQuestion(
                            stem_lines=[
                                _text_line("B班走读的学生人数占该班总人数的19"),
                                _text_line("45,C班走读的学生人数比寄宿的少1"),
                                _text_line("12,则共有多少名学生?"),
                            ],
                            option_lines=[_text_line("A. 1"), _text_line("B. 2")],
                            source_number="67",
                        ),
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")

        stems = [question.stem for question in project.sections[0].questions]
        self.assertIn("4/5,当甲走了全程的1/3时", stems[0])
        self.assertIn("19/45,C班走读的学生人数比寄宿的少1/2,则共有多少名学生?", stems[1])

    def test_build_project_prefers_block_bounds_for_material_regions(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("表格第一行"), _text_line("表格第二行")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="111",
                                )
                            ],
                        )
                    ],
                )
            ]
        )
        layout = [
            PageTextLine(
                text="材料一",
                page_number=2,
                x0=40,
                y0=20,
                x1=80,
                y1=30,
                block_x0=12,
                block_y0=18,
                block_x1=260,
                block_y1=44,
            ),
            PageTextLine(
                text="表格第一行",
                page_number=2,
                x0=80,
                y0=60,
                x1=120,
                y1=72,
                block_x0=12,
                block_y0=52,
                block_x1=260,
                block_y1=96,
            ),
            PageTextLine(
                text="表格第二行",
                page_number=2,
                x0=85,
                y0=100,
                x1=125,
                y1=112,
                block_x0=12,
                block_y0=96,
                block_x1=260,
                block_y1=140,
            ),
            PageTextLine(text="111题题干", page_number=2, x0=10, y0=160, x1=90, y1=172),
        ]

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            layout_lines=layout,
        )

        material = project.sections[0].material_sets[0]
        self.assertEqual(material.body_regions[0].x0, 12)
        self.assertEqual(material.body_regions[0].x1, 260)
        self.assertEqual(material.body_regions[0].y0, 18)
        self.assertEqual(material.body_regions[0].y1, 140)

    def test_build_project_merges_material_image_regions_into_body_regions(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("表格标题"), _image_line("chart.png")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="111",
                                )
                            ],
                        )
                    ],
                )
            ]
        )
        layout = [
            PageTextLine(
                text="材料一",
                page_number=2,
                x0=10,
                y0=20,
                x1=40,
                y1=30,
                block_x0=10,
                block_y0=20,
                block_x1=140,
                block_y1=44,
            ),
            PageTextLine(
                text="表格标题",
                page_number=2,
                x0=10,
                y0=50,
                x1=60,
                y1=62,
                block_x0=10,
                block_y0=48,
                block_x1=140,
                block_y1=72,
            ),
            PageTextLine(text="111题题干", page_number=2, x0=10, y0=170, x1=90, y1=182),
        ]
        image_regions = {
            "chart.png": ExtractedImageRegion(
                path="chart.png",
                page_number=2,
                x0=150,
                y0=80,
                x1=320,
                y1=200,
            )
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            layout_lines=layout,
            image_regions=image_regions,
        )

        material = project.sections[0].material_sets[0]
        self.assertEqual(material.body_assets[0].source_page, 2)
        self.assertIsNotNone(material.body_assets[0].page_region)
        self.assertEqual(material.body_regions[0].x0, 10)
        self.assertEqual(material.body_regions[0].x1, 320)
        self.assertEqual(material.body_regions[0].y0, 20)
        self.assertEqual(material.body_regions[0].y1, 200)

    def test_build_project_attaches_orphan_continuation_page_image_to_material(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料二",
                            intro_lines=[_text_line("第一页材料正文"), _text_line("第二页续行文字")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("106题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="106",
                                )
                            ],
                        )
                    ],
                )
            ]
        )
        layout = [
            PageTextLine(
                text="材料二",
                page_number=1,
                x0=10,
                y0=20,
                x1=40,
                y1=30,
                block_x0=10,
                block_y0=20,
                block_x1=160,
                block_y1=44,
            ),
            PageTextLine(
                text="第一页材料正文",
                page_number=1,
                x0=10,
                y0=50,
                x1=100,
                y1=62,
                block_x0=10,
                block_y0=48,
                block_x1=220,
                block_y1=74,
            ),
            PageTextLine(
                text="第二页续行文字",
                page_number=2,
                x0=18,
                y0=132,
                x1=110,
                y1=146,
                block_x0=18,
                block_y0=128,
                block_x1=240,
                block_y1=154,
            ),
            PageTextLine(
                text="106题题干",
                page_number=2,
                x0=18,
                y0=282,
                x1=96,
                y1=296,
                block_x0=18,
                block_y0=280,
                block_x1=180,
                block_y1=306,
            ),
        ]
        image_regions = {
            "continued-chart.png": ExtractedImageRegion(
                path="continued-chart.png",
                page_number=2,
                x0=22,
                y0=24,
                x1=236,
                y1=118,
            )
        }

        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path="sample.pdf",
            layout_lines=layout,
            image_regions=image_regions,
        )

        material = project.sections[0].material_sets[0]
        self.assertEqual(len(material.body_assets), 1)
        self.assertEqual(material.body_assets[0].path, "continued-chart.png")
        page2_region = [region for region in material.body_regions if region.page_number == 2][0]
        self.assertEqual(page2_region.y0, 24)
        self.assertEqual(page2_region.y1, 154)


if __name__ == "__main__":
    unittest.main()
