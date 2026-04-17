import tempfile
import unittest
from pathlib import Path

from PIL import Image, ImageDraw

from core.pdf_exam_extract import ExtractedImageRegion
from core.pdf_exam_models import DataAnalysisSection, ExamQuestion, MaterialUnit, ParsedExam, QuantSection, ReasoningSection, RichLine
from ingest.pdf.project_builder import build_project_from_parsed_exam


def _text_line(text: str) -> RichLine:
    return RichLine(parts=[(text, None)])


def _image_line(path: str) -> RichLine:
    return RichLine(parts=[("", path)])


class PdfProjectBuilderAutofixTest(unittest.TestCase):
    def test_extracts_inline_options_from_stem_text(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料（回答81-85题）",
                            intro_lines=[_text_line("材料正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("的多少倍?A.0.5B.0.8C.1.3D.2.1")],
                                    option_lines=[],
                                    source_number="81",
                                )
                            ],
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")
        question = project.sections[0].material_sets[0].questions[0]

        self.assertEqual(question.stem, "的多少倍?")
        self.assertEqual([option.text for option in question.options], ["0.5", "0.8", "1.3", "2.1"])

    def test_extracts_multiline_tail_options_from_stem_text(self):
        exam = ParsedExam(
            reasoning_sections=[
                ReasoningSection(
                    title="判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("有人说喝隔夜茶会致癌,下列选项如果为真,最能质疑该说法的是()。"),
                                _text_line("A 只有过量摄取且身体对缺乏维生素c 的情况下才能对人体造成危害"),
                                _text_line("B 生成亚硝胺的基础物质亚硝酸盐、硝酸盐、和胺类在食物中是普遍存在的"),
                                _text_line("C 亚硝胺是重要的化学致癌物之一,也是四大食品污染物之一"),
                                _text_line("D 茶叶富含维生素c 和茶多酚,是合成亚硝胺的天然抑制剂"),
                            ],
                            option_lines=[],
                            source_number="1020",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")
        question = project.sections[0].questions[0]

        self.assertEqual(question.stem, "有人说喝隔夜茶会致癌,下列选项如果为真,最能质疑该说法的是()。")
        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual(question.options[0].text[:4], "只有过量")
        self.assertIn("天然抑制剂", question.options[-1].text)

    def test_extracts_tail_options_without_confusing_earlier_abcd_variables(self):
        exam = ParsedExam(
            reasoning_sections=[
                ReasoningSection(
                    title="判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("某派出所计划派出甲、乙、丙、丁、戊、己六位民警到A、B、C 三个区域去巡逻。"),
                                _text_line("如果己民警去A 区域巡逻,则下列一定为真的是()。"),
                                _text_line("A、丙民警去B 区域"),
                                _text_line("B、戊民警去C 区域"),
                                _text_line("C、甲民警去C 区域D、丁民警去B 区域"),
                            ],
                            option_lines=[],
                            source_number="759",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")
        question = project.sections[0].questions[0]

        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual(question.options[2].text, "甲民警去C 区域")
        self.assertEqual(question.options[3].text, "丁民警去B 区域")

    def test_does_not_split_entity_letters_inside_option_body(self):
        exam = ParsedExam(
            reasoning_sections=[
                ReasoningSection(
                    title="判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("从上述条件中,可以推出()。")],
                            option_lines=[
                                _text_line("A．B、C、D 不参会"),
                                _text_line("B．A、B、C 参会"),
                                _text_line("C．D、C、E 不参会"),
                                _text_line("D．A、B、E 参会"),
                            ],
                            source_number="951",
                        )
                    ],
                )
            ]
        )

        project = build_project_from_parsed_exam(exam, source_pdf_path="sample.pdf")
        question = project.sections[0].questions[0]

        self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
        self.assertEqual(question.options[0].text, "B、C、D 不参会")
        self.assertEqual(question.options[2].text, "D、C、E 不参会")

    def test_moves_four_stem_assets_into_image_options(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[
                                _text_line("下图中最符合题意的是:"),
                                _image_line("a.png"),
                                _image_line("b.png"),
                                _image_line("c.png"),
                                _image_line("d.png"),
                            ],
                            option_lines=[],
                            source_number="43",
                        )
                    ],
                )
            ]
        )
        image_regions = {
            "a.png": ExtractedImageRegion("a.png", 7, 0, 0, 10, 10),
            "b.png": ExtractedImageRegion("b.png", 7, 10, 0, 20, 10),
            "c.png": ExtractedImageRegion("c.png", 7, 20, 0, 30, 10),
            "d.png": ExtractedImageRegion("d.png", 7, 30, 0, 40, 10),
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

    def test_splits_vertical_choice_image_into_four_options(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "vertical.png"
            image = Image.new("RGB", (220, 360), "white")
            draw = ImageDraw.Draw(image)
            for idx in range(4):
                top = 20 + idx * 85
                draw.rectangle((30, top, 190, top + 45), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                data_sections=[
                    DataAnalysisSection(
                        title="资料分析",
                        materials=[
                            MaterialUnit(
                                header="材料（回答21-25题）",
                                intro_lines=[_text_line("材料正文")],
                                questions=[
                                    ExamQuestion(
                                        stem_lines=[
                                            _text_line("以下图形中,最符合题意的是:"),
                                            _image_line(str(image_path)),
                                        ],
                                        option_lines=[],
                                        source_number="25",
                                    )
                                ],
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(image_path): ExtractedImageRegion(str(image_path), 5, 0, 0, 220, 360),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            question = project.sections[0].material_sets[0].questions[0]

            self.assertEqual(len(question.stem_assets), 0)
            self.assertEqual(len(question.options), 4)
            self.assertTrue(all(option.image_path for option in question.options))

    def test_splits_two_stacked_choice_images_into_four_options(self):
        with tempfile.TemporaryDirectory() as tmp:
            left_path = Path(tmp) / "left.png"
            right_path = Path(tmp) / "right.png"
            for path in (left_path, right_path):
                image = Image.new("RGB", (220, 320), "white")
                draw = ImageDraw.Draw(image)
                draw.rectangle((30, 20, 190, 125), fill="black")
                draw.rectangle((30, 180, 190, 285), fill="black")
                image.save(path)

            exam = ParsedExam(
                data_sections=[
                    DataAnalysisSection(
                        title="资料分析",
                        materials=[
                            MaterialUnit(
                                header="材料（回答170-170题）",
                                intro_lines=[_text_line("材料正文")],
                                questions=[
                                    ExamQuestion(
                                        stem_lines=[
                                            _text_line("下列选项最符合题意的是:"),
                                            _image_line(str(left_path)),
                                            _image_line(str(right_path)),
                                        ],
                                        option_lines=[],
                                        source_number="170",
                                    )
                                ],
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(left_path): ExtractedImageRegion(str(left_path), 5, 0, 0, 220, 320),
                str(right_path): ExtractedImageRegion(str(right_path), 5, 240, 0, 460, 320),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            question = project.sections[0].material_sets[0].questions[0]

            self.assertLessEqual(len(question.stem_assets), 1)
            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))

    def test_extracts_options_from_nonfinal_stem_asset_when_trailing_noise_asset_exists(self):
        with tempfile.TemporaryDirectory() as tmp:
            option_image = Path(tmp) / "options.png"
            noise_image = Path(tmp) / "noise.png"

            image = Image.new("RGB", (320, 180), "white")
            draw = ImageDraw.Draw(image)
            for idx in range(4):
                top = 20 + idx * 35
                draw.rectangle((40, top, 260, top + 18), fill="black")
            image.save(option_image)

            noise = Image.new("RGB", (60, 20), "black")
            noise.save(noise_image)

            exam = ParsedExam(
                reasoning_sections=[
                    ReasoningSection(
                        title="判断推理",
                        questions=[
                            ExamQuestion(
                                stem_lines=[
                                    _text_line("下列化学变化能够实现降碳目标的是:"),
                                    _image_line(str(option_image)),
                                    _image_line(str(noise_image)),
                                ],
                                option_lines=[],
                                source_number="43",
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(option_image): ExtractedImageRegion(str(option_image), 8, 0, 0, 320, 180),
                str(noise_image): ExtractedImageRegion(str(noise_image), 8, 0, 184, 60, 204),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            question = project.sections[0].questions[0]

            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))

    def test_promotes_trailing_stem_asset_to_previous_visual_question(self):
        with tempfile.TemporaryDirectory() as tmp:
            context_image = Path(tmp) / "context.png"
            option_image = Path(tmp) / "options.png"

            image = Image.new("RGB", (220, 120), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((30, 20, 190, 90), fill="black")
            image.save(context_image)

            image = Image.new("RGB", (220, 360), "white")
            draw = ImageDraw.Draw(image)
            for idx in range(4):
                top = 20 + idx * 85
                draw.rectangle((30, top, 190, top + 45), fill="black")
            image.save(option_image)

            exam = ParsedExam(
                reasoning_sections=[
                    ReasoningSection(
                        title="判断推理",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_text_line("从所给的四个选项中,选择最合适的一个填入问号处,使之呈现一定的规律性:")],
                                option_lines=[],
                                source_number="1383",
                            ),
                            ExamQuestion(
                                stem_lines=[
                                    _text_line("下图为给定的多面体及其外表面展开图,问对应关系为:"),
                                    _image_line(str(context_image)),
                                    _image_line(str(option_image)),
                                ],
                                option_lines=[
                                    _text_line("A.1-C,2-A,3-B,4-D"),
                                    _text_line("B.1-A,2-C,3-B,4-D"),
                                    _text_line("C.1-A,2-C,3-D,4-B"),
                                    _text_line("D.1-C,2-A,3-D,4-B"),
                                ],
                                source_number="1384",
                            ),
                        ],
                    )
                ]
            )
            image_regions = {
                str(context_image): ExtractedImageRegion(str(context_image), 254, 0, 0, 220, 120),
                str(option_image): ExtractedImageRegion(str(option_image), 254, 0, 140, 220, 500),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            previous, current = project.sections[0].questions[:2]

            self.assertEqual(previous.source_number, "1383")
            self.assertEqual([option.letter for option in previous.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in previous.options))
            self.assertEqual(len(previous.stem_assets), 0)

            self.assertEqual(current.source_number, "1384")
            self.assertEqual(len(current.stem_assets), 1)
            self.assertTrue(current.stem_assets[0].path.endswith("context.png"))
            self.assertEqual([option.text for option in current.options], ["1-C,2-A,3-B,4-D", "1-A,2-C,3-B,4-D", "1-A,2-C,3-D,4-B", "1-C,2-A,3-D,4-B"])

    def test_synthesizes_labeled_point_options_for_diagram_prompt(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "diagram.png"
            image = Image.new("RGB", (200, 220), "white")
            draw = ImageDraw.Draw(image)
            draw.ellipse((30, 90, 120, 200), outline="black", width=3)
            draw.line((80, 90, 80, 30), fill="black", width=3)
            image.save(image_path)

            exam = ParsedExam(
                reasoning_sections=[
                    ReasoningSection(
                        title="判断推理",
                        questions=[
                            ExamQuestion(
                                stem_lines=[
                                    _text_line("下图中卫星位于点时,对于O 点的人来说仰角为90°。"),
                                    _image_line(str(image_path)),
                                ],
                                option_lines=[],
                                source_number="674",
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(image_path): ExtractedImageRegion(str(image_path), 127, 0, 0, 200, 220),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            question = project.sections[0].questions[0]

            self.assertEqual([option.text for option in question.options], ["A点", "B点", "C点", "D点"])

    def test_splits_combined_choice_image_into_stem_and_options(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "combined.png"
            image = Image.new("RGB", (420, 180), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((30, 20, 390, 65), fill="black")
            for idx in range(4):
                left = 30 + idx * 95
                draw.rectangle((left, 105, left + 55, 150), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                reasoning_sections=[
                    ReasoningSection(
                        title="判断推理",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_image_line(str(image_path))],
                                option_lines=[],
                                source_number="1",
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(image_path): ExtractedImageRegion(str(image_path), 1, 0, 0, 420, 180),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            question = project.sections[0].questions[0]

            self.assertEqual(len(question.stem_assets), 1)
            self.assertEqual(len(question.options), 4)
            self.assertTrue(all(option.image_path for option in question.options))

    def test_promotes_leading_question_assets_into_empty_material(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "material_chart.png"
            image = Image.new("RGB", (320, 180), "white")
            draw = ImageDraw.Draw(image)
            draw.rectangle((20, 20, 300, 150), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                data_sections=[
                    DataAnalysisSection(
                        title="资料分析",
                        materials=[
                            MaterialUnit(
                                header="材料（回答566-570题）",
                                intro_lines=[],
                                questions=[
                                    ExamQuestion(
                                        stem_lines=[
                                            _text_line("2018 年中国在线旅游收入约占旅游业总收入的:"),
                                            _image_line(str(image_path)),
                                        ],
                                        option_lines=[
                                            _text_line("A.20%"),
                                            _text_line("B.25%"),
                                            _text_line("C.12%"),
                                            _text_line("D.16%"),
                                        ],
                                        source_number="566",
                                    ),
                                    ExamQuestion(
                                        stem_lines=[_text_line("2017 年中国在线旅游收入同比约增长多少万亿元?")],
                                        option_lines=[
                                            _text_line("A.0.15"),
                                            _text_line("B.0.20"),
                                            _text_line("C.0.25"),
                                            _text_line("D.0.30"),
                                        ],
                                        source_number="567",
                                    ),
                                ],
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(image_path): ExtractedImageRegion(str(image_path), 5, 0, 0, 320, 180),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            material = project.sections[0].material_sets[0]
            question = material.questions[0]

            self.assertEqual(len(material.body_assets), 1)
            self.assertEqual(material.body_assets[0].path, str(image_path))
            self.assertEqual(len(question.stem_assets), 0)

    def test_promotes_misplaced_leading_option_image_back_to_previous_visual_question(self):
        with tempfile.TemporaryDirectory() as tmp:
            image_path = Path(tmp) / "network_choices.png"
            image = Image.new("RGB", (360, 120), "white")
            draw = ImageDraw.Draw(image)
            for idx in range(4):
                left = 18 + idx * 84
                draw.rectangle((left, 24, left + 46, 92), fill="black")
            image.save(image_path)

            exam = ParsedExam(
                reasoning_sections=[
                    ReasoningSection(
                        title="判断推理",
                        questions=[
                            ExamQuestion(
                                stem_lines=[_text_line("根据上述定义,下列图形反映了轮式网络特点的是()")],
                                option_lines=[],
                                source_number="1309",
                            ),
                            ExamQuestion(
                                stem_lines=[_text_line("下列使用了连文这一修辞方法的是()")],
                                option_lines=[
                                    _text_line("A．《史记》中的句子"),
                                    _image_line(str(image_path)),
                                    _text_line("B．《易经》中的句子"),
                                    _text_line("C．《出师表》中的句子"),
                                    _text_line("D．《过秦论》中的句子"),
                                ],
                                source_number="1310",
                            ),
                        ],
                    )
                ]
            )
            image_regions = {
                str(image_path): ExtractedImageRegion(str(image_path), 241, 0, 0, 360, 120),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            previous = project.sections[0].questions[0]
            nxt = project.sections[0].questions[1]

            self.assertEqual([option.letter for option in previous.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in previous.options))
            self.assertFalse(nxt.options[0].image_path)

    def test_extracts_options_from_duplicate_overlapping_stem_assets(self):
        with tempfile.TemporaryDirectory() as tmp:
            first_path = Path(tmp) / "duplicate_a.png"
            second_path = Path(tmp) / "duplicate_b.png"
            image = Image.new("RGB", (180, 260), "white")
            draw = ImageDraw.Draw(image)
            for idx in range(4):
                top = 18 + idx * 58
                draw.rectangle((32, top, 148, top + 32), fill="black")
            image.save(first_path)
            image.save(second_path)

            exam = ParsedExam(
                reasoning_sections=[
                    ReasoningSection(
                        title="判断推理",
                        questions=[
                            ExamQuestion(
                                stem_lines=[
                                    _text_line("从所给的四个选项中,选择最合适的一个填入问号处,使之呈现一定的规律性。()"),
                                    _image_line(str(first_path)),
                                    _image_line(str(second_path)),
                                ],
                                option_lines=[],
                                source_number="1806",
                            )
                        ],
                    )
                ]
            )
            image_regions = {
                str(first_path): ExtractedImageRegion(str(first_path), 332, 0, 0, 180, 260),
                str(second_path): ExtractedImageRegion(str(second_path), 332, 0, 0, 180, 260),
            }

            project = build_project_from_parsed_exam(
                exam,
                source_pdf_path="sample.pdf",
                image_regions=image_regions,
            )
            question = project.sections[0].questions[0]

            self.assertEqual([option.letter for option in question.options], ["A", "B", "C", "D"])
            self.assertTrue(all(option.image_path for option in question.options))


if __name__ == "__main__":
    unittest.main()
