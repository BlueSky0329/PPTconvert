import unittest

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
)
from domain.selectors import parse_question_ranges, select_project
from exporters.pptx_slides import project_to_ppt_questions
from ingest.pdf.layout import PageTextLine
from ingest.pdf.project_builder import build_project_from_parsed_exam


def _text_line(text: str) -> RichLine:
    return RichLine(parts=[(text, None)])


def _image_line(path: str) -> RichLine:
    return RichLine(parts=[("", path)])


class ExamProjectTest(unittest.TestCase):
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


if __name__ == "__main__":
    unittest.main()
