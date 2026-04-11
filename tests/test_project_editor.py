import unittest

from core.pdf_exam_models import DataAnalysisSection, ExamQuestion, MaterialUnit, ParsedExam, QuantSection, RichLine
from domain.models import AssetRef, PageRegion
from domain.project_editor import (
    clear_option_image,
    insert_option_after,
    insert_material_after,
    merge_adjacent_materials,
    move_option,
    move_data_question,
    remove_question,
    remove_option,
    replace_option_image,
    rename_material,
    renumber_question,
    set_question_option_layout,
    update_option_text,
    update_question_stem,
)
from ingest.pdf.project_builder import build_project_from_parsed_exam
from core.pdf_exam_extract import ExtractedImageRegion


def _text_line(text: str) -> RichLine:
    return RichLine(parts=[(text, None)])


class ProjectEditorTest(unittest.TestCase):
    def test_remove_question_cleans_up_empty_quant_section(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="数量关系",
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
        project = build_project_from_parsed_exam(exam)

        removed = remove_question(project, project.sections[0].questions[0])

        self.assertTrue(removed)
        self.assertEqual(project.sections, [])

    def test_move_data_question_to_next_material(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("材料一正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="111",
                                )
                            ],
                        ),
                        MaterialUnit(
                            header="材料二",
                            intro_lines=[_text_line("材料二正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("116题题干")],
                                    option_lines=[_text_line("A. 丙"), _text_line("B. 丁")],
                                    source_number="116",
                                )
                            ],
                        ),
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam)
        question = project.sections[0].material_sets[0].questions[0]

        moved = move_data_question(project, question, 1)

        self.assertTrue(moved)
        self.assertEqual(len(project.sections[0].material_sets), 1)
        self.assertEqual(
            [q.source_number for q in project.sections[0].material_sets[0].questions],
            ["111", "116"],
        )

    def test_rename_and_renumber(self):
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
        project = build_project_from_parsed_exam(exam)
        material = project.sections[0].material_sets[0]
        question = material.questions[0]

        rename_material(material, "材料甲")
        renumber_question(question, "211")

        self.assertEqual(material.header, "材料甲")
        self.assertEqual(question.source_number, "211")

    def test_update_question_stem_and_option_layout(self):
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
        project = build_project_from_parsed_exam(exam)
        question = project.sections[0].questions[0]

        update_question_stem(question, "新的题干")
        set_question_option_layout(question, "one_row")

        self.assertEqual(question.stem, "新的题干")
        self.assertEqual(question.option_layout, "one_row")

        set_question_option_layout(question, "unknown")
        self.assertIsNone(question.option_layout)

    def test_update_option_text_and_image(self):
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
        project = build_project_from_parsed_exam(exam)
        question = project.sections[0].questions[0]

        self.assertTrue(update_option_text(question, "A", "10"))
        self.assertTrue(replace_option_image(question, "B", "sample.png"))
        self.assertEqual(question.options[0].text, "10")
        self.assertEqual(question.options[1].image_path, "sample.png")

        self.assertTrue(clear_option_image(question, "B"))
        self.assertIsNone(question.options[1].image_path)
        self.assertFalse(update_option_text(question, "Z", "x"))

    def test_move_insert_and_remove_option(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[
                                _text_line("A. 1"),
                                _text_line("B. 2"),
                                _text_line("C. 3"),
                            ],
                            source_number="66",
                        )
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam)
        question = project.sections[0].questions[0]
        question.answer = "AC"

        self.assertTrue(move_option(question, "C", -1))
        self.assertEqual([(o.letter, o.text) for o in question.options], [("A", "1"), ("B", "3"), ("C", "2")])
        self.assertEqual(question.answer, "AB")

        self.assertTrue(insert_option_after(question, "B"))
        self.assertEqual([o.letter for o in question.options], ["A", "B", "C", "D"])
        self.assertEqual(question.options[2].text, "")
        self.assertEqual(question.answer, "AB")

        self.assertTrue(remove_option(question, "B"))
        self.assertEqual([o.letter for o in question.options], ["A", "B", "C"])
        self.assertEqual([o.text for o in question.options], ["1", "", "2"])
        self.assertEqual(question.answer, "A")

    def test_remove_answered_option_drops_that_answer_letter(self):
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[
                                _text_line("A. 1"),
                                _text_line("B. 2"),
                                _text_line("C. 3"),
                            ],
                            source_number="66",
                        )
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam)
        question = project.sections[0].questions[0]
        question.answer = "BC"

        self.assertTrue(remove_option(question, "B"))
        self.assertEqual([o.letter for o in question.options], ["A", "B"])
        self.assertEqual([o.text for o in question.options], ["1", "3"])
        self.assertEqual(question.answer, "B")

    def test_build_project_preserves_option_image_region(self):
        image_path = "option_a.png"
        exam = ParsedExam(
            quant_sections=[
                QuantSection(
                    title="四. 数量关系",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("66题题干")],
                            option_lines=[
                                RichLine(parts=[("A. ", None), ("", image_path)]),
                                _text_line("B. 2"),
                            ],
                            source_number="66",
                        )
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(
            exam,
            image_regions={
                image_path: ExtractedImageRegion(
                    path=image_path,
                    page_number=3,
                    x0=10.0,
                    y0=20.0,
                    x1=40.0,
                    y1=60.0,
                )
            },
        )

        option = project.sections[0].questions[0].options[0]
        self.assertEqual(option.image_path, image_path)
        self.assertEqual(option.source_page, 3)
        self.assertIsNotNone(option.page_region)

    def test_insert_material_after(self):
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
        project = build_project_from_parsed_exam(exam)

        inserted = insert_material_after(project, project.sections[0].material_sets[0], header="材料二")

        self.assertTrue(inserted)
        self.assertEqual([m.header for m in project.sections[0].material_sets], ["材料一", "材料二"])

    def test_merge_adjacent_materials_with_next(self):
        exam = ParsedExam(
            data_sections=[
                DataAnalysisSection(
                    title="资料分析",
                    materials=[
                        MaterialUnit(
                            header="材料一",
                            intro_lines=[_text_line("材料一正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("111题题干")],
                                    option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                                    source_number="111",
                                )
                            ],
                        ),
                        MaterialUnit(
                            header="材料二",
                            intro_lines=[_text_line("材料二正文")],
                            questions=[
                                ExamQuestion(
                                    stem_lines=[_text_line("116题题干")],
                                    option_lines=[_text_line("A. 丙"), _text_line("B. 丁")],
                                    source_number="116",
                                )
                            ],
                        ),
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam)
        first = project.sections[0].material_sets[0]
        second = project.sections[0].material_sets[1]
        first.body_assets = [AssetRef(kind="material_inline_image", path="one.png")]
        first.body_regions = [PageRegion(page_number=1, x0=1, y0=1, x1=2, y1=2)]
        second.body_assets = [AssetRef(kind="material_inline_image", path="two.png")]
        second.body_regions = [PageRegion(page_number=2, x0=3, y0=3, x1=4, y1=4)]

        merged = merge_adjacent_materials(project, first, 1)

        self.assertTrue(merged)
        self.assertEqual(len(project.sections[0].material_sets), 1)
        merged_material = project.sections[0].material_sets[0]
        self.assertEqual([q.source_number for q in merged_material.questions], ["111", "116"])
        self.assertEqual(merged_material.body_lines, ["材料一正文", "材料二正文"])
        self.assertEqual(merged_material.body, "材料一正文\n材料二正文")
        self.assertEqual([asset.path for asset in merged_material.body_assets], ["one.png", "two.png"])
        self.assertEqual([region.page_number for region in merged_material.body_regions], [1, 2])

    def test_insert_material_after_generates_unique_ids(self):
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
        project = build_project_from_parsed_exam(exam)
        first = project.sections[0].material_sets[0]

        inserted_once = insert_material_after(project, first, header="材料二")
        inserted_twice = insert_material_after(project, first, header="材料三")

        self.assertTrue(inserted_once)
        self.assertTrue(inserted_twice)
        material_ids = [material.material_id for material in project.sections[0].material_sets]
        self.assertEqual(len(material_ids), len(set(material_ids)))


if __name__ == "__main__":
    unittest.main()
