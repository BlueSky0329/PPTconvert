import unittest

from core.pdf_exam_models import DataAnalysisSection, ExamQuestion, MaterialUnit, ParsedExam, QuantSection, RichLine
from domain.models import AssetRef, PageRegion
from domain.project_editor import (
    insert_material_after,
    merge_adjacent_materials,
    move_data_question,
    remove_question,
    rename_material,
    renumber_question,
)
from ingest.pdf.project_builder import build_project_from_parsed_exam


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
