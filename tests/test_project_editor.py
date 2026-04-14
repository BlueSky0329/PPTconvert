import unittest

from core.pdf_exam_models import DataAnalysisSection, ExamQuestion, MaterialUnit, ParsedExam, QuantSection, RichLine
from core.pdf_exam_models import CommonSenseSection, ReasoningSection
from domain.models import AssetRef, MaterialSet, OptionNode, PageRegion, QuestionNode
from domain.project_editor import (
    apply_all_safe_subject_suggestions,
    apply_section_subject_suggestion,
    clear_option_image,
    insert_question_before,
    insert_option_after,
    insert_material_after,
    merge_adjacent_materials,
    move_option,
    move_stem_assets_to_material,
    move_data_question,
    remove_question,
    remove_option,
    replace_option_image,
    reclassify_objective_section,
    rename_material,
    renumber_question,
    section_subject_suggestion,
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

    def test_move_stem_assets_to_material(self):
        material = MaterialSet(
            material_id="m1",
            header="材料一",
            body="2024年相关数据如下。",
            questions=[
                QuestionNode(
                    source_number="101",
                    stem="根据上述资料，下列说法正确的是",
                    stem_assets=[
                        AssetRef(
                            kind="image",
                            path="chart.png",
                            source_page=2,
                            page_region=PageRegion(page_number=2, x0=10, y0=20, x1=110, y1=160),
                        )
                    ],
                    options=[
                        OptionNode(letter="A", text="甲"),
                        OptionNode(letter="B", text="乙"),
                        OptionNode(letter="C", text="丙"),
                        OptionNode(letter="D", text="丁"),
                    ],
                )
            ],
        )
        question = material.questions[0]

        moved = move_stem_assets_to_material(material, question)

        self.assertEqual(moved, 1)
        self.assertEqual(question.stem_assets, [])
        self.assertEqual(len(material.body_assets), 1)
        self.assertEqual(material.body_assets[0].path, "chart.png")
        self.assertEqual(len(material.body_regions), 1)

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

    def test_insert_question_before(self):
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
        existing = project.sections[0].questions[0]
        inserted = insert_question_before(
            project,
            existing,
            QuestionNode(
                source_number="65",
                stem="新插入题干",
                options=[
                    OptionNode(letter="A", text="甲"),
                    OptionNode(letter="B", text="乙"),
                    OptionNode(letter="C", text="丙"),
                    OptionNode(letter="D", text="丁"),
                ],
            ),
        )

        self.assertTrue(inserted)
        self.assertEqual([q.source_number for q in project.sections[0].questions], ["65", "66"])

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

    def test_reclassify_section_merges_adjacent_equivalent_sections_and_refreshes_subjects(self):
        exam = ParsedExam(
            common_sense_sections=[
                CommonSenseSection(
                    title="二. 常识判断",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("1题题干")],
                            option_lines=[_text_line("A. 甲"), _text_line("B. 乙")],
                            source_number="1",
                        )
                    ]
                )
            ],
            reasoning_sections=[
                ReasoningSection(
                    title="判断推理",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("76题题干")],
                            option_lines=[_text_line("A. 丙"), _text_line("B. 丁")],
                            source_number="76",
                        )
                    ]
                )
            ],
        )
        project = build_project_from_parsed_exam(exam)

        changed = reclassify_objective_section(project.sections[0], "reasoning", project=project)

        self.assertTrue(changed)
        self.assertEqual(len(project.sections), 1)
        self.assertEqual(project.sections[0].kind, "reasoning")
        self.assertEqual(project.sections[0].title, "判断推理")
        self.assertEqual([question.source_number for question in project.sections[0].questions], ["1", "76"])
        self.assertEqual(project.selected_subjects, ["reasoning"])

    def test_apply_section_subject_suggestion_updates_unknown_section(self):
        exam = ParsedExam(
            common_sense_sections=[
                CommonSenseSection(
                    title="题目列表",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("下列关于宪法的说法正确的是")],
                            option_lines=[
                                _text_line("A. 甲"),
                                _text_line("B. 乙"),
                                _text_line("C. 丙"),
                                _text_line("D. 丁"),
                            ],
                            source_number="1",
                        )
                    ],
                )
            ]
        )
        project = build_project_from_parsed_exam(exam)
        project.sections[0].kind = "unknown"
        project.sections[0].title = "题目列表"
        project.sections[0].questions[0].suggested_subject = "common_sense"

        target, reason = section_subject_suggestion(project.sections[0])
        changed = apply_section_subject_suggestion(project.sections[0], project=project)

        self.assertEqual(target, "common_sense")
        self.assertIn("一致建议改为", reason)
        self.assertTrue(changed)
        self.assertEqual(project.sections[0].kind, "common_sense")

    def test_apply_all_safe_subject_suggestions_skips_conflicting_section(self):
        exam = ParsedExam(
            common_sense_sections=[
                CommonSenseSection(
                    title="题目列表",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("下列关于宪法的说法正确的是")],
                            option_lines=[
                                _text_line("A. 甲"),
                                _text_line("B. 乙"),
                                _text_line("C. 丙"),
                                _text_line("D. 丁"),
                            ],
                            source_number="1",
                        )
                    ],
                )
            ],
            reasoning_sections=[
                ReasoningSection(
                    title="题目列表",
                    questions=[
                        ExamQuestion(
                            stem_lines=[_text_line("根据上述定义，下列符合定义的是")],
                            option_lines=[
                                _text_line("A. 甲"),
                                _text_line("B. 乙"),
                                _text_line("C. 丙"),
                                _text_line("D. 丁"),
                            ],
                            source_number="71",
                        ),
                        ExamQuestion(
                            stem_lines=[_text_line("如果甲成立，那么乙成立。由此可以推出")],
                            option_lines=[
                                _text_line("A. 甲"),
                                _text_line("B. 乙"),
                                _text_line("C. 丙"),
                                _text_line("D. 丁"),
                            ],
                            source_number="72",
                        ),
                    ],
                )
            ],
        )
        project = build_project_from_parsed_exam(exam)
        project.sections[0].kind = "unknown"
        project.sections[0].title = "题目列表"
        project.sections[0].questions[0].suggested_subject = "common_sense"
        project.sections[1].questions[0].suggested_subject = "reasoning"
        project.sections[1].questions[1].suggested_subject = "common_sense"

        applied = apply_all_safe_subject_suggestions(project)

        self.assertEqual(applied, 1)
        self.assertEqual(project.sections[0].kind, "common_sense")
        self.assertEqual(project.sections[1].kind, "reasoning")


if __name__ == "__main__":
    unittest.main()
