import unittest

from core.project_quality import annotate_project_quality, is_flagged_question
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section


class ProjectQualityTest(unittest.TestCase):
    def test_annotate_project_quality_flags_unknown_subject_and_blank_option(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="unknown",
                    title="题目列表",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="下列关于宪法的说法正确的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text=""),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        )
                    ],
                )
            ],
        )

        summary = annotate_project_quality(project)
        question = project.sections[0].questions[0]

        self.assertEqual(summary.question_count, 1)
        self.assertEqual(summary.flagged_questions, 1)
        self.assertTrue(is_flagged_question(question))
        self.assertIn("unknown_subject", {issue.code for issue in question.review_issues})
        self.assertIn("blank_option", {issue.code for issue in question.review_issues})
        self.assertEqual(question.suggested_subject, "common_sense")

    def test_annotate_project_quality_flags_number_gap_and_duplicate_number(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="根据上述定义，下列符合定义的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="74",
                            stem="根据上述定义，下列不符合定义的是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="74",
                            stem="如果甲成立，那么乙成立。由此可以推出",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                    ],
                )
            ],
        )

        annotate_project_quality(project)
        second = project.sections[0].questions[1]
        third = project.sections[0].questions[2]

        self.assertIn("number_gap", {issue.code for issue in second.review_issues})
        self.assertIn("duplicate_number", {issue.code for issue in second.review_issues})
        self.assertIn("duplicate_number", {issue.code for issue in third.review_issues})

    def test_annotate_project_quality_flags_empty_material(self):
        project = ExamProject(
            title="资料分析",
            sections=[
                Section(
                    kind="data",
                    title="资料分析",
                    material_sets=[
                        MaterialSet(
                            material_id="m1",
                            header="材料一",
                            body="",
                            questions=[
                                QuestionNode(
                                    source_number="101",
                                    stem="根据所给资料，下列说法正确的是",
                                    options=[
                                        OptionNode(letter="A", text="甲"),
                                        OptionNode(letter="B", text="乙"),
                                        OptionNode(letter="C", text="丙"),
                                        OptionNode(letter="D", text="丁"),
                                    ],
                                )
                            ],
                        )
                    ],
                )
            ],
        )

        annotate_project_quality(project)
        question = project.sections[0].material_sets[0].questions[0]
        self.assertIn("material_empty", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_flags_embedded_data_intro(self):
        project = ExamProject(
            title="资料分析",
            sections=[
                Section(
                    kind="data",
                    title="资料分析",
                    material_sets=[
                        MaterialSet(
                            material_id="m1",
                            header="材料一",
                            body="",
                            questions=[
                                QuestionNode(
                                    source_number="101",
                                    stem="2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。根据上述资料，下列说法正确的是",
                                    options=[
                                        OptionNode(letter="A", text="甲"),
                                        OptionNode(letter="B", text="乙"),
                                        OptionNode(letter="C", text="丙"),
                                        OptionNode(letter="D", text="丁"),
                                    ],
                                )
                            ],
                        )
                    ],
                )
            ],
        )

        annotate_project_quality(project)
        question = project.sections[0].material_sets[0].questions[0]
        issue_codes = {issue.code for issue in question.review_issues}
        self.assertIn("material_intro_embedded_in_stem", issue_codes)

    def test_annotate_project_quality_flags_data_assets_bound_to_first_question(self):
        project = ExamProject(
            title="资料分析",
            sections=[
                Section(
                    kind="data",
                    title="资料分析",
                    material_sets=[
                        MaterialSet(
                            material_id="m1",
                            header="材料一",
                            body="2024年全市工业增加值继续增长。",
                            questions=[
                                QuestionNode(
                                    source_number="101",
                                    stem="根据上述资料，下列说法正确的是",
                                    stem_assets=[
                                        AssetRef(kind="image", path="chart.png", source_page=5),
                                        AssetRef(kind="image", path="table.png", source_page=5),
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
                    ],
                )
            ],
        )

        annotate_project_quality(project)
        question = project.sections[0].material_sets[0].questions[0]
        issue_codes = {issue.code for issue in question.review_issues}

        self.assertIn("material_asset_binding", issue_codes)

    def test_annotate_project_quality_flags_graphic_assets_bound_to_stem(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="问号处应填入",
                            stem_assets=[
                                AssetRef(kind="image", path="a.png"),
                                AssetRef(kind="image", path="b.png"),
                                AssetRef(kind="image", path="c.png"),
                                AssetRef(kind="image", path="d.png"),
                            ],
                            options=[
                                OptionNode(letter="A", text="A"),
                                OptionNode(letter="B", text="B"),
                                OptionNode(letter="C", text="C"),
                                OptionNode(letter="D", text="D"),
                            ],
                        )
                    ],
                )
            ],
        )

        annotate_project_quality(project)
        question = project.sections[0].questions[0]

        self.assertEqual(question.inferred_subtype, "图形推理")
        self.assertIn("graphic_asset_binding", {issue.code for issue in question.review_issues})


if __name__ == "__main__":
    unittest.main()
