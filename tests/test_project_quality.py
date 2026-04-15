import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

import core.learned_subject_model as learned_subject_model
from core.project_quality import annotate_project_quality, is_flagged_question
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section


class DummyDecisionSubjectModel:
    class _DecisionMatrix(list):
        ndim = 2

    def decision_function(self, rows):
        row = rows[0]
        text = row["text"]
        if "下列说法正确的是" in text:
            return self._DecisionMatrix([[2.4, -0.6, 0.1]])
        return self._DecisionMatrix([[2.0, -0.8, 0.0]])


class ProjectQualityTest(unittest.TestCase):
    def tearDown(self):
        learned_subject_model._CACHE_KEY = None
        learned_subject_model._CACHE_BUNDLE = None

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

    def test_annotate_project_quality_uses_enabled_trained_subject_model(self):
        baseline_project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="unknown",
                    title="题目列表",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="下列说法正确的是",
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
        annotate_project_quality(baseline_project)
        baseline_question = baseline_project.sections[0].questions[0]

        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="unknown",
                    title="题目列表",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="下列说法正确的是",
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
        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummyDecisionSubjectModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "1",
                    "PPTCONVERT_TRUST_SUBJECT_MODEL": "1",
                    "PPTCONVERT_SUBJECT_MODEL": str(model_path),
                },
                clear=False,
            ):
                annotate_project_quality(project)

        question = project.sections[0].questions[0]
        self.assertFalse(any("学习模型" in signal for signal in baseline_question.inferred_signals))
        self.assertTrue(any("学习模型:常识判断" in signal for signal in question.inferred_signals))

    def test_annotate_project_quality_uses_ready_default_subject_model_without_enable_flag(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="unknown",
                    title="题目列表",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="下列说法正确的是",
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

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "subject_classifier.pkl"
            bundle = {
                "labels": ["common_sense", "reasoning", "verbal"],
                "model": DummyDecisionSubjectModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with (
                patch.object(learned_subject_model, "_DEFAULT_MODEL_PATH", model_path),
                patch.object(learned_subject_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_PICKLED_SUBJECT_MODEL": "",
                        "PPTCONVERT_SUBJECT_MODEL": "",
                        "PPTCONVERT_TRUST_SUBJECT_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                annotate_project_quality(project)

        question = project.sections[0].questions[0]
        self.assertEqual(question.suggested_subject, "common_sense")
        self.assertTrue(any("学习模型:常识判断" in signal for signal in question.inferred_signals))


if __name__ == "__main__":
    unittest.main()
