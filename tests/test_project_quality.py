import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

import core.learned_subject_model as learned_subject_model
from core.project_quality import annotate_project_quality, is_flagged_question
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, PaperSource, QuestionNode, Section


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

    def test_annotate_project_quality_downgrades_single_blank_option_when_three_slots_are_recovered(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="176",
                            stem="",
                            stem_assets=[AssetRef(kind="stem_image", path="stem.png", source_page=1)],
                            options=[
                                OptionNode(letter="A", text="", image_path="a.png"),
                                OptionNode(letter="B", text=""),
                                OptionNode(letter="C", text="", image_path="c.png"),
                                OptionNode(letter="D", text="", image_path="d.png"),
                            ],
                        )
                    ],
                )
            ],
        )

        annotate_project_quality(project)
        question = project.sections[0].questions[0]
        self.assertIn("source_text_missing", {issue.code for issue in question.review_issues})
        self.assertNotIn("blank_option", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_marks_source_missing_visual_question(self):
        project = ExamProject(
            title="示例",
            source=PaperSource(pdf_path="missing-visual.pdf"),
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="474",
                            stem="问以下哪个坐标图能准确表示甲、乙生产线产量之差与总生产时间之间的关系（）",
                            options=[],
                            page_numbers=[53],
                        )
                    ],
                )
            ],
        )

        with patch("core.project_quality._pages_have_visual_candidates", return_value=False):
            summary = annotate_project_quality(project)

        question = project.sections[0].questions[0]
        self.assertEqual(summary.source_defect_questions, 1)
        self.assertIn("source_visual_missing", {issue.code for issue in question.review_issues})
        self.assertNotIn("option_count", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_marks_source_missing_text_placeholder(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="970",
                            stem="缺失",
                            options=[],
                        )
                    ],
                )
            ],
        )

        summary = annotate_project_quality(project)
        question = project.sections[0].questions[0]

        self.assertEqual(summary.source_defect_questions, 1)
        self.assertIn("source_text_missing", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_marks_missing_option_placeholders_as_source_missing(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="56",
                            stem="某药物剂量判断题",
                            options=[
                                OptionNode(letter="A", text="缺失"),
                                OptionNode(letter="B", text="40毫克"),
                                OptionNode(letter="C", text="缺失"),
                                OptionNode(letter="D", text="缺失"),
                            ],
                        )
                    ],
                )
            ],
        )

        summary = annotate_project_quality(project)
        question = project.sections[0].questions[0]

        self.assertEqual(summary.source_defect_questions, 1)
        self.assertIn("source_text_missing", {issue.code for issue in question.review_issues})
        self.assertNotIn("duplicate_option_text", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_marks_empty_placeholder_between_consecutive_questions_as_source_missing(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="1332",
                            stem="上一题图形题",
                            stem_assets=[AssetRef(kind="stem_image", path="a.png", source_page=244)],
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="1333",
                            stem="",
                            options=[],
                        ),
                        QuestionNode(
                            source_number="1334",
                            stem="下一题图形题",
                            stem_assets=[AssetRef(kind="stem_image", path="b.png", source_page=245)],
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

        summary = annotate_project_quality(project)
        question = project.sections[0].questions[1]

        self.assertEqual(summary.source_defect_questions, 1)
        self.assertIn("source_text_missing", {issue.code for issue in question.review_issues})
        self.assertNotIn("missing_stem", {issue.code for issue in question.review_issues})
        self.assertNotIn("option_count", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_marks_partial_inline_options_before_next_question_as_source_missing(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="verbal",
                    title="言语理解",
                    questions=[
                        QuestionNode(
                            source_number="1911",
                            stem=(
                                "这段文字中，盘口壶的例子意在说明()。"
                                "A.类型学是常用的考古研究方法"
                                "B.新技术手段如何帮助考古分析"
                            ),
                            options=[],
                        ),
                        QuestionNode(
                            source_number="1912",
                            stem="随着中国文化在全球的传播范围愈来愈广，下列说法与文意相符的是()。",
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

        summary = annotate_project_quality(project)
        question = project.sections[0].questions[0]

        self.assertEqual(summary.source_defect_questions, 1)
        self.assertIn("source_text_missing", {issue.code for issue in question.review_issues})
        self.assertNotIn("option_count", {issue.code for issue in question.review_issues})

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

    def test_annotate_project_quality_ignores_source_gap_number_jump(self):
        project = ExamProject(
            title="示例",
            source=PaperSource(pdf_path="reasoning-gap.pdf"),
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="565",
                            stem="上一题",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                            page_numbers=[108],
                        ),
                        QuestionNode(
                            source_number="567",
                            stem="下一题",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                            page_numbers=[108],
                        ),
                    ],
                )
            ],
        )

        with patch("core.project_quality._page_text_map", return_value={108: "565. 上一题\n567. 下一题"}):
            annotate_project_quality(project)

        question = project.sections[0].questions[1]
        self.assertNotIn("number_gap", {issue.code for issue in question.review_issues})

    def test_annotate_project_quality_ignores_politics_chapter_number_reset(self):
        project = ExamProject(
            title="政治理论题本",
            sections=[
                Section(
                    kind="politics",
                    title="政治理论",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="（2025·联考）关于新时代中国特色社会主义思想，下列说法正确的是：",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="2",
                            stem="（2025·联考）关于高质量发展的理解，下列说法错误的是：",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="1",
                            stem="（2025·上海）下列关于马克思主义哲学的说法不正确的是：",
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
        reset_question = project.sections[0].questions[2]

        self.assertNotIn("number_order", {issue.code for issue in reset_question.review_issues})

    def test_annotate_project_quality_ignores_politics_reset_to_small_number(self):
        project = ExamProject(
            title="政治理论题本",
            sections=[
                Section(
                    kind="politics",
                    title="政治理论",
                    questions=[
                        QuestionNode(
                            source_number="29",
                            stem="（2025·北京）上一章节的最后一道题。",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="10",
                            stem="（2025·江苏）新章节重新编号后的题目。",
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
        reset_question = project.sections[0].questions[1]

        self.assertNotIn("number_order", {issue.code for issue in reset_question.review_issues})

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
