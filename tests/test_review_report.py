import json
import os
import tempfile
import unittest

from core.project_quality import annotate_project_quality, question_max_severity
from domain.models import ExamProject, MaterialSet, OptionNode, QuestionNode, Section
from exporters.review_report import build_quality_report_payload, export_quality_report


class ReviewReportTest(unittest.TestCase):
    def test_build_quality_report_payload_includes_flagged_items(self):
        project = ExamProject(
            title="示例工程",
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
                ),
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
                                    stem="根据资料，下列说法正确的是",
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
                ),
            ],
        )
        annotate_project_quality(project)

        payload = build_quality_report_payload(project)

        self.assertEqual(payload["title"], "示例工程")
        self.assertEqual(payload["question_count"], 2)
        self.assertEqual(payload["flagged_question_count"], 2)
        self.assertEqual(payload["items"][0]["source_number"], "1")
        self.assertIn("issues", payload["items"][0])
        self.assertIsNotNone(payload["items"][0]["suggested_subject"])

    def test_export_quality_report_writes_json(self):
        project = ExamProject(
            title="示例工程",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="66",
                            stem="",
                            options=[OptionNode(letter="A", text="")],
                        )
                    ],
                )
            ],
        )
        annotate_project_quality(project)

        with tempfile.TemporaryDirectory() as temp_dir:
            path = os.path.join(temp_dir, "report.json")
            export_quality_report(project, path)
            with open(path, "r", encoding="utf-8") as file_obj:
                payload = json.load(file_obj)

        self.assertEqual(payload["flagged_question_count"], 1)
        self.assertEqual(payload["items"][0]["severity"], "error")
        self.assertEqual(question_max_severity(project.sections[0].questions[0]), "error")


if __name__ == "__main__":
    unittest.main()
