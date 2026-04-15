import json
import os
import tempfile
import unittest
from pathlib import Path

from core.repair_log import append_question_repair_log, capture_question_state
from domain.models import ExamProject, MaterialSet, OptionNode, PaperSource, QuestionNode, Section
from exporters.manifest_json import export_project_manifest
from scripts.build_gui_repair_log_dataset import build_gui_repair_log_dataset


class RepairLogTest(unittest.TestCase):
    def _build_project(self) -> ExamProject:
        return ExamProject(
            title="repair_log_demo",
            source=PaperSource(pdf_path="sample.pdf"),
            sections=[
                Section(
                    kind="data",
                    title="资料分析",
                    material_sets=[
                        MaterialSet(
                            material_id="m1",
                            header="材料一",
                            body="2024年数据",
                            questions=[
                                QuestionNode(
                                    source_number="101",
                                    stem="根据上述资料，下列说法正确的是",
                                    options=[
                                        OptionNode("A", "甲"),
                                        OptionNode("B", "乙"),
                                        OptionNode("C", "丙"),
                                        OptionNode("D", "丁"),
                                    ],
                                )
                            ],
                        )
                    ],
                )
            ],
        )

    def test_append_question_repair_log_assigns_tracking_ids(self):
        project = self._build_project()
        section = project.sections[0]
        material = section.material_sets[0]
        question = material.questions[0]
        before_state = capture_question_state(section=section, material=material, question=question)

        question.stem = "根据上述资料，2024年同比增长5.2%，下列说法正确的是"
        entry = append_question_repair_log(
            project,
            source="gui_manual",
            action="update_question_stem",
            section=section,
            material=material,
            question=question,
            before_state=before_state,
            metadata={"field": "stem"},
        )

        self.assertIsNotNone(entry)
        self.assertTrue(project.repair_session_id)
        self.assertTrue(question.question_id)
        self.assertEqual(len(project.repair_log), 1)
        self.assertEqual(project.repair_log[0].action, "update_question_stem")
        self.assertEqual(project.repair_log[0].before_state["question_no"], "101")

    def test_build_gui_repair_log_dataset_reads_manifest_logs(self):
        project = self._build_project()
        section = project.sections[0]
        material = section.material_sets[0]
        question = material.questions[0]
        before_state = capture_question_state(section=section, material=material, question=question)
        question.stem = "根据上述资料，2024年同比增长5.2%，下列说法正确的是"
        append_question_repair_log(
            project,
            source="gui_manual",
            action="update_question_stem",
            section=section,
            material=material,
            question=question,
            before_state=before_state,
            metadata={"field": "stem"},
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            manifest_path = os.path.join(temp_dir, "project.json")
            output_path = Path(temp_dir) / "gui_repair_logs.jsonl"
            export_project_manifest(project, manifest_path)

            summary = build_gui_repair_log_dataset([manifest_path], output_path)

            self.assertEqual(summary["row_count"], 1)
            self.assertTrue(output_path.exists())
            rows = [
                json.loads(line)
                for line in output_path.read_text(encoding="utf-8").splitlines()
                if line.strip()
            ]
            self.assertEqual(rows[0]["action"], "update_question_stem")
            self.assertEqual(rows[0]["source"], "gui_manual")
            self.assertEqual(rows[0]["source_pdf"], "sample.pdf")


if __name__ == "__main__":
    unittest.main()
