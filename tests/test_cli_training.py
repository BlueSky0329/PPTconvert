import tempfile
import unittest
from argparse import Namespace
from pathlib import Path
from types import SimpleNamespace
from unittest.mock import patch

from domain.models import ExamProject, OptionNode, QuestionNode, Section
from main import run_pdf_workflow
from scripts.build_gold_subject_dataset import build_dataset
from scripts.train_local_subject_model import (
    _build_sample_weights,
    _choose_grouped_split,
    _minimum_train_label_support,
)


class CliWorkflowTest(unittest.TestCase):
    def test_run_pdf_workflow_exports_repaired_project(self):
        project = ExamProject(
            title="示例工程",
            sections=[
                Section(
                    kind="verbal",
                    title="言语理解",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="原始题干",
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
        quality = SimpleNamespace(flagged_questions=0, severe_questions=0, total_issue_count=0)

        def fake_ai_repair(_args, repaired_project):
            repaired_project.sections[0].questions[0].stem = "AI修复后的题干"

        def fake_export_project_outputs(exported_project, **_kwargs):
            self.assertEqual(exported_project.sections[0].questions[0].stem, "AI修复后的题干")
            return SimpleNamespace(asset_dir="assets", docx_path="题本.docx", pptx_path=None, manifest_path=None)

        with tempfile.TemporaryDirectory() as tmp:
            pdf_path = Path(tmp) / "sample.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            args = Namespace(
                pdf_input=str(pdf_path),
                docx_output=None,
                ppt_output=None,
                output=None,
                manifest_output=None,
                layout=None,
                font_size=None,
                subject="all",
                question_range=None,
                document_subject="auto",
                template=None,
                ai_repair=True,
                ai_limit=12,
                ai_all_questions=False,
            )
            with patch("workflows.project_flow.build_pdf_project", return_value=(project, "assets")), patch(
                "workflows.project_flow.export_project_outputs",
                side_effect=fake_export_project_outputs,
            ), patch("core.project_quality.annotate_project_quality", return_value=quality), patch(
                "main.maybe_run_ai_repair",
                side_effect=fake_ai_repair,
            ):
                run_pdf_workflow(args)


class TrainingSplitTest(unittest.TestCase):
    def test_choose_grouped_split_preserves_pdf_boundaries(self):
        labels = [
            "verbal",
            "verbal",
            "reasoning",
            "reasoning",
            "verbal",
            "reasoning",
        ]
        groups = [
            "pdf-a",
            "pdf-a",
            "pdf-b",
            "pdf-b",
            "pdf-c",
            "pdf-d",
        ]

        train_idx, test_idx, full_coverage = _choose_grouped_split(labels, groups)

        train_groups = {groups[index] for index in train_idx}
        test_groups = {groups[index] for index in test_idx}
        self.assertTrue(full_coverage)
        self.assertTrue(train_groups.isdisjoint(test_groups))
        self.assertEqual({labels[index] for index in train_idx}, {"verbal", "reasoning"})
        self.assertEqual({labels[index] for index in test_idx}, {"verbal", "reasoning"})

    def test_build_sample_weights_prefers_question_bank_and_backfills_missing_labels(self):
        rows = [
            {"source_form": "single_subject_book", "subject": "reasoning"},
            {"source_form": "set_paper", "subject": "reasoning"},
            {"source_form": "set_paper", "subject": "politics"},
        ]

        weights, summary = _build_sample_weights(rows, preferred_forms={"single_subject_book"})

        self.assertEqual(weights, [3.0, 0.35, 1.5])
        self.assertEqual(summary["preferred"], 1.0)
        self.assertEqual(summary["fallback_same_label"], 1.0)
        self.assertEqual(summary["fallback_missing_label"], 1.0)

    def test_build_dataset_can_filter_forms_and_pass_subject_hint(self):
        project = ExamProject(
            title="示例工程",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="题干",
                            options=[OptionNode("A", "甲"), OptionNode("B", "乙")],
                        )
                    ],
                )
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            single_pdf = root / "single.pdf"
            set_pdf = root / "set.pdf"
            single_pdf.write_bytes(b"%PDF-1.4\n")
            set_pdf.write_bytes(b"%PDF-1.4\n")
            catalog_path = root / "catalog.json"
            output_path = root / "subject_gold.jsonl"
            catalog_path.write_text(
                """
                {
                  "pdfs": [
                    {"path": "single.pdf", "form": "single_subject_book", "subject": "reasoning"},
                    {"path": "set.pdf", "form": "set_paper", "sections": {"reasoning": ["1-1"]}}
                  ]
                }
                """,
                encoding="utf-8",
            )

            with patch("scripts.build_gold_subject_dataset.ROOT", root), patch(
                "scripts.build_gold_subject_dataset.build_exam_project_from_pdf",
                return_value=project,
            ) as build_mock:
                summary = build_dataset(catalog_path, output_path, forms={"single_subject_book"})

        build_mock.assert_called_once_with(
            str(single_pdf),
            mode="all",
            document_subject_hint="reasoning",
        )
        self.assertEqual(summary["pdf_count"], 1)
        self.assertEqual(summary["subject_distribution"]["reasoning"], 1)

    def test_choose_grouped_split_keeps_dominant_label_source_in_training(self):
        labels = (
            ["politics"] * 50
            + ["quant"] * 50
            + ["verbal"] * 50
            + ["politics", "quant", "verbal"] * 5
            + ["politics", "quant", "verbal"] * 5
        )
        groups = (
            ["politics-book"] * 50
            + ["quant-book"] * 50
            + ["verbal-book"] * 50
            + ["set-a"] * 15
            + ["set-b"] * 15
        )

        train_idx, test_idx, full_coverage = _choose_grouped_split(labels, groups)

        train_groups = {groups[index] for index in train_idx}
        test_groups = {groups[index] for index in test_idx}
        self.assertTrue(full_coverage)
        self.assertTrue(train_groups.isdisjoint(test_groups))
        self.assertNotIn("politics-book", test_groups)
        self.assertNotIn("quant-book", test_groups)
        politics_train = sum(1 for index in train_idx if labels[index] == "politics")
        quant_train = sum(1 for index in train_idx if labels[index] == "quant")
        self.assertGreaterEqual(politics_train, _minimum_train_label_support(labels.count("politics")))
        self.assertGreaterEqual(quant_train, _minimum_train_label_support(labels.count("quant")))


if __name__ == "__main__":
    unittest.main()
