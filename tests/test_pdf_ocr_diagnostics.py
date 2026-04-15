import unittest

from core.ai_repair import AIRepairService
from core.pdf_ocr_diagnostics import auto_repair_ocr_project, diagnose_project_ocr_risks
from domain.models import AssetRef, ExamProject, OptionNode, QuestionNode, Section


class PDFOCRDiagnosticsTest(unittest.TestCase):
    def test_diagnose_project_ocr_risks_detects_noise_and_image_only_questions(self):
        project = ExamProject(
            title="ocr_diag",
            sections=[
                Section(
                    kind="unknown",
                    title="待确认",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="A B C D E F G H",
                            options=[
                                OptionNode("A", "甲"),
                                OptionNode("B", "乙"),
                                OptionNode("C", "丙"),
                                OptionNode("D", "丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="2",
                            stem="",
                            stem_assets=[AssetRef(kind="image", path="missing.png")],
                            options=[
                                OptionNode("A", ""),
                                OptionNode("B", ""),
                                OptionNode("C", ""),
                                OptionNode("D", ""),
                            ],
                        ),
                    ],
                )
            ],
        )

        report = diagnose_project_ocr_risks(project)

        self.assertEqual(report.question_count, 2)
        self.assertEqual(report.suspicious_text_questions, 1)
        self.assertGreaterEqual(report.fragmented_questions, 1)
        self.assertEqual(report.image_only_questions, 1)
        self.assertTrue(report.likely_scanned_pdf)
        self.assertTrue(any(issue.code == "ocr_noise_text" for issue in report.issues))

    def test_auto_repair_ocr_project_can_apply_safe_subject_correction(self):
        project = ExamProject(
            title="ocr_repair",
            sections=[
                Section(
                    kind="unknown",
                    title="待确认",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="下列关于宪法的说法正确的是",
                            options=[
                                OptionNode("A", "甲"),
                                OptionNode("B", "乙"),
                                OptionNode("C", "丙"),
                                OptionNode("D", "丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="2",
                            stem="下列关于行政处罚法的说法正确的是",
                            options=[
                                OptionNode("A", "甲"),
                                OptionNode("B", "乙"),
                                OptionNode("C", "丙"),
                                OptionNode("D", "丁"),
                            ],
                        ),
                    ],
                )
            ],
        )

        summary = auto_repair_ocr_project(
            project,
            service=AIRepairService(mode="balanced"),
            only_flagged=True,
            limit=6,
        )

        self.assertGreaterEqual(summary.section_subject_changes, 1)
        self.assertEqual(project.sections[0].kind, "common_sense")
        self.assertLessEqual(
            summary.diagnostics_after.flagged_questions,
            summary.diagnostics_before.flagged_questions,
        )


if __name__ == "__main__":
    unittest.main()
