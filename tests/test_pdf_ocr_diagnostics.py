import unittest
from unittest.mock import patch

from core.ai_repair import AIRepairService
import core.pdf_ocr_diagnostics as pdf_ocr_diagnostics
from core.pdf_ocr_diagnostics import (
    OCRRecoveryPreview,
    auto_repair_ocr_project,
    diagnose_project_ocr_risks,
)
from core.pdf_ocr_engine import OCRLine
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


class OCRProbeFieldTest(unittest.TestCase):
    def _build_image_only_project(self) -> ExamProject:
        return ExamProject(
            title="ocr_probe",
            sections=[
                Section(
                    kind="unknown",
                    title="待确认",
                    questions=[
                        QuestionNode(
                            source_number=str(idx),
                            stem="",
                            stem_assets=[AssetRef(kind="image", path="missing.png")],
                            options=[
                                OptionNode("A", ""),
                                OptionNode("B", ""),
                                OptionNode("C", ""),
                                OptionNode("D", ""),
                            ],
                        )
                        for idx in range(1, 4)
                    ],
                )
            ],
        )

    def test_default_does_not_run_probe(self):
        project = self._build_image_only_project()
        with (
            patch("core.pdf_ocr_engine.is_ocr_dependency_available", return_value=True)
            as dependency_mock,
            patch(
                "core.pdf_ocr_engine.is_ocr_available",
                side_effect=AssertionError("default diagnosis should not load OCR"),
            ),
            patch("core.pdf_ocr_diagnostics._probe_ocr_recovery") as probe_mock,
        ):
            report = diagnose_project_ocr_risks(project)
        dependency_mock.assert_called_once()
        probe_mock.assert_not_called()
        self.assertTrue(report.likely_scanned_pdf)
        self.assertTrue(report.ocr_available)
        self.assertEqual(report.ocr_recoverable_samples, [])

    def test_probe_skipped_when_engine_unavailable(self):
        project = self._build_image_only_project()
        with (
            patch("core.pdf_ocr_engine.is_ocr_available", return_value=False),
            patch("core.pdf_ocr_diagnostics._probe_ocr_recovery") as probe_mock,
        ):
            report = diagnose_project_ocr_risks(project, include_ocr_probe=True)
        probe_mock.assert_not_called()
        self.assertFalse(report.ocr_available)
        self.assertTrue(
            any("OCR 引擎" in suggestion for suggestion in report.suggestions),
            report.suggestions,
        )

    def test_probe_runs_when_requested_and_engine_available(self):
        project = self._build_image_only_project()
        fake_previews = [
            OCRRecoveryPreview(
                page_number=1,
                sample_text="回收到的文字示例",
                line_count=4,
            )
        ]
        with (
            patch("core.pdf_ocr_engine.is_ocr_available", return_value=True),
            patch(
                "core.pdf_ocr_diagnostics._probe_ocr_recovery",
                return_value=fake_previews,
            ) as probe_mock,
        ):
            report = diagnose_project_ocr_risks(project, include_ocr_probe=True)
        probe_mock.assert_called_once()
        self.assertTrue(report.ocr_available)
        self.assertEqual(report.ocr_recoverable_samples, fake_previews)
        self.assertTrue(
            any("OCR 预览" in suggestion for suggestion in report.suggestions),
            report.suggestions,
        )

    def test_report_to_dict_contains_new_fields(self):
        project = self._build_image_only_project()
        with patch("core.pdf_ocr_engine.is_ocr_dependency_available", return_value=False):
            report = diagnose_project_ocr_risks(project)
        payload = report.to_dict()
        self.assertIn("ocr_available", payload)
        self.assertIn("ocr_recoverable_samples", payload)
        self.assertEqual(payload["ocr_available"], False)
        self.assertEqual(payload["ocr_recoverable_samples"], [])


class ProbeOCRRecoveryHelperTest(unittest.TestCase):
    def test_returns_empty_when_pdf_path_missing(self):
        self.assertEqual(
            pdf_ocr_diagnostics._probe_ocr_recovery(None, [1, 2, 3]), []
        )
        self.assertEqual(
            pdf_ocr_diagnostics._probe_ocr_recovery("/tmp/x.pdf", []), []
        )

    def test_caps_probe_count_and_trims_sample(self):
        long_text_lines = [
            OCRLine(
                text="甲" * 30,
                bbox=(0.0, float(i * 10), 100.0, float(i * 10 + 10)),
                confidence=0.9,
                page_number=1,
            )
            for i in range(3)
        ]
        with (
            patch("core.pdf_ocr_engine.is_ocr_available", return_value=True),
            patch(
                "core.pdf_ocr_engine.ocr_pdf_page",
                return_value=long_text_lines,
            ) as ocr_mock,
        ):
            previews = pdf_ocr_diagnostics._probe_ocr_recovery(
                "/tmp/sample.pdf",
                [1, 2, 3, 4, 5],
            )
        self.assertEqual(ocr_mock.call_count, pdf_ocr_diagnostics._OCR_PROBE_PAGE_LIMIT)
        self.assertEqual(len(previews), pdf_ocr_diagnostics._OCR_PROBE_PAGE_LIMIT)
        self.assertTrue(all(preview.sample_text.endswith("...") for preview in previews))


if __name__ == "__main__":
    unittest.main()
