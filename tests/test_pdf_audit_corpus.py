import unittest
from pathlib import Path
from unittest.mock import patch

from core.project_quality import ProjectQualitySummary
from domain.models import ExamProject, OptionNode, QuestionNode, ReviewIssue, Section
from scripts.audit_pdf_corpus import _detect_source_gap_numbers, audit_pdf


class PdfAuditCorpusTest(unittest.TestCase):
    def test_audit_pdf_reports_quality_and_source_defects(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="问以下哪个坐标图能准确表示关系（）",
                            options=[],
                            review_issues=[
                                ReviewIssue(
                                    code="source_visual_missing",
                                    title="源 PDF 疑似缺少图形选项",
                                    severity="error",
                                ),
                                ReviewIssue(
                                    code="option_count",
                                    title="选项数量异常",
                                    severity="error",
                                ),
                            ],
                        ),
                        QuestionNode(
                            source_number="2",
                            stem="普通题目",
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

        with patch("scripts.audit_pdf_corpus.build_exam_project_from_pdf", return_value=project), patch(
            "scripts.audit_pdf_corpus.annotate_project_quality",
            return_value=ProjectQualitySummary(
                question_count=2,
                flagged_questions=1,
                severe_questions=1,
                source_defect_questions=1,
                total_issue_count=2,
            ),
        ):
            row = audit_pdf(Path("示例（2题）.pdf"))

        self.assertEqual(row.flagged_questions, 1)
        self.assertEqual(row.severe_questions, 1)
        self.assertEqual(row.source_defect_questions, 1)
        self.assertEqual(row.top_issue_codes[0]["code"], "option_count")
        self.assertEqual(row.top_issue_codes[0]["count"], 1)
        self.assertEqual(row.severe_samples[0]["issue_codes"], ["source_visual_missing", "option_count"])
        self.assertEqual(row.severe_samples[0]["defect_type"], "source_visual_missing")
        self.assertEqual(row.severe_samples[0]["page_numbers"], [])
        self.assertFalse(row.severe_samples[0]["page_numbers_inferred"])

    def test_detect_source_gap_numbers_marks_direct_page_number_jumps(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(source_number="565", stem="上一题", page_numbers=[108]),
                        QuestionNode(source_number="567", stem="下一题", page_numbers=[108]),
                        QuestionNode(source_number="974", stem="上一题", page_numbers=[182]),
                        QuestionNode(source_number="976", stem="下一题", page_numbers=[182]),
                    ],
                )
            ],
        )

        with patch(
            "scripts.audit_pdf_corpus._page_text_map",
            return_value={
                108: "565、上一题\n567、下一题",
                182: "974. 上一题\n976. 下一题",
            },
        ):
            source_gap_numbers = _detect_source_gap_numbers(project, Path("示例.pdf"), [566, 975])

        self.assertEqual(source_gap_numbers, [566, 975])

    def test_audit_pdf_separates_parser_missing_from_source_gap(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(source_number="1", stem="第一题"),
                        QuestionNode(source_number="3", stem="第三题"),
                    ],
                )
            ],
        )

        with patch("scripts.audit_pdf_corpus.build_exam_project_from_pdf", return_value=project), patch(
            "scripts.audit_pdf_corpus.annotate_project_quality",
            return_value=ProjectQualitySummary(question_count=2),
        ), patch("scripts.audit_pdf_corpus._detect_source_gap_numbers", return_value=[2]):
            row = audit_pdf(Path("示例（3题）.pdf"))

        self.assertEqual(row.missing_numbers, [2])
        self.assertEqual(row.parser_missing_numbers, [])
        self.assertEqual(row.source_gap_numbers, [2])

    def test_effective_question_pages_can_infer_placeholder_page_from_neighbors(self):
        project = ExamProject(
            title="示例",
            sections=[
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(source_number="1332", stem="上一题", page_numbers=[244]),
                        QuestionNode(source_number="1333", stem="", page_numbers=[]),
                        QuestionNode(source_number="1334", stem="下一题", page_numbers=[245]),
                    ],
                )
            ],
        )

        with patch("scripts.audit_pdf_corpus.build_exam_project_from_pdf", return_value=project), patch(
            "scripts.audit_pdf_corpus.annotate_project_quality",
            return_value=ProjectQualitySummary(
                question_count=3,
                flagged_questions=1,
                severe_questions=1,
                source_defect_questions=1,
            ),
        ):
            project.sections[0].questions[1].review_issues = [
                ReviewIssue(code="source_text_missing", title="源 PDF 题目文本缺失", severity="error")
            ]
            row = audit_pdf(Path("示例.pdf"))

        sample = row.severe_samples[0]
        self.assertEqual(sample["page_numbers"], [244, 245])
        self.assertTrue(sample["page_numbers_inferred"])
        self.assertEqual(sample["defect_type"], "source_text_missing")


if __name__ == "__main__":
    unittest.main()
