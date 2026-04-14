import os
import tempfile
import unittest

from domain.models import (
    AssetRef,
    ExamProject,
    MaterialSet,
    OptionNode,
    PageRegion,
    PaperSource,
    QuestionNode,
    QuestionRange,
    ReviewIssue,
    Section,
)
from exporters.manifest_json import export_project_manifest, load_project_manifest_project


class ManifestJsonTest(unittest.TestCase):
    def test_manifest_roundtrip_restores_project_dataclasses(self):
        with tempfile.TemporaryDirectory() as temp_dir:
            manifest_path = os.path.join(temp_dir, "project.json")
            project = ExamProject(
                title="示例工程",
                source=PaperSource(
                    pdf_path="sample.pdf",
                    asset_dir=os.path.join(temp_dir, "assets"),
                ),
                sections=[
                    Section(
                        kind="data",
                        title="资料分析",
                        material_sets=[
                            MaterialSet(
                                material_id="data-1-1",
                                header="材料一",
                                body="材料正文",
                                body_lines=["材料正文"],
                                body_assets=[
                                    AssetRef(
                                        kind="material_inline_image",
                                        path="material.png",
                                        source_page=2,
                                        page_region=PageRegion(page_number=2, x0=1, y0=2, x1=3, y1=4),
                                    )
                                ],
                                body_regions=[PageRegion(page_number=2, x0=1, y0=2, x1=30, y1=40)],
                                questions=[
                                    QuestionNode(
                                        source_number="111",
                                        stem="题干",
                                        options=[
                                            OptionNode(
                                                letter="A",
                                                text="甲",
                                                image_path="option-a.png",
                                                source_page=2,
                                                page_region=PageRegion(page_number=2, x0=5, y0=6, x1=7, y1=8),
                                            )
                                        ],
                                        stem_assets=[
                                            AssetRef(
                                                kind="stem_image",
                                                path="stem.png",
                                                source_page=2,
                                            )
                                        ],
                                        page_numbers=[2],
                                        option_layout="one_row",
                                        review_confidence=0.62,
                                        review_issues=[
                                            ReviewIssue(
                                                code="option_count",
                                                title="选项数量异常",
                                                detail="当前识别到 1 个选项。",
                                                severity="warning",
                                            )
                                        ],
                                        suggested_subject="data",
                                        suggested_subject_confidence=0.84,
                                        suggested_subject_reason="题干里有明显资料分析信号。",
                                        inferred_subtype="表格型资料分析",
                                        inferred_subtype_confidence=0.88,
                                        inferred_signals=["根据以下资料", "同比", "表中"],
                                    )
                                ],
                            )
                        ],
                    )
                ],
                selected_subjects=["data"],
                selected_ranges=[QuestionRange(start=111, end=115)],
            )

            export_project_manifest(project, manifest_path)
            loaded = load_project_manifest_project(manifest_path)

            self.assertEqual(loaded.title, "示例工程")
            self.assertEqual(loaded.source.pdf_path, "sample.pdf")
            self.assertEqual(loaded.selected_subjects, ["data"])
            self.assertEqual(len(loaded.selected_ranges), 1)
            self.assertEqual(loaded.selected_ranges[0].start, 111)
            question = loaded.sections[0].material_sets[0].questions[0]
            self.assertEqual(question.options[0].image_path, "option-a.png")
            self.assertEqual(question.options[0].source_page, 2)
            self.assertIsNotNone(question.options[0].page_region)
            self.assertEqual(question.review_issues[0].code, "option_count")
            self.assertAlmostEqual(question.review_confidence, 0.62, places=2)
            self.assertEqual(question.suggested_subject, "data")
            self.assertEqual(question.inferred_subtype, "表格型资料分析")
            self.assertAlmostEqual(question.inferred_subtype_confidence, 0.88, places=2)
            self.assertEqual(question.inferred_signals[:2], ["根据以下资料", "同比"])


if __name__ == "__main__":
    unittest.main()
