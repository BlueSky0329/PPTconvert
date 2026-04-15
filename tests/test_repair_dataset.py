import json
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section
from scripts.build_repair_action_dataset import build_repair_action_dataset, build_repair_state_record, generate_project_repair_rows


class RepairActionDatasetTest(unittest.TestCase):
    def _build_project(self) -> ExamProject:
        return ExamProject(
            title="修复训练集示例",
            sections=[
                Section(
                    kind="quant",
                    title="数量关系",
                    questions=[
                        QuestionNode(
                            source_number="71",
                            stem="甲乙两地相距 100 千米，问相遇时间。",
                            options=[
                                OptionNode("A", "4"),
                                OptionNode("B", "5"),
                                OptionNode("C", "6"),
                                OptionNode("D", "8"),
                            ],
                        ),
                        QuestionNode(
                            source_number="72",
                            stem="某商品按8折出售后利润率为20%，其成本是多少？",
                            options=[
                                OptionNode("A", "80"),
                                OptionNode("B", "96"),
                                OptionNode("C", "100"),
                                OptionNode("D", "120"),
                            ],
                        ),
                    ],
                ),
                Section(
                    kind="reasoning",
                    title="判断推理",
                    questions=[
                        QuestionNode(
                            source_number="81",
                            stem="请选择最合适的一项。",
                            options=[
                                OptionNode("A", "", image_path="a.png"),
                                OptionNode("B", "", image_path="b.png"),
                                OptionNode("C", "", image_path="c.png"),
                                OptionNode("D", "", image_path="d.png"),
                            ],
                        ),
                    ],
                ),
                Section(
                    kind="data",
                    title="资料分析",
                    material_sets=[
                        MaterialSet(
                            material_id="m1",
                            header="材料一",
                            body="2024年全市工业增加值同比增长8.4%。",
                            body_lines=["2024年全市工业增加值同比增长8.4%。"],
                            body_assets=[AssetRef(kind="image", path="chart.png")],
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
                ),
            ],
        )

    def test_generate_project_repair_rows_covers_core_actions(self):
        project = self._build_project()

        rows = generate_project_repair_rows(
            project,
            source_pdf="sample.pdf",
            source_form="set_paper",
        )

        actions = {row["action"] for row in rows}
        self.assertIn("renumber_current_question", actions)
        self.assertIn("split_embedded_next_question", actions)
        self.assertIn("move_spilled_option_back", actions)
        self.assertIn("move_data_intro_back_to_material", actions)
        self.assertIn("move_data_assets_to_material", actions)
        self.assertIn("reassign_stem_image_to_options", actions)
        families = {row["action_family"] for row in rows}
        self.assertIn("boundary_repair", families)
        self.assertIn("material_repair", families)
        self.assertIn("asset_repair", families)

    def test_build_repair_state_record_includes_neighbor_context(self):
        project = self._build_project()
        section = project.sections[0]
        previous_question = section.questions[0]
        question = section.questions[1]

        record = build_repair_state_record(
            section=section,
            material=None,
            question=question,
            previous_question=previous_question,
            next_question=None,
        )

        self.assertIn("[SECTION] quant / 数量关系", record["text"])
        self.assertIn("[PREV]", record["text"])
        self.assertIn("[CURRENT]", record["text"])
        self.assertEqual(record["meta"]["prev_present"], 1.0)
        self.assertEqual(record["meta"]["next_present"], 0.0)
        self.assertGreater(record["meta"]["stem_length"], 0.0)
        self.assertEqual(record["meta"]["previous_option_count"], 4.0)

    def test_build_repair_action_dataset_uses_subject_hint_for_single_subject_books(self):
        project = self._build_project()

        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            pdf_path = root / "sample.pdf"
            pdf_path.write_bytes(b"%PDF-1.4\n")
            catalog_path = root / "catalog.json"
            output_path = root / "repair_actions.jsonl"
            catalog_path.write_text(
                json.dumps(
                    {
                        "pdfs": [
                            {
                                "path": "sample.pdf",
                                "form": "single_subject_book",
                                "subject": "data",
                            }
                        ]
                    },
                    ensure_ascii=False,
                ),
                encoding="utf-8",
            )

            with patch("scripts.build_repair_action_dataset.ROOT", root), patch(
                "scripts.build_repair_action_dataset.build_exam_project_from_pdf",
                return_value=project,
            ) as build_mock:
                summary = build_repair_action_dataset(catalog_path, output_path)

        build_mock.assert_called_once_with(
            str(pdf_path),
            mode="all",
            document_subject_hint="data",
        )
        self.assertEqual(summary["pdf_count"], 1)
        self.assertGreater(summary["row_count"], 0)


if __name__ == "__main__":
    unittest.main()
