import json
import os
import tempfile
import unittest

from core.repair_log import append_question_repair_log, capture_question_state
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section
from exporters.manifest_json import export_project_manifest
from scripts.build_repair_trajectory_dataset import (
    build_trajectories_from_gui_dataset_files,
    build_trajectories_from_gui_logs,
    generate_trajectories_for_project,
)


def _build_project() -> ExamProject:
    return ExamProject(
        title="trajectory_test",
        sections=[
            Section(
                kind="verbal",
                title="言语理解",
                questions=[
                    QuestionNode(
                        source_number="1",
                        stem="这段文字意在说明",
                        options=[
                            OptionNode(letter="A", text="甲"),
                            OptionNode(letter="B", text="乙"),
                            OptionNode(letter="C", text="丙"),
                            OptionNode(letter="D", text="丁"),
                        ],
                    ),
                    QuestionNode(
                        source_number="2",
                        stem="这段文字的主旨是",
                        options=[
                            OptionNode(letter="A", text="甲"),
                            OptionNode(letter="B", text="乙"),
                            OptionNode(letter="C", text="丙"),
                            OptionNode(letter="D", text="丁"),
                        ],
                    ),
                    QuestionNode(
                        source_number="3",
                        stem="作者意在强调的是",
                        options=[
                            OptionNode(letter="A", text="甲", image_path="a.png"),
                            OptionNode(letter="B", text="乙", image_path="b.png"),
                            OptionNode(letter="C", text="丙", image_path="c.png"),
                            OptionNode(letter="D", text="丁", image_path="d.png"),
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
                        body="2023年全市GDP同比增长5.2%",
                        body_assets=[
                            AssetRef(kind="material_inline_image", path="chart.png"),
                        ],
                        questions=[
                            QuestionNode(
                                source_number="101",
                                stem="根据上述资料，下列说法正确的是",
                                options=[
                                    OptionNode(letter="A", text="甲"),
                                    OptionNode(letter="B", text="乙"),
                                    OptionNode(letter="C", text="丙"),
                                    OptionNode(letter="D", text="丁"),
                                ],
                            ),
                            QuestionNode(
                                source_number="102",
                                stem="下列说法错误的是",
                                options=[
                                    OptionNode(letter="A", text="甲"),
                                    OptionNode(letter="B", text="乙"),
                                    OptionNode(letter="C", text="丙"),
                                    OptionNode(letter="D", text="丁"),
                                ],
                            ),
                        ],
                    ),
                ],
            ),
        ],
    )


class TrajectoryDatasetTest(unittest.TestCase):
    def test_generates_single_step_trajectories(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        self.assertGreater(len(trajectories), 0)
        for traj in trajectories:
            self.assertIn("source_pdf", traj)
            self.assertEqual(traj["source_pdf"], "test.pdf")
            self.assertIn("steps", traj)
            self.assertGreater(len(traj["steps"]), 0)
            for step in traj["steps"]:
                self.assertIn("action", step)
                self.assertIn("state_record", step)
                self.assertIn("text", step["state_record"])
                self.assertIn("meta", step["state_record"])

    def test_single_step_trajectories_have_length_one(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        single_step = [t for t in trajectories if t["trajectory_length"] == 1]
        self.assertGreater(len(single_step), 0)
        for traj in single_step:
            self.assertEqual(len(traj["steps"]), 1)
            self.assertEqual(traj["steps"][0]["step"], 0)

    def test_multi_step_trajectories_combine_injectors(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        multi_step = [t for t in trajectories if t["trajectory_length"] >= 2]
        for traj in multi_step:
            self.assertEqual(len(traj["steps"]), traj["trajectory_length"])
            actions = [step["action"] for step in traj["steps"]]
            self.assertEqual(len(actions), len(set(actions)), "multi-step actions should be unique")

    def test_trajectories_include_action_family(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        valid_families = {"boundary_repair", "material_repair", "asset_repair", "other_repair"}
        for traj in trajectories:
            for step in traj["steps"]:
                self.assertIn("action_family", step)
                self.assertIn(step["action_family"], valid_families)

    def test_data_section_questions_produce_trajectories(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        data_trajectories = [t for t in trajectories if t["subject"] == "data"]
        self.assertGreater(len(data_trajectories), 0)

    def test_material_and_asset_actions_are_present_when_applicable(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        actions = {
            step["action"]
            for trajectory in trajectories
            for step in trajectory["steps"]
        }
        self.assertIn("move_data_intro_back_to_material", actions)
        self.assertIn("move_data_assets_to_material", actions)
        self.assertIn("reassign_stem_image_to_options", actions)


class TrajectoryInjectorTest(unittest.TestCase):
    def test_number_shift_injector(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        renumber_trajectories = [
            t for t in trajectories
            if any(step["action"] == "renumber_current_question" for step in t["steps"])
        ]
        self.assertGreater(len(renumber_trajectories), 0)

    def test_spilled_option_injector(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        spilled_trajectories = [
            t for t in trajectories
            if any(step["action"] == "move_spilled_option_back" for step in t["steps"])
        ]
        self.assertGreater(len(spilled_trajectories), 0)

    def test_embedded_next_injector(self):
        project = _build_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="test.pdf", source_form="test",
        )

        embedded_trajectories = [
            t for t in trajectories
            if any(step["action"] == "split_embedded_next_question" for step in t["steps"])
        ]
        self.assertGreater(len(embedded_trajectories), 0)

    def test_builds_suffix_trajectories_from_gui_repair_logs(self):
        project = ExamProject(
            title="gui_trajectory_test",
            sections=[
                Section(
                    kind="verbal",
                    title="言语理解",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="这段文字意在说明",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="3",
                            stem="D. 丁 3. 这段文字的主旨是",
                            options=[
                                OptionNode(letter="A", text="甲"),
                                OptionNode(letter="B", text="乙"),
                                OptionNode(letter="C", text="丙"),
                                OptionNode(letter="D", text="丁"),
                            ],
                        ),
                    ],
                ),
            ],
        )
        project.source.pdf_path = "gui_demo.pdf"
        section = project.sections[0]
        previous_question = section.questions[0]
        question = section.questions[1]

        before_state = capture_question_state(
            section=section,
            material=None,
            question=question,
            previous_question=previous_question,
            next_question=None,
        )
        question.source_number = "2"
        append_question_repair_log(
            project,
            source="gui_manual",
            action="renumber_question",
            section=section,
            material=None,
            question=question,
            previous_question=previous_question,
            next_question=None,
            before_state=before_state,
            metadata={"new_number": "2"},
        )

        before_state = capture_question_state(
            section=section,
            material=None,
            question=question,
            previous_question=previous_question,
            next_question=None,
        )
        question.stem = "这段文字的主旨是"
        append_question_repair_log(
            project,
            source="gui_manual",
            action="update_question_stem",
            section=section,
            material=None,
            question=question,
            previous_question=previous_question,
            next_question=None,
            before_state=before_state,
            metadata={"field": "stem"},
        )

        with tempfile.TemporaryDirectory() as temp_dir:
            manifest_path = os.path.join(temp_dir, "project.json")
            export_project_manifest(project, manifest_path)
            trajectories = build_trajectories_from_gui_logs([manifest_path])

        self.assertEqual(len(trajectories), 2)
        full_trajectory = next(item for item in trajectories if item["trajectory_length"] == 2)
        self.assertEqual(full_trajectory["source_form"], "gui_repair_log")
        self.assertEqual(full_trajectory["trajectory_origin"], "gui_log_manifest")
        self.assertEqual(
            [step["action"] for step in full_trajectory["steps"]],
            ["renumber_current_question", "move_spilled_option_back"],
        )
        self.assertEqual(full_trajectory["source_pdf"], "gui_demo.pdf")

    def test_builds_suffix_trajectories_from_gui_jsonl_dataset(self):
        rows = [
            {
                "session_id": "session-jsonl",
                "manifest_path": "demo_manifest.json",
                "source_pdf": "jsonl_demo.pdf",
                "timestamp": "2026-04-15T10:00:00Z",
                "source": "gui_manual",
                "action": "renumber_question",
                "question_id": "q-jsonl-1",
                "question_no": "2",
                "section_kind": "verbal",
                "material_id": "",
                "before_state": {
                    "question_id": "q-jsonl-1",
                    "question_no": "3",
                    "section_kind": "verbal",
                    "material_id": "",
                    "state_record": {
                        "text": "题号：3\n题干：D. 丁 3. 这段文字的主旨是",
                        "meta": {
                            "number_gap_from_previous": 2.0,
                            "stem_has_embedded_question_no": 1.0,
                            "stem_starts_with_option_marker": 1.0,
                        },
                    },
                },
                "after_state": {
                    "question_id": "q-jsonl-1",
                    "question_no": "2",
                    "section_kind": "verbal",
                    "material_id": "",
                    "state_record": {
                        "text": "题号：2\n题干：D. 丁 2. 这段文字的主旨是",
                        "meta": {
                            "number_gap_from_previous": 1.0,
                            "stem_has_embedded_question_no": 1.0,
                            "stem_starts_with_option_marker": 1.0,
                        },
                    },
                },
                "metadata": {"new_number": "2"},
            },
            {
                "session_id": "session-jsonl",
                "manifest_path": "demo_manifest.json",
                "source_pdf": "jsonl_demo.pdf",
                "timestamp": "2026-04-15T10:01:00Z",
                "source": "gui_manual",
                "action": "update_question_stem",
                "question_id": "q-jsonl-1",
                "question_no": "2",
                "section_kind": "verbal",
                "material_id": "",
                "before_state": {
                    "question_id": "q-jsonl-1",
                    "question_no": "2",
                    "section_kind": "verbal",
                    "material_id": "",
                    "state_record": {
                        "text": "题号：2\n题干：D. 丁 2. 这段文字的主旨是",
                        "meta": {
                            "number_gap_from_previous": 1.0,
                            "stem_has_embedded_question_no": 1.0,
                            "stem_starts_with_option_marker": 1.0,
                        },
                    },
                },
                "after_state": {
                    "question_id": "q-jsonl-1",
                    "question_no": "2",
                    "section_kind": "verbal",
                    "material_id": "",
                    "state_record": {
                        "text": "题号：2\n题干：这段文字的主旨是",
                        "meta": {
                            "number_gap_from_previous": 1.0,
                            "stem_has_embedded_question_no": 0.0,
                            "stem_starts_with_option_marker": 0.0,
                        },
                    },
                },
                "metadata": {"field": "stem"},
            },
        ]

        with tempfile.TemporaryDirectory() as temp_dir:
            dataset_path = os.path.join(temp_dir, "gui_repair_logs.jsonl")
            with open(dataset_path, "w", encoding="utf-8") as fh:
                for row in rows:
                    fh.write(json.dumps(row, ensure_ascii=False) + "\n")
            trajectories = build_trajectories_from_gui_dataset_files([dataset_path])

        self.assertEqual(len(trajectories), 2)
        full_trajectory = next(item for item in trajectories if item["trajectory_length"] == 2)
        self.assertEqual(full_trajectory["source_form"], "gui_repair_log_jsonl")
        self.assertEqual(full_trajectory["trajectory_origin"], "gui_log_jsonl")
        self.assertEqual(full_trajectory["source_pdf"], "jsonl_demo.pdf")
        self.assertEqual(
            [step["action"] for step in full_trajectory["steps"]],
            ["renumber_current_question", "move_spilled_option_back"],
        )


if __name__ == "__main__":
    unittest.main()
