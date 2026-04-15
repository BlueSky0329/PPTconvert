import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

import core.repair_trajectory_model as trajectory_model
from core.repair_trajectory_model import (
    TrajectoryPlan,
    TrajectoryStep,
    build_rule_based_trajectory_plan,
    clear_trajectory_model_cache,
    predict_trajectory_plan,
)
from domain.models import MaterialSet, OptionNode, QuestionNode, Section


class DummyTrajectoryModel:
    class _DecisionMatrix(list):
        ndim = 2

    def decision_function(self, rows):
        text = rows[0]["text"]
        if "同比增长8.4%" in text:
            return self._DecisionMatrix([[-0.3, 0.1, 0.5, -0.8, 2.6, -0.1]])
        return self._DecisionMatrix([[2.5, -0.4, 0.1, -0.6, 0.3, -0.2]])


class TrajectoryModelLoadTest(unittest.TestCase):
    def tearDown(self):
        clear_trajectory_model_cache()

    def test_predict_trajectory_plan_returns_plan_from_trusted_model(self):
        section = Section(kind="data", title="资料分析")
        material = MaterialSet(material_id="m1", header="材料一", body="")
        question = QuestionNode(
            source_number="101",
            stem="2023年全市规模以上工业增加值同比增长8.4%，根据上述资料，下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_trajectory_policy.pkl"
            bundle = {
                "labels": [
                    "move_data_assets_to_material",
                    "move_data_intro_back_to_material",
                    "move_spilled_option_back",
                    "reassign_stem_image_to_options",
                    "renumber_current_question",
                    "split_embedded_next_question",
                ],
                "label_field": "action",
                "model": DummyTrajectoryModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_TRAJECTORY_MODEL": "1",
                    "PPTCONVERT_TRUST_TRAJECTORY_MODEL": "1",
                    "PPTCONVERT_TRAJECTORY_MODEL": str(model_path),
                },
                clear=False,
            ):
                plan = predict_trajectory_plan(
                    section=section,
                    material=material,
                    question=question,
                )

        self.assertIsNotNone(plan)
        assert plan is not None
        self.assertFalse(plan.is_empty)
        self.assertEqual(plan.primary_action, "renumber_current_question")
        self.assertGreater(plan.primary_confidence, 0.3)

    def test_predict_trajectory_plan_returns_none_without_model(self):
        section = Section(kind="verbal", title="言语理解")
        question = QuestionNode(
            source_number="1",
            stem="下列关于法律常识的说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with patch.dict(
            environ,
            {"PPTCONVERT_TRAJECTORY_MODEL": "", "PPTCONVERT_ENABLE_TRAJECTORY_MODEL": "0"},
            clear=False,
        ):
            plan = predict_trajectory_plan(section=section, material=None, question=question)

        self.assertIsNone(plan)

    def test_predict_trajectory_plan_skips_not_ready_bundle(self):
        section = Section(kind="data", title="资料分析")
        material = MaterialSet(material_id="m1", header="材料一", body="")
        question = QuestionNode(
            source_number="101",
            stem="同比增长8.4%，下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_trajectory_policy.pkl"
            bundle = {
                "labels": ["renumber_current_question", "split_embedded_next_question"],
                "label_field": "action",
                "model": DummyTrajectoryModel(),
                "ready_for_runtime": False,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_TRAJECTORY_MODEL": "1",
                    "PPTCONVERT_TRUST_TRAJECTORY_MODEL": "1",
                    "PPTCONVERT_TRAJECTORY_MODEL": str(model_path),
                },
                clear=False,
            ):
                plan = predict_trajectory_plan(
                    section=section,
                    material=material,
                    question=question,
                )

        self.assertIsNone(plan)

    def test_predict_trajectory_loads_default_path_model(self):
        section = Section(kind="verbal", title="言语理解")
        question = QuestionNode(
            source_number="5",
            stem="下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_trajectory_policy.pkl"
            bundle = {
                "labels": [
                    "move_data_assets_to_material",
                    "move_data_intro_back_to_material",
                    "move_spilled_option_back",
                    "reassign_stem_image_to_options",
                    "renumber_current_question",
                    "split_embedded_next_question",
                ],
                "label_field": "action",
                "model": DummyTrajectoryModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with (
                patch.object(trajectory_model, "_DEFAULT_MODEL_PATH", model_path),
                patch.object(trajectory_model, "_TRUSTED_MODEL_DIR", model_path.parent.resolve()),
                patch.dict(
                    environ,
                    {
                        "PPTCONVERT_ENABLE_TRAJECTORY_MODEL": "",
                        "PPTCONVERT_TRAJECTORY_MODEL": "",
                        "PPTCONVERT_TRUST_TRAJECTORY_MODEL": "",
                    },
                    clear=False,
                ),
            ):
                plan = predict_trajectory_plan(section=section, material=None, question=question)

        self.assertIsNotNone(plan)
        assert plan is not None
        self.assertEqual(plan.primary_action, "renumber_current_question")


class RuleBasedTrajectoryTest(unittest.TestCase):
    def test_rule_detects_number_gap(self):
        section = Section(kind="verbal", title="言语理解")
        prev_q = QuestionNode(
            source_number="5",
            stem="这段文字意在说明",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )
        cur_q = QuestionNode(
            source_number="7",
            stem="这段文字的主旨是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=None,
            question=cur_q,
            previous_question=prev_q,
        )

        self.assertFalse(plan.is_empty)
        self.assertEqual(plan.primary_action, "renumber_current_question")

    def test_rule_detects_embedded_next_question(self):
        section = Section(kind="quant", title="数量关系")
        cur_q = QuestionNode(
            source_number="10",
            stem="某工厂有100名工人 11. 下一题题干内容",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )
        next_q = QuestionNode(source_number="11", stem="", options=[])

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=None,
            question=cur_q,
            next_question=next_q,
        )

        actions = plan.actions_in_order()
        self.assertIn("split_embedded_next_question", actions)

    def test_rule_detects_spilled_option(self):
        section = Section(kind="common_sense", title="常识判断")
        prev_q = QuestionNode(
            source_number="3",
            stem="下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
            ],
        )
        cur_q = QuestionNode(
            source_number="4",
            stem="D. 丁选项文字 4. 真正的题干内容",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=None,
            question=cur_q,
            previous_question=prev_q,
        )

        actions = plan.actions_in_order()
        self.assertIn("move_spilled_option_back", actions)

    def test_rule_detects_material_intro_recovery(self):
        section = Section(kind="data", title="资料分析")
        material = MaterialSet(material_id="m1", header="材料一", body="", body_lines=[], body_assets=[])
        question = QuestionNode(
            source_number="101",
            stem="2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。根据上述资料，下列说法正确的是",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )
        material.questions = [question]

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=material,
            question=question,
        )

        actions = plan.actions_in_order()
        self.assertIn("move_data_intro_back_to_material", actions)

    def test_rule_detects_stem_assets_to_material(self):
        section = Section(kind="data", title="资料分析")
        material = MaterialSet(
            material_id="m1", header="材料一", body="一段材料",
            body_lines=[], body_assets=[],
        )
        question = QuestionNode(
            source_number="101",
            stem="根据上述资料，下列说法正确的是",
            stem_assets=[{"kind": "image", "path": "chart.png"}],
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )
        material.questions = [question]

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=material,
            question=question,
        )

        actions = plan.actions_in_order()
        self.assertIn("move_data_assets_to_material", actions)

    def test_rule_detects_image_option_reassignment(self):
        section = Section(kind="reasoning", title="判断推理")
        question = QuestionNode(
            source_number="50",
            stem="请选出正确的一项",
            stem_assets=[
                {"kind": "image", "path": "a.png"},
                {"kind": "image", "path": "b.png"},
                {"kind": "image", "path": "c.png"},
                {"kind": "image", "path": "d.png"},
            ],
            options=[
                OptionNode(letter="A", text=""),
                OptionNode(letter="B", text=""),
                OptionNode(letter="C", text=""),
                OptionNode(letter="D", text=""),
            ],
        )

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=None,
            question=question,
        )

        actions = plan.actions_in_order()
        self.assertIn("reassign_stem_image_to_options", actions)

    def test_rule_returns_empty_for_clean_question(self):
        section = Section(kind="verbal", title="言语理解")
        prev_q = QuestionNode(
            source_number="1",
            stem="题干内容",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )
        cur_q = QuestionNode(
            source_number="2",
            stem="正常的题干内容",
            options=[
                OptionNode(letter="A", text="甲"),
                OptionNode(letter="B", text="乙"),
                OptionNode(letter="C", text="丙"),
                OptionNode(letter="D", text="丁"),
            ],
        )

        plan = build_rule_based_trajectory_plan(
            section=section,
            material=None,
            question=cur_q,
            previous_question=prev_q,
        )

        self.assertTrue(plan.is_empty)

    def test_trajectory_plan_properties(self):
        plan = TrajectoryPlan(steps=[
            TrajectoryStep(action="renumber_current_question", confidence=0.9, action_family="boundary_repair"),
            TrajectoryStep(action="move_spilled_option_back", confidence=0.7, action_family="boundary_repair"),
        ])

        self.assertFalse(plan.is_empty)
        self.assertEqual(plan.primary_action, "renumber_current_question")
        self.assertEqual(plan.primary_confidence, 0.9)
        self.assertEqual(plan.actions_in_order(), ["renumber_current_question", "move_spilled_option_back"])

    def test_empty_trajectory_plan(self):
        plan = TrajectoryPlan(steps=[])

        self.assertTrue(plan.is_empty)
        self.assertIsNone(plan.primary_action)
        self.assertEqual(plan.primary_confidence, 0.0)
        self.assertEqual(plan.actions_in_order(), [])


if __name__ == "__main__":
    unittest.main()
