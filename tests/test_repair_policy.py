import pickle
import tempfile
import unittest
from os import environ
from pathlib import Path
from unittest.mock import patch

from core.ai_repair import AIRepairService, repair_project_questions
from core.project_quality import annotate_project_quality
from core.repair_action_model import clear_repair_model_cache
from domain.models import ExamProject, MaterialSet, OptionNode, QuestionNode, Section


class DummyRepairFamilyModel:
    class _DecisionMatrix(list):
        ndim = 2

    def decision_function(self, rows):
        text = rows[0]["text"]
        if "同比增长8.4%" in text:
            return self._DecisionMatrix([[-0.4, 0.0, 2.8]])
        return self._DecisionMatrix([[0.0, 2.5, -0.4]])


def _build_policy_project() -> ExamProject:
    return ExamProject(
        title="policy",
        sections=[
            Section(
                kind="unknown",
                title="待确认",
                questions=[
                    QuestionNode(
                        source_number="1",
                        stem="下列关于法律常识的说法正确的是",
                        options=[
                            OptionNode(letter="A", text="甲"),
                            OptionNode(letter="B", text="乙"),
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
                                stem="2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。根据上述资料，下列说法正确的是",
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


class RepairPolicyTest(unittest.TestCase):
    def tearDown(self):
        clear_repair_model_cache()

    def test_policy_mode_reranks_flagged_questions_by_predicted_repair_family(self):
        baseline_project = _build_policy_project()
        annotate_project_quality(baseline_project)

        baseline_summary = repair_project_questions(
            baseline_project,
            service=AIRepairService(mode="balanced"),
            only_flagged=True,
            limit=1,
        )

        self.assertEqual(baseline_summary.attempted_questions, 1)
        self.assertEqual(baseline_project.sections[0].kind, "common_sense")
        self.assertEqual(baseline_project.sections[1].material_sets[0].body, "")

        policy_project = _build_policy_project()
        annotate_project_quality(policy_project)

        with tempfile.TemporaryDirectory() as tmp:
            model_path = Path(tmp) / "repair_action_family_classifier.pkl"
            bundle = {
                "labels": ["asset_repair", "boundary_repair", "material_repair"],
                "label_field": "action_family",
                "model": DummyRepairFamilyModel(),
                "ready_for_runtime": True,
            }
            with model_path.open("wb") as fh:
                pickle.dump(bundle, fh)

            with patch.dict(
                environ,
                {
                    "PPTCONVERT_ENABLE_PICKLED_REPAIR_MODEL": "1",
                    "PPTCONVERT_TRUST_REPAIR_MODEL": "1",
                    "PPTCONVERT_REPAIR_FAMILY_MODEL": str(model_path),
                },
                clear=False,
            ):
                policy_summary = repair_project_questions(
                    policy_project,
                    service=AIRepairService(mode="policy"),
                    only_flagged=True,
                    limit=1,
                )

        self.assertEqual(policy_summary.attempted_questions, 1)
        self.assertEqual(policy_project.sections[0].kind, "unknown")
        material = policy_project.sections[1].material_sets[0]
        question = material.questions[0]
        self.assertEqual(material.body, "2023年全市规模以上工业增加值同比增长8.4%，其中制造业增加值占比进一步提升。")
        self.assertEqual(question.stem, "根据上述资料，下列说法正确的是")


if __name__ == "__main__":
    unittest.main()
