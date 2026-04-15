import json
import tempfile
import unittest
from pathlib import Path

from scripts.build_repair_trajectory_dataset import generate_trajectories_for_project
from domain.models import ExamProject, MaterialSet, OptionNode, QuestionNode, Section


def _build_training_project() -> ExamProject:
    """Build a project with enough questions to generate diverse trajectories."""
    verbal_questions = []
    for i in range(1, 6):
        verbal_questions.append(QuestionNode(
            source_number=str(i),
            stem=f"这段文字意在说明第{i}题的主旨",
            options=[
                OptionNode(letter="A", text=f"选项A内容{i}"),
                OptionNode(letter="B", text=f"选项B内容{i}"),
                OptionNode(letter="C", text=f"选项C内容{i}"),
                OptionNode(letter="D", text=f"选项D内容{i}"),
            ],
        ))

    data_questions = []
    for i in range(101, 104):
        data_questions.append(QuestionNode(
            source_number=str(i),
            stem=f"根据上述资料，第{i}题下列说法正确的是",
            options=[
                OptionNode(letter="A", text=f"选项A内容{i}"),
                OptionNode(letter="B", text=f"选项B内容{i}"),
                OptionNode(letter="C", text=f"选项C内容{i}"),
                OptionNode(letter="D", text=f"选项D内容{i}"),
            ],
        ))

    return ExamProject(
        title="training_test",
        sections=[
            Section(kind="verbal", title="言语理解", questions=verbal_questions),
            Section(
                kind="data",
                title="资料分析",
                material_sets=[
                    MaterialSet(
                        material_id="m1",
                        header="材料一",
                        body="2023年全市GDP同比增长5.2%",
                        questions=data_questions,
                    ),
                ],
            ),
        ],
    )


class TrajectoryTrainingTest(unittest.TestCase):
    def test_train_trajectory_model_with_generated_data(self):
        try:
            from scripts.train_repair_trajectory_model import train_trajectory_model
        except (ImportError, SystemExit):
            self.skipTest("scikit-learn not installed, skipping training test")

        project = _build_training_project()
        trajectories = generate_trajectories_for_project(
            project, source_pdf="train_test.pdf", source_form="test",
        )

        self.assertGreater(len(trajectories), 3)

        with tempfile.TemporaryDirectory() as tmp:
            dataset_path = Path(tmp) / "trajectories.jsonl"
            model_path = Path(tmp) / "trajectory_policy.pkl"

            with dataset_path.open("w", encoding="utf-8") as fh:
                for traj in trajectories:
                    fh.write(json.dumps(traj, ensure_ascii=False) + "\n")

            summary = train_trajectory_model(dataset_path, model_path)

            self.assertTrue(model_path.exists())
            self.assertIn("macro_f1", summary)
            self.assertIn("trajectory_count", summary)
            self.assertEqual(summary["trajectory_count"], len(trajectories))
            self.assertIn("ready_for_runtime", summary)

            metrics_path = model_path.with_suffix(".metrics.json")
            self.assertTrue(metrics_path.exists())


if __name__ == "__main__":
    unittest.main()
