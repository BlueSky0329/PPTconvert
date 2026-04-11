import unittest

from core.models import Option, Question
from ingest.docx.project_builder import build_exam_project_from_word_questions


class WordProjectBuilderTest(unittest.TestCase):
    def test_build_project_groups_sections_and_data_materials(self):
        questions = [
            Question(
                number=1,
                stem="言语题",
                options=[Option("A", "甲"), Option("B", "乙")],
                source_question_number="36",
                section_kind="verbal",
                section_title="三. 言语理解与表达",
            ),
            Question(
                number=2,
                stem="资料题一",
                options=[Option("A", "1"), Option("B", "2")],
                source_question_number="116",
                section_kind="data",
                section_title="六. 资料分析",
                material_header="材料一",
                material_text="表格一",
                material_image_paths=["m1.png"],
                question_image_paths=["q1.png"],
                image_paths=["m1.png", "q1.png"],
            ),
            Question(
                number=3,
                stem="资料题二",
                options=[Option("A", "3"), Option("B", "4")],
                source_question_number="117",
                section_kind="data",
                section_title="六. 资料分析",
                material_header="材料一",
                material_text="表格一",
                material_image_paths=["m1.png"],
            ),
            Question(
                number=4,
                stem="资料题三",
                options=[Option("A", "5"), Option("B", "6")],
                source_question_number="121",
                section_kind="data",
                section_title="六. 资料分析",
                material_header="材料二",
                material_text="表格二",
                material_image_paths=["m2.png"],
            ),
        ]

        project = build_exam_project_from_word_questions(
            questions,
            title="示例题库",
            docx_path="sample.docx",
            asset_dir="sample_assets",
        )

        self.assertEqual(project.title, "示例题库")
        self.assertEqual(project.source.docx_path, "sample.docx")
        self.assertEqual(project.source.asset_dir, "sample_assets")
        self.assertEqual([section.kind for section in project.sections], ["verbal", "data"])

        verbal_section = project.sections[0]
        self.assertEqual(len(verbal_section.questions), 1)
        self.assertEqual(verbal_section.questions[0].source_number, "36")

        data_section = project.sections[1]
        self.assertEqual(len(data_section.material_sets), 2)
        self.assertEqual(data_section.material_sets[0].header, "材料一")
        self.assertEqual(len(data_section.material_sets[0].questions), 2)
        self.assertEqual(data_section.material_sets[0].body_assets[0].path, "m1.png")
        self.assertEqual(data_section.material_sets[0].questions[0].stem_assets[0].path, "q1.png")
        self.assertEqual(data_section.material_sets[1].header, "材料二")
        self.assertEqual(data_section.material_sets[1].body_assets[0].path, "m2.png")


if __name__ == "__main__":
    unittest.main()
