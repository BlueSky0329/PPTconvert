import unittest

from core.project_quality import annotate_project_quality
from domain.models import ExamProject, OptionNode, QuestionNode, Section
from gui.app import PPTConvertApp


class GuiPreviewBehaviorTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.app = PPTConvertApp()
        cls.app.root.withdraw()

    @classmethod
    def tearDownClass(cls):
        try:
            cls.app.root.update_idletasks()
            cls.app.root.update()
        except Exception:
            pass
        cls.app.root.destroy()

    def _current_left_tab(self) -> str:
        selected = self.app._pdf_preview_left_tabs.select()
        if selected == str(self.app._pdf_preview_structure_tab):
            return "structure"
        if selected == str(self.app._pdf_preview_review_tab):
            return "review"
        if selected == str(self.app._pdf_preview_slide_tab):
            return "slide"
        return selected

    def _build_flagged_project(self) -> ExamProject:
        project = ExamProject(
            title="示例工程",
            sections=[
                Section(
                    kind="common_sense",
                    title="常识判断",
                    questions=[
                        QuestionNode(
                            source_number="1",
                            stem="根据上述定义，下列符合定义的是",
                            options=[
                                OptionNode("A", "甲"),
                                OptionNode("B", "乙"),
                                OptionNode("C", "丙"),
                                OptionNode("D", "丁"),
                            ],
                        ),
                        QuestionNode(
                            source_number="2",
                            stem="下列关于宪法的说法正确的是",
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
        annotate_project_quality(project)
        return project

    def test_structure_question_selection_keeps_structure_tab_active(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        section_item = self.app.pdf_tree.get_children()[0]
        question_item = next(
            child
            for child in self.app.pdf_tree.get_children(section_item)
            if self.app._pdf_preview_payloads.get(child, {}).get("kind") == "question"
        )

        self.app.pdf_tree.selection_set(question_item)
        self.app.pdf_tree.focus(question_item)
        self.app._on_pdf_preview_select()
        self.app.root.update_idletasks()
        self.app.root.update()

        self.assertEqual(self._current_left_tab(), "structure")
        self.assertEqual(self.app._selected_pdf_item_id(), question_item)

    def test_ai_suggestion_mentions_explicit_subject_name(self):
        project = self._build_flagged_project()
        self.app._populate_pdf_preview(project)
        self.app.root.update_idletasks()
        self.app.root.update()

        review_item = self.app._pdf_review_tree.get_children()[0]
        self.app._pdf_review_tree.selection_set(review_item)
        self.app._pdf_review_tree.focus(review_item)
        self.app._on_pdf_review_select()
        self.app.root.update_idletasks()
        self.app.root.update()

        suggestion = self.app._pdf_ai_suggestion_var.get()
        self.assertIn("判断推理", suggestion)


if __name__ == "__main__":
    unittest.main()
