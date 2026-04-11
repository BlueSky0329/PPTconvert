import unittest

from domain.models import ExamProject, QuestionNode, Section
from domain.selectors import parse_subject_spec, select_project


class SelectorTest(unittest.TestCase):
    def test_select_project_preserves_unknown_sections_under_subject_filter(self):
        project = ExamProject(
            title="demo",
            sections=[
                Section(kind="quant", title="数量关系", questions=[QuestionNode(source_number="66", stem="q1")]),
                Section(kind="unknown", title="待确认科目", questions=[QuestionNode(source_number="67", stem="q2")]),
            ],
        )

        filtered = select_project(project, subjects=parse_subject_spec("quant"))
        self.assertEqual([section.kind for section in filtered.sections], ["quant", "unknown"])


if __name__ == "__main__":
    unittest.main()
