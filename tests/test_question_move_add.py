# -*- coding: utf-8 -*-
import unittest

from domain.models import ExamProject, OptionNode, QuestionNode, Section
from domain.project_editor import insert_blank_question_after, move_question_to_kind


def _q(number):
    return QuestionNode(
        source_number=str(number),
        stem=f"题{number}",
        options=[OptionNode(letter=l, text=t) for l, t in zip("ABCD", "abcd")],
    )


class InsertBlankQuestionTest(unittest.TestCase):
    def test_insert_blank_after(self):
        q1, q2 = _q(1), _q(2)
        proj = ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=[q1, q2])])
        new_q = insert_blank_question_after(proj, q1)
        self.assertIsNotNone(new_q)
        self.assertEqual(len(new_q.options), 4)
        self.assertEqual([o.letter for o in new_q.options], ["A", "B", "C", "D"])
        order = [qq for _s, _m, qq in proj.iter_questions()]
        self.assertEqual(order, [q1, new_q, q2])


class MoveQuestionToKindTest(unittest.TestCase):
    def test_move_to_existing_section(self):
        qv, qc = _q(30), _q(31)
        verbal = Section(kind="verbal", title="言语", questions=[qv])
        common = Section(kind="common_sense", title="常识", questions=[qc])
        proj = ExamProject(title="t", sections=[verbal, common])
        self.assertTrue(move_question_to_kind(proj, qv, "common_sense"))
        self.assertIn(qv, common.questions)
        self.assertNotIn("verbal", {s.kind for s in proj.sections})  # emptied + cleaned

    def test_move_creates_section_when_missing(self):
        qv = _q(30)
        proj = ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=[qv, _q(31)])])
        self.assertTrue(move_question_to_kind(proj, qv, "common_sense"))
        common_sections = [s for s in proj.sections if s.kind == "common_sense"]
        self.assertEqual(len(common_sections), 1)
        self.assertIn(qv, common_sections[0].questions)

    def test_move_rejects_data_and_same_kind(self):
        qv = _q(30)
        proj = ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=[qv])])
        self.assertFalse(move_question_to_kind(proj, qv, "data"))
        self.assertFalse(move_question_to_kind(proj, qv, "verbal"))


if __name__ == "__main__":
    unittest.main()
