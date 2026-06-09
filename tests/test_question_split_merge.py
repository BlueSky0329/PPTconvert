# -*- coding: utf-8 -*-
import unittest

from domain.models import ExamProject, OptionNode, QuestionNode, Section
from domain.project_editor import merge_with_next_question, split_question_options


def _q(number, stem, option_texts, answer=None):
    return QuestionNode(
        source_number=str(number),
        stem=stem,
        options=[OptionNode(letter=l, text=t) for l, t in zip("ABCDEFGH", option_texts)],
        answer=answer,
    )


class SplitQuestionTest(unittest.TestCase):
    def _proj(self, questions):
        return ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=list(questions))])

    def test_split_options_into_new_question(self):
        q = _q(5, "题干一", ["a", "b", "c", "d", "e", "f", "g", "h"], answer="B")
        proj = self._proj([q])
        new_q = split_question_options(proj, q, 4, new_number="6", new_stem="题干二")
        self.assertIsNotNone(new_q)
        self.assertEqual([o.letter for o in q.options], ["A", "B", "C", "D"])
        self.assertEqual([o.text for o in q.options], ["a", "b", "c", "d"])
        self.assertEqual([o.text for o in new_q.options], ["e", "f", "g", "h"])
        self.assertEqual([o.letter for o in new_q.options], ["A", "B", "C", "D"])
        self.assertEqual(q.answer, "B")
        self.assertEqual([qq.source_number for _s, _m, qq in proj.iter_questions()], ["5", "6"])

    def test_split_moves_answer_to_new_question(self):
        q = _q(5, "题干", ["a", "b", "c", "d", "e", "f"], answer="E")
        proj = self._proj([q])
        new_q = split_question_options(proj, q, 4)
        self.assertEqual(new_q.options[0].text, "e")
        self.assertEqual(new_q.answer, "A")
        self.assertIsNone(q.answer)

    def test_split_invalid_index_returns_none(self):
        q = _q(1, "s", ["a", "b", "c", "d"])
        proj = self._proj([q])
        self.assertIsNone(split_question_options(proj, q, 0))
        self.assertIsNone(split_question_options(proj, q, 4))


class MergeQuestionTest(unittest.TestCase):
    def _proj(self, questions):
        return ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=list(questions))])

    def test_merge_with_next_continuation(self):
        q1 = _q(5, "题干前半", ["a", "b"], answer="A")
        q2 = _q(6, "", ["c", "d"])
        proj = self._proj([q1, q2])
        self.assertTrue(merge_with_next_question(proj, q1))
        self.assertEqual([o.letter for o in q1.options], ["A", "B", "C", "D"])
        self.assertEqual([o.text for o in q1.options], ["a", "b", "c", "d"])
        self.assertEqual(q1.answer, "A")
        self.assertEqual(proj.question_count, 1)

    def test_merge_appends_next_stem(self):
        q1 = _q(5, "前", ["a", "b"])
        q2 = _q(6, "后", ["c", "d"])
        proj = self._proj([q1, q2])
        merge_with_next_question(proj, q1)
        self.assertIn("前", q1.stem)
        self.assertIn("后", q1.stem)

    def test_merge_no_next_returns_false(self):
        q = _q(1, "s", ["a", "b", "c", "d"])
        proj = self._proj([q])
        self.assertFalse(merge_with_next_question(proj, q))


if __name__ == "__main__":
    unittest.main()
