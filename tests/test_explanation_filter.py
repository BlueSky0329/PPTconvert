# -*- coding: utf-8 -*-
import unittest

from core.explanation_filter import (
    classify_explanation_content,
    filter_explanation_questions,
    is_explanation_question,
    is_real_question,
)
from domain.models import ExamProject, OptionNode, QuestionNode, Section


def _q(number, stem, option_texts):
    return QuestionNode(
        source_number=str(number),
        stem=stem,
        options=[OptionNode(letter=letter, text=text) for letter, text in zip("ABCD", option_texts)],
    )


def _real(number):
    return _q(number, f"下列说法正确的是第{number}题：", ["选项甲", "选项乙", "选项丙", "选项丁"])


def _explanation_stem(number):
    return _q(
        number,
        "本题考查政治常识。",
        ["A项正确，报告指出……", "B项错误，规定：……", "C项错误，表述错误", "D项错误，对应不符"],
    )


def _explanation_answer(number):
    return QuestionNode(source_number=str(number), stem=".【答案】A【解析】先看第一空……", options=[])


class ExplanationClassifierTest(unittest.TestCase):
    def test_real_vs_explanation_questions(self):
        self.assertTrue(is_real_question(_real(1)))
        self.assertFalse(is_real_question(_explanation_stem(1)))
        self.assertFalse(is_real_question(_explanation_answer(1)))
        self.assertTrue(is_explanation_question(_explanation_stem(1)))
        self.assertTrue(is_explanation_question(_explanation_answer(1)))
        self.assertFalse(is_explanation_question(_real(1)))


class ExplanationFilterTest(unittest.TestCase):
    def _project(self, questions, kind="politics", title="政治"):
        return ExamProject(title="t", sections=[Section(kind=kind, title=title, questions=list(questions))])

    def test_clean_project_is_untouched(self):
        project = self._project([_real(1), _real(2), _real(3)], kind="common_sense", title="常识")
        info = filter_explanation_questions(project)
        self.assertEqual(info["category"], "clean")
        self.assertEqual(info["removed_questions"], 0)
        self.assertFalse(info["is_answer_booklet"])
        self.assertEqual(project.question_count, 3)
        self.assertEqual(project.import_notices, [])

    def test_combined_drops_explanations_keeps_real(self):
        questions = [_real(i) for i in range(1, 21)] + [_explanation_answer(i) for i in range(1, 11)]
        project = self._project(questions)
        info = filter_explanation_questions(project)
        self.assertEqual(info["category"], "combined")
        self.assertEqual(info["removed_questions"], 10)
        self.assertEqual(project.question_count, 20)
        self.assertTrue(
            all(not is_explanation_question(q) for _s, _m, q in project.iter_questions())
        )
        self.assertTrue(any("合订版" in note for note in project.import_notices))

    def test_answer_booklet_is_detected_and_flagged(self):
        questions = [_explanation_stem(i) for i in range(1, 13)] + [_real(99)]
        project = self._project(questions, kind="unknown", title="")
        info = classify_explanation_content(project)
        self.assertEqual(info["category"], "answer_booklet")
        info = filter_explanation_questions(project)
        self.assertTrue(info["is_answer_booklet"])
        self.assertLessEqual(project.question_count, 2)
        self.assertTrue(any(note.lstrip().startswith("⚠") for note in project.import_notices))


if __name__ == "__main__":
    unittest.main()
