# -*- coding: utf-8 -*-
import unittest

from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section
from domain.project_editor import (
    move_question_in_container,
    reassign_stem_assets_to_options,
    set_material_body,
    set_question_answer,
)


def _q(number, options=("a", "b", "c", "d")):
    return QuestionNode(
        source_number=str(number),
        stem=f"题{number}",
        options=[OptionNode(letter=l, text=t) for l, t in zip("ABCD", options)],
    )


class ReorderQuestionTest(unittest.TestCase):
    def test_move_up_and_down(self):
        q1, q2, q3 = _q(1), _q(2), _q(3)
        proj = ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=[q1, q2, q3])])
        self.assertTrue(move_question_in_container(proj, q3, -1))  # q3 上移
        self.assertEqual([qq for _s, _m, qq in proj.iter_questions()], [q1, q3, q2])
        self.assertTrue(move_question_in_container(proj, q1, 1))  # q1 下移
        self.assertEqual([qq for _s, _m, qq in proj.iter_questions()], [q3, q1, q2])

    def test_move_at_boundary_returns_false(self):
        q1, q2 = _q(1), _q(2)
        proj = ExamProject(title="t", sections=[Section(kind="verbal", title="言语", questions=[q1, q2])])
        self.assertFalse(move_question_in_container(proj, q1, -1))
        self.assertFalse(move_question_in_container(proj, q2, 1))


class SetAnswerTest(unittest.TestCase):
    def test_set_and_normalize_answer(self):
        q = _q(1)
        set_question_answer(q, "b")
        self.assertEqual(q.answer, "B")
        set_question_answer(q, "a, b, b")
        self.assertEqual(q.answer, "AB")
        set_question_answer(q, "")
        self.assertIsNone(q.answer)


class ReassignStemToOptionsTest(unittest.TestCase):
    def test_reassign_stem_assets_to_options(self):
        q = _q(1, options=("", "", "", ""))
        q.stem_assets = [AssetRef(kind="image", path=f"/img{i}.png", source_page=1) for i in range(4)]
        changed = reassign_stem_assets_to_options(q)
        self.assertEqual(changed, 4)
        self.assertEqual([o.image_path for o in q.options], ["/img0.png", "/img1.png", "/img2.png", "/img3.png"])
        self.assertEqual(q.stem_assets, [])


class SetMaterialBodyTest(unittest.TestCase):
    def test_replace_material_body(self):
        material = MaterialSet(material_id="m1", header="材料一", body="旧正文", body_lines=["旧正文"])
        set_material_body(material, "新正文第一行\n新正文第二行")
        self.assertIn("新正文第一行", material.body)
        self.assertIn("新正文第二行", material.body)
        self.assertNotIn("旧正文", material.body)
        self.assertEqual(material.body_lines, ["新正文第一行", "新正文第二行"])


if __name__ == "__main__":
    unittest.main()
