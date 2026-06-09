# -*- coding: utf-8 -*-
import unittest

from core.subject_inference import infer_subject_diagnostics


class PoliticsAnchorClassificationTest(unittest.TestCase):
    """强政治锚点（习近平/党的二十大/中央文件/马克思恩格斯）应被判为 politics，
    且不能因此把法律常识题误拉成 politics。"""

    def _kind(self, stem, options):
        return infer_subject_diagnostics(stem=stem, options=list(options), allow_data=True).kind

    def test_strong_anchor_questions_are_politics(self):
        cases = [
            "党的二十大报告指出，全面依法治国是国家治理的一场深刻革命。下列说法正确的是：",
            "习近平总书记在中央城市工作会议上强调，做好城市工作的总体要求。下列表述正确的是：",
            "恩格斯指出，马克思的整个世界观不是教义而是方法。下列理解正确的是：",
            "党的二十届三中全会提出，进一步全面深化改革、推进中国式现代化。下列表述正确的是：",
        ]
        options = ["甲表述内容", "乙表述内容", "丙表述内容", "丁表述内容"]
        for stem in cases:
            self.assertEqual(self._kind(stem, options), "politics", stem[:24])

    def test_law_common_sense_is_not_pulled_into_politics(self):
        kind = self._kind(
            "根据我国《民法典》的相关规定，下列关于合同效力的说法正确的是：",
            ["甲说法内容", "乙说法内容", "丙说法内容", "丁说法内容"],
        )
        self.assertNotEqual(kind, "politics")


if __name__ == "__main__":
    unittest.main()
