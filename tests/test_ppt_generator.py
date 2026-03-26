import unittest

from pptx.dml.color import RGBColor
from pptx.util import Inches, Pt

from core.models import Option, Question
from core.ppt_generator import PPTConfig, PPTGenerator


class PPTConfigTest(unittest.TestCase):
    def test_from_mapping_applies_overrides(self):
        config = PPTConfig.from_mapping(
            {
                "margin_left_in": 1.2,
                "stem_font_size": Pt(24),
                "image_max_width": Inches(4),
                "option_letter_color": RGBColor(1, 2, 3),
            }
        )
        self.assertEqual(config.margin_left_in, 1.2)
        self.assertEqual(config.stem_font_size.pt, 24)
        self.assertEqual(config.image_max_width.inches, 4)
        self.assertEqual(config.option_letter_color, RGBColor(1, 2, 3))
        self.assertEqual(config.number_color, RGBColor(1, 2, 3))


class PPTGeneratorTest(unittest.TestCase):
    def test_renderable_options_warns_when_question_has_more_than_four_options(self):
        question = Question(
            number=1,
            stem="stem",
            options=[
                Option("A", "a"),
                Option("B", "b"),
                Option("C", "c"),
                Option("D", "d"),
                Option("E", "e"),
            ],
        )
        generator = PPTGenerator()

        with self.assertLogs("core.ppt_generator", level="WARNING") as captured:
            options = generator._renderable_options(question)

        self.assertEqual([option.letter for option in options], ["A", "B", "C", "D"])
        self.assertIn("only the first 4 can be rendered", captured.output[0])

