from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Option:
    """单个选项"""
    letter: str          # A / B / C / D
    text: str            # 选项文本内容
    image_path: Optional[str] = None


@dataclass
class Question:
    """一道选择题"""
    number: int                              # 题号
    stem: str                                # 题干文本
    options: list[Option] = field(default_factory=list)
    image_paths: list[str] = field(default_factory=list)  # 题目关联图片
    answer: Optional[str] = None             # 正确答案（可选）

    @property
    def is_complete(self) -> bool:
        return bool(self.stem and len(self.options) >= 2)

    def get_option_text(self, letter: str) -> Optional[str]:
        for opt in self.options:
            if opt.letter.upper() == letter.upper():
                return opt.text
        return None
