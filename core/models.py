from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Option:
    letter: str
    text: str
    image_path: Optional[str] = None


@dataclass
class Question:
    number: int
    stem: str
    options: list[Option] = field(default_factory=list)
    image_paths: list[str] = field(default_factory=list)
    question_image_paths: list[str] = field(default_factory=list)
    material_image_paths: list[str] = field(default_factory=list)
    answer: Optional[str] = None
    source_label: Optional[str] = None
    source_question_number: Optional[str] = None
    material_header: Optional[str] = None
    material_text: Optional[str] = None
    option_layout: Optional[str] = None
    section_kind: Optional[str] = None
    section_title: Optional[str] = None

    @property
    def display_stem(self) -> str:
        parts = [self.source_label, self.stem]
        return " ".join(part for part in parts if part).strip()

    @property
    def is_complete(self) -> bool:
        return bool(self.display_stem and len(self.options) >= 2)

    def get_option_text(self, letter: str) -> Optional[str]:
        for option in self.options:
            if option.letter.upper() == letter.upper():
                return option.text
        return None
