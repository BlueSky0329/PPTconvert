from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Optional

SubjectKind = Literal[
    "politics",
    "common_sense",
    "verbal",
    "quant",
    "reasoning",
    "data",
    "unknown",
]

ALL_SUBJECT_KINDS: tuple[SubjectKind, ...] = (
    "politics",
    "common_sense",
    "verbal",
    "quant",
    "reasoning",
    "data",
)

SUBJECT_DISPLAY_NAMES: dict[SubjectKind, str] = {
    "politics": "政治理论",
    "common_sense": "常识判断",
    "verbal": "言语理解与表达",
    "quant": "数量关系",
    "reasoning": "判断推理",
    "data": "资料分析",
    "unknown": "未知科目",
}


@dataclass
class PageRegion:
    page_number: int
    x0: float
    y0: float
    x1: float
    y1: float

    def padded(self, margin: float = 8.0) -> "PageRegion":
        return PageRegion(
            page_number=self.page_number,
            x0=max(0.0, self.x0 - margin),
            y0=max(0.0, self.y0 - margin),
            x1=self.x1 + margin,
            y1=self.y1 + margin,
        )


@dataclass
class AssetRef:
    kind: str
    path: str
    source_page: Optional[int] = None
    page_region: Optional[PageRegion] = None
    label: str = ""


@dataclass
class OptionNode:
    letter: str
    text: str
    image_path: Optional[str] = None
    source_page: Optional[int] = None
    page_region: Optional[PageRegion] = None


@dataclass
class QuestionNode:
    source_number: str
    stem: str
    options: list[OptionNode] = field(default_factory=list)
    stem_assets: list[AssetRef] = field(default_factory=list)
    answer: Optional[str] = None
    page_numbers: list[int] = field(default_factory=list)
    option_layout: Optional[str] = None

    @property
    def numeric_source_number(self) -> Optional[int]:
        value = (self.source_number or "").strip()
        if value.isdigit():
            return int(value)
        return None


@dataclass
class MaterialSet:
    material_id: str
    header: str
    body: str
    body_lines: list[str] = field(default_factory=list)
    body_assets: list[AssetRef] = field(default_factory=list)
    body_regions: list[PageRegion] = field(default_factory=list)
    questions: list[QuestionNode] = field(default_factory=list)


@dataclass
class Section:
    kind: SubjectKind
    title: str
    questions: list[QuestionNode] = field(default_factory=list)
    material_sets: list[MaterialSet] = field(default_factory=list)


@dataclass
class PaperSource:
    pdf_path: Optional[str] = None
    docx_path: Optional[str] = None
    asset_dir: Optional[str] = None


@dataclass
class QuestionRange:
    start: int
    end: int

    def contains(self, value: int) -> bool:
        return self.start <= value <= self.end


@dataclass
class ExamProject:
    title: str
    source: PaperSource = field(default_factory=PaperSource)
    sections: list[Section] = field(default_factory=list)
    selected_subjects: list[SubjectKind] = field(default_factory=list)
    selected_ranges: list[QuestionRange] = field(default_factory=list)

    def iter_questions(self):
        for section in self.sections:
            if section.kind == "data":
                for material in section.material_sets:
                    for question in material.questions:
                        yield section, material, question
            else:
                for question in section.questions:
                    yield section, None, question

    @property
    def question_count(self) -> int:
        return sum(1 for _section, _material, _question in self.iter_questions())
