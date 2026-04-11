"""PDF 真题整理用的中间结构（资料分析 / 数量关系）。"""

from __future__ import annotations

from dataclasses import dataclass, field

from domain.models import SubjectKind


@dataclass
class RichLine:
    """一行内容：纯文本和/或内嵌图片路径。"""

    parts: list[tuple[str, str | None]] = field(default_factory=list)
    # (text, image_path_or_none); image_path 为临时文件路径


@dataclass
class ExamOption:
    letter: str
    text: str


@dataclass
class ExamQuestion:
    stem_lines: list[RichLine]
    option_lines: list[RichLine]
    source_number: str = ""


@dataclass
class MaterialUnit:
    header: str  # 如 材料一
    intro_lines: list[RichLine]  # 材料正文（可含图）
    questions: list[ExamQuestion]


@dataclass
class ObjectiveSection:
    title: str
    questions: list[ExamQuestion]
    kind: SubjectKind = "unknown"


@dataclass
class PoliticsSection(ObjectiveSection):
    kind: SubjectKind = "politics"


@dataclass
class CommonSenseSection(ObjectiveSection):
    kind: SubjectKind = "common_sense"


@dataclass
class VerbalSection(ObjectiveSection):
    kind: SubjectKind = "verbal"


@dataclass
class ReasoningSection(ObjectiveSection):
    kind: SubjectKind = "reasoning"


@dataclass
class UnknownSection(ObjectiveSection):
    kind: SubjectKind = "unknown"


@dataclass
class QuantSection(ObjectiveSection):
    kind: SubjectKind = "quant"


@dataclass
class DataAnalysisSection:
    title: str  # 如 2026年·天津·资料分析
    materials: list[MaterialUnit]
    kind: SubjectKind = "data"


@dataclass
class ParsedExam:
    politics_sections: list[PoliticsSection] = field(default_factory=list)
    common_sense_sections: list[CommonSenseSection] = field(default_factory=list)
    verbal_sections: list[VerbalSection] = field(default_factory=list)
    data_sections: list[DataAnalysisSection] = field(default_factory=list)
    quant_sections: list[QuantSection] = field(default_factory=list)
    reasoning_sections: list[ReasoningSection] = field(default_factory=list)
    unknown_sections: list[UnknownSection] = field(default_factory=list)

    def iter_objective_sections(self) -> list[ObjectiveSection]:
        return [
            *self.politics_sections,
            *self.common_sense_sections,
            *self.verbal_sections,
            *self.quant_sections,
            *self.reasoning_sections,
            *self.unknown_sections,
        ]
