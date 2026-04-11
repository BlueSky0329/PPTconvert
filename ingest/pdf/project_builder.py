from __future__ import annotations

from collections import defaultdict
import os
import re
import shutil
import tempfile
from typing import Iterable, Mapping, Optional

from core.pdf_exam_extract import ExtractedImageRegion, extract_pdf_line_items_with_metadata
from core.pdf_exam_models import ObjectiveSection, ParsedExam, RichLine
from core.pdf_exam_parse import parse_line_items
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, PageRegion, PaperSource, QuestionNode, Section, SubjectKind
from ingest.pdf.layout import PageTextLine, extract_pdf_text_lines

_OPTION_PREFIX = re.compile(r"^\s*([A-Z])\s*[.．、:：\)）]\s*", re.IGNORECASE)


class _LayoutLocator:
    def __init__(self, lines: Iterable[PageTextLine]):
        self._lines = list(lines)
        self._index: dict[str, list[int]] = defaultdict(list)
        for idx, line in enumerate(self._lines):
            if line.text:
                self._index[line.text].append(idx)
        self._cursor = -1

    def consume(self, texts: Iterable[str]) -> list[PageTextLine]:
        matched: list[PageTextLine] = []
        for text in texts:
            key = (text or "").strip()
            if not key:
                continue
            candidates = self._index.get(key, [])
            chosen = None
            for idx in candidates:
                if idx > self._cursor:
                    chosen = idx
                    break
            if chosen is None and candidates:
                chosen = candidates[-1]
            if chosen is None:
                continue
            matched.append(self._lines[chosen])
            self._cursor = max(self._cursor, chosen)
        return matched


def _rich_line_text(rich: RichLine) -> str:
    return "".join(text for text, _img in rich.parts).strip()


def _rich_line_images(rich: RichLine) -> list[str]:
    return [img for _text, img in rich.parts if img]


def _needs_space_between(left: str, right: str) -> bool:
    if not left or not right:
        return False
    return left[-1].isascii() and left[-1].isalnum() and right[0].isascii() and right[0].isalnum()


def _join_wrapped_lines(lines: list[str]) -> str:
    merged = ""
    for line in (item.strip() for item in lines if item and item.strip()):
        if not merged:
            merged = line
            continue
        merged += (" " if _needs_space_between(merged, line) else "") + line
    return merged.strip()


def _regions_from_lines(lines: Iterable[PageTextLine]) -> list[PageRegion]:
    grouped: dict[int, list[tuple[float, float, float, float]]] = defaultdict(list)
    for line in lines:
        if None not in (line.block_x0, line.block_y0, line.block_x1, line.block_y1):
            bbox = (
                float(line.block_x0),
                float(line.block_y0),
                float(line.block_x1),
                float(line.block_y1),
            )
        else:
            bbox = (line.x0, line.y0, line.x1, line.y1)
        grouped[line.page_number].append(bbox)

    regions: list[PageRegion] = []
    for page_number, page_boxes in grouped.items():
        unique_boxes = list(dict.fromkeys(page_boxes))
        x0 = min(box[0] for box in unique_boxes)
        y0 = min(box[1] for box in unique_boxes)
        x1 = max(box[2] for box in unique_boxes)
        y1 = max(box[3] for box in unique_boxes)
        regions.append(PageRegion(page_number=page_number, x0=x0, y0=y0, x1=x1, y1=y1))
    regions.sort(key=lambda region: region.page_number)
    return regions


def _coalesce_regions(regions: Iterable[PageRegion]) -> list[PageRegion]:
    grouped: dict[int, list[PageRegion]] = defaultdict(list)
    for region in regions:
        grouped[region.page_number].append(region)

    merged: list[PageRegion] = []
    for page_number, page_regions in grouped.items():
        merged.append(
            PageRegion(
                page_number=page_number,
                x0=min(region.x0 for region in page_regions),
                y0=min(region.y0 for region in page_regions),
                x1=max(region.x1 for region in page_regions),
                y1=max(region.y1 for region in page_regions),
            )
        )
    merged.sort(key=lambda region: region.page_number)
    return merged


def _region_from_extracted_image(info: ExtractedImageRegion | None) -> PageRegion | None:
    if info is None:
        return None
    return PageRegion(
        page_number=info.page_number,
        x0=info.x0,
        y0=info.y0,
        x1=info.x1,
        y1=info.y1,
    )


def _copy_asset(path: str, asset_dir: str, seen: dict[str, str]) -> str:
    source = os.path.abspath(path)
    cached = seen.get(source)
    if cached:
        return cached

    name = os.path.basename(source)
    stem, ext = os.path.splitext(name)
    candidate = os.path.join(asset_dir, name)
    counter = 1
    while os.path.exists(candidate):
        candidate = os.path.join(asset_dir, f"{stem}_{counter}{ext}")
        counter += 1
    shutil.copy2(source, candidate)
    seen[source] = candidate
    return candidate


def _materialize_project_assets(project: ExamProject, asset_dir: str) -> None:
    os.makedirs(asset_dir, exist_ok=True)
    copied: dict[str, str] = {}
    project.source.asset_dir = asset_dir

    for _section, material, question in project.iter_questions():
        for asset in question.stem_assets:
            if asset.path and os.path.exists(asset.path):
                asset.path = _copy_asset(asset.path, asset_dir, copied)
        for option in question.options:
            if option.image_path and os.path.exists(option.image_path):
                option.image_path = _copy_asset(option.image_path, asset_dir, copied)
        if material:
            for asset in material.body_assets:
                if asset.path and os.path.exists(asset.path):
                    asset.path = _copy_asset(asset.path, asset_dir, copied)


def _question_from_rich(
    source_number: str,
    stem_lines: list[RichLine],
    option_lines: list[RichLine],
    locator: _LayoutLocator,
    image_regions: Mapping[str, ExtractedImageRegion] | None = None,
) -> QuestionNode:
    stem_text_lines = [_rich_line_text(line) for line in stem_lines if _rich_line_text(line)]
    located_lines = locator.consume(stem_text_lines)
    page_numbers = sorted({line.page_number for line in located_lines})

    stem_assets: list[AssetRef] = []
    for rich_line in stem_lines:
        for image_path in _rich_line_images(rich_line):
            region = _region_from_extracted_image((image_regions or {}).get(image_path))
            stem_assets.append(
                AssetRef(
                    kind="stem_image",
                    path=image_path,
                    source_page=region.page_number if region else None,
                    page_region=region,
                )
            )

    options: list[OptionNode] = []
    current: Optional[OptionNode] = None
    for rich_line in option_lines:
        text = _rich_line_text(rich_line)
        images = _rich_line_images(rich_line)
        if text:
            match = _OPTION_PREFIX.match(text)
            if match:
                if current is not None:
                    options.append(current)
                current = OptionNode(
                    letter=match.group(1).upper(),
                    text=text[match.end() :].strip(),
                )
            elif current is not None:
                current.text = _join_wrapped_lines([current.text, text])
        if current is not None and images and not current.image_path:
            current.image_path = images[0]
            region = _region_from_extracted_image((image_regions or {}).get(images[0]))
            current.source_page = region.page_number if region else None
            current.page_region = region
    if current is not None:
        options.append(current)

    stem = _join_wrapped_lines(stem_text_lines)
    return QuestionNode(
        source_number=(source_number or "").strip(),
        stem=stem,
        options=options,
        stem_assets=stem_assets,
        page_numbers=page_numbers,
    )


def _append_objective_project_section(
    project: ExamProject,
    parsed_section: ObjectiveSection,
    locator: _LayoutLocator,
    image_regions: Mapping[str, ExtractedImageRegion] | None = None,
) -> None:
    section = Section(kind=parsed_section.kind, title=parsed_section.title)
    for parsed_question in parsed_section.questions:
        section.questions.append(
            _question_from_rich(
                parsed_question.source_number,
                parsed_question.stem_lines,
                parsed_question.option_lines,
                locator,
                image_regions=image_regions,
            )
        )
    if section.questions:
        project.sections.append(section)


def build_project_from_parsed_exam(
    exam: ParsedExam,
    *,
    source_pdf_path: Optional[str] = None,
    layout_lines: Optional[Iterable[PageTextLine]] = None,
    image_regions: Mapping[str, ExtractedImageRegion] | None = None,
    title: Optional[str] = None,
) -> ExamProject:
    project_title = title or (
        os.path.splitext(os.path.basename(source_pdf_path))[0] if source_pdf_path else "Exam Project"
    )
    locator = _LayoutLocator(layout_lines or [])
    project = ExamProject(
        title=project_title,
        source=PaperSource(pdf_path=source_pdf_path),
    )

    for parsed_section in exam.iter_objective_sections():
        _append_objective_project_section(project, parsed_section, locator, image_regions=image_regions)

    for section_index, data_section in enumerate(exam.data_sections, 1):
        section = Section(kind="data", title=data_section.title)
        for material_index, parsed_material in enumerate(data_section.materials, 1):
            intro_lines = [_rich_line_text(line) for line in parsed_material.intro_lines if _rich_line_text(line)]
            intro_assets: list[AssetRef] = []
            intro_asset_regions: list[PageRegion] = []
            for rich_line in parsed_material.intro_lines:
                for image_path in _rich_line_images(rich_line):
                    region = _region_from_extracted_image((image_regions or {}).get(image_path))
                    if region is not None:
                        intro_asset_regions.append(region)
                    intro_assets.append(
                        AssetRef(
                            kind="material_inline_image",
                            path=image_path,
                            source_page=region.page_number if region else None,
                            page_region=region,
                        )
                    )

            located_lines = locator.consume([parsed_material.header] + intro_lines)
            material = MaterialSet(
                material_id=f"data-{section_index}-{material_index}",
                header=parsed_material.header.strip(),
                body="\n".join(intro_lines).strip(),
                body_lines=intro_lines,
                body_assets=intro_assets,
                body_regions=_coalesce_regions([*_regions_from_lines(located_lines), *intro_asset_regions]),
            )
            for parsed_question in parsed_material.questions:
                material.questions.append(
                    _question_from_rich(
                        parsed_question.source_number,
                        parsed_question.stem_lines,
                        parsed_question.option_lines,
                        locator,
                        image_regions=image_regions,
                    )
                )
            if material.questions:
                section.material_sets.append(material)
        if section.material_sets:
            project.sections.append(section)

    return project


def build_exam_project_from_pdf(
    pdf_path: str,
    *,
    mode: str = "all",
    asset_dir: Optional[str] = None,
    document_subject_hint: SubjectKind | None = None,
) -> ExamProject:
    items, temp_dir, image_regions = extract_pdf_line_items_with_metadata(pdf_path)
    try:
        exam = parse_line_items(
            items,
            mode=mode,
            document_subject_hint=document_subject_hint,
        )  # type: ignore[arg-type]
        layout_lines = extract_pdf_text_lines(pdf_path)
        project = build_project_from_parsed_exam(
            exam,
            source_pdf_path=pdf_path,
            layout_lines=layout_lines,
            image_regions=image_regions,
        )
        target_asset_dir = asset_dir or tempfile.mkdtemp(prefix="pptconvert_project_assets_")
        _materialize_project_assets(project, target_asset_dir)
        return project
    finally:
        shutil.rmtree(temp_dir, ignore_errors=True)
