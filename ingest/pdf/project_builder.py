from __future__ import annotations

from collections import defaultdict
import os
import re
import shutil
import tempfile
import unicodedata
from typing import Iterable, Mapping, Optional

from PIL import Image, ImageFilter

from core import pdf_exam_extract as pdf_extract
from core.pdf_exam_extract import ExtractedImageRegion, extract_pdf_line_items_with_metadata
from core.pdf_exam_models import ObjectiveSection, ParsedExam, RichLine
from core.pdf_exam_parse import parse_line_items
from core.project_quality import annotate_project_quality
from core.subject_inference import preferred_subject_title, should_merge_subject_sections
from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, PageRegion, PaperSource, QuestionNode, Section, SubjectKind
from ingest.pdf.layout import PageTextLine, extract_pdf_text_lines

_OPTION_PREFIX = re.compile(r"^\s*([ABCD])\s*[.．、:：\)）]\s*", re.IGNORECASE)
_OPTION_MARKER = re.compile(r"(?<![A-Za-z])([ABCD])\s*[.\uFF0E\u3001)\uFF09:：]\s*", re.IGNORECASE)
_INLINE_OPTION_MARKER = re.compile(r"(?<![A-Za-z])([ABCD])\s*[.\uFF0E\u3001)\uFF09:：\-－—–]\s*", re.IGNORECASE)
_INLINE_BARE_OPTION_MARKER = re.compile(
    r"(?<![A-Za-z])([ABCD])(?:\s*[.\uFF0E\u3001)\uFF09:：\-－—–]\s*|\s+)(?=\S)",
)
_TAIL_OPTION_LINE = re.compile(r"^\s*([ABCD])(?:\s*[.\uFF0E\u3001)\uFF09:：\-－—–]\s*|\s+)(.*)$", re.IGNORECASE)
_ENUMERATED_STEM_LINE = re.compile(r"^(?:[1-9](?!\d)[\u4e00-\u9fffA-Za-z]|[①②③④⑤⑥⑦⑧⑨⑩])")
_LEADING_BLANK_PUNCT = re.compile(r"^[,，、:：;；]")
_PROMPT_AFTER_ENUMERATION = re.compile(r"^(?:将以上|将下列|下列|根据|依次|这段|作者|最适合|填入|与原文|以下|对此)")
_TRAILING_SHORT_NUMBER = re.compile(r"(\d{1,2})$")
_LEADING_SHORT_NUMBER = re.compile(r"^(\d{1,2})(.*)$")
_POINT_DIAGRAM_PROMPT = re.compile(r"(?:下图|图中).*(?:位于|哪个|哪一).*(?:点|位置)", re.IGNORECASE)
_OPTION_LEFT_CONTEXT_CHARS = set(":：?？(（[【")
_OPTION_TAIL_PASSAGE_MARKER = re.compile(
    r"(?:以下是[^。\n]{0,80}?阅读之后回答\d{1,4}[—\-]\d{1,4}\s*题|阅读之后回答\d{1,4}[—\-]\d{1,4}\s*题)"
)
_OPTION_FOLLOWING_PASSAGE_LINE = re.compile(
    r"^\s*(?:以下是[^。\n]{0,80}?阅读之后回答\d{1,4}[—\-]\d{1,4}\s*题|阅读之后回答\d{1,4}[—\-]\d{1,4}\s*题)"
)
_TWO_COLUMN_OPTION_SPLIT = re.compile(r"\s{2,}")
_COMBINATION_OPTION_BODY = re.compile(r"^[0-9①②③④⑤⑥⑦⑧⑨⑩ⅠⅡⅢⅣⅤⅥⅦⅧⅨⅩIVXivx]+$")


def _option_marker_has_valid_left_context(text: str, start: int) -> bool:
    if start <= 0:
        return True
    previous = text[start - 1]
    return previous.isspace() or previous in _OPTION_LEFT_CONTEXT_CHARS


def _match_looks_like_embedded_list_marker(text: str, match: re.Match[str]) -> bool:
    if _option_marker_has_valid_left_context(text, match.start()):
        return False
    token = text[match.start() : match.end()]
    return "、" in token


def _looks_like_following_passage_line(text: str) -> bool:
    normalized = (text or "").strip()
    if not normalized:
        return False
    return bool(_OPTION_FOLLOWING_PASSAGE_LINE.match(normalized))


def _looks_like_combination_option_body(text: str) -> bool:
    normalized = re.sub(r"\s+", "", (text or "").strip())
    if not normalized:
        return False
    return bool(_COMBINATION_OPTION_BODY.fullmatch(normalized))


def _select_ordered_option_matches(matches: list[re.Match[str]]) -> list[re.Match[str]]:
    ordered: list[re.Match[str]] = []
    expected_letter: str | None = None
    for match in matches:
        letter = match.group(1).upper()
        if letter not in {"A", "B", "C", "D"}:
            continue
        if not ordered:
            ordered.append(match)
            expected_letter = chr(ord(letter) + 1) if letter < "D" else None
            continue
        if expected_letter is None or letter != expected_letter:
            continue
        ordered.append(match)
        expected_letter = chr(ord(letter) + 1) if letter < "D" else None
    return ordered


def _best_ordered_option_match_sequence(text: str, matches: list[re.Match[str]]) -> list[re.Match[str]]:
    filtered = [match for match in matches if match.group(1).upper() in {"A", "B", "C", "D"}]
    best: list[re.Match[str]] = []
    best_rank = (-1, -1, -1, -10**9)

    def rank(chosen: list[re.Match[str]]) -> tuple[int, int, int, int]:
        valid_left_context = sum(
            1 for match in chosen if _option_marker_has_valid_left_context(text, match.start())
        )
        if not chosen:
            return (0, valid_left_context, 0, 0)
        embedded_marker_penalty = 0
        short_body_penalty = 0
        leading_punct_penalty = 0
        body_length_score = 0
        bodies: list[str] = []
        for index, match in enumerate(chosen):
            body_start = match.end()
            body_end = chosen[index + 1].start() if index + 1 < len(chosen) else len(text)
            body = text[body_start:body_end].strip()
            bodies.append(body)
            body_length_score += min(len(body), 40)
            if body and len(body) <= 3:
                short_body_penalty += 1
            if body[:1] in ".,，。:：;；)）]】":
                leading_punct_penalty += 1
            embedded_marker_penalty += sum(
                1
                for marker in _OPTION_MARKER.finditer(body)
                if marker.group(1).upper() in {"A", "B", "C", "D"}
            )
        last_body = bodies[-1] if bodies else ""
        last_body_clean = 0 if any(
            marker.group(1).upper() in {"A", "B", "C", "D"} for marker in _OPTION_MARKER.finditer(last_body)
        ) else 1
        return (
            len(chosen),
            valid_left_context,
            last_body_clean,
            body_length_score - embedded_marker_penalty * 12 - short_body_penalty * 6 - leading_punct_penalty * 8,
        )

    def walk(position: int, expected_index: int | None, chosen: list[re.Match[str]]) -> None:
        nonlocal best, best_rank
        current_rank = rank(chosen)
        if current_rank > best_rank:
            best = list(chosen)
            best_rank = current_rank
        if expected_index is not None and expected_index >= 4:
            return
        for index in range(position, len(filtered)):
            match = filtered[index]
            letter = match.group(1).upper()
            if not chosen:
                if not _option_marker_has_valid_left_context(text, match.start()):
                    continue
                next_index = ord(letter) - ord("A") + 1 if letter < "D" else None
            else:
                if expected_index is None:
                    continue
                wanted = "ABCD"[expected_index]
                if letter != wanted:
                    continue
                next_index = expected_index + 1 if letter < "D" else None
            chosen.append(match)
            walk(index + 1, next_index, chosen)
            chosen.pop()

    walk(0, None, [])
    return best


def _split_option_fragments(text: str, *, marker_pattern: re.Pattern[str] = _OPTION_MARKER) -> list[tuple[str, str]]:
    normalized = (text or "").strip()
    if not normalized:
        return []

    all_matches = [
        match
        for match in marker_pattern.finditer(normalized)
        if match.group(1).upper() in {"A", "B", "C", "D"}
    ]
    matches = _best_ordered_option_match_sequence(normalized, all_matches)
    if len(matches) < 2:
        matches = [
            match for match in all_matches if _option_marker_has_valid_left_context(normalized, match.start())
        ]
    if not matches:
        match = _OPTION_PREFIX.match(normalized)
        if not match:
            return []
        return [(match.group(1).upper(), normalized[match.end() :].strip())]

    fragments: list[tuple[str, str]] = []
    for idx, match in enumerate(matches):
        body_start = match.end()
        body_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(normalized)
        body = normalized[body_start:body_end].strip()
        marker = _OPTION_TAIL_PASSAGE_MARKER.search(body)
        if marker and body[: marker.start()].strip():
            body = body[: marker.start()].rstrip()
        fragments.append((match.group(1).upper(), body))
    if len(fragments) >= 2 and fragments[0][0] != "A":
        suspicious_tail = any(
            body.strip()[:1] in ".,，。:：;；)）]】"
            for _letter, body in fragments[1:]
        )
        suspicious_embedded_marker = any(
            _match_looks_like_embedded_list_marker(normalized, match)
            for match in matches[1:]
        )
        if suspicious_tail or suspicious_embedded_marker:
            match = _OPTION_PREFIX.match(normalized)
            if match:
                return [(match.group(1).upper(), normalized[match.end() :].strip())]
    if len(fragments) > 1:
        letters = [letter for letter, _body in fragments]
        nonempty_count = sum(1 for _letter, body in fragments if body.strip())
        if not (letters[:4] == ["A", "B", "C", "D"] and nonempty_count == 0) and nonempty_count < 2:
            match = _OPTION_PREFIX.match(normalized)
            if match:
                return [(match.group(1).upper(), normalized[match.end() :].strip())]
    return fragments


def _normalize_option_letters(options: list[OptionNode]) -> list[OptionNode]:
    if not options:
        return options
    expected = [chr(ord("A") + idx) for idx in range(len(options))]
    actual = [option.letter for option in options]
    if actual == expected:
        return options
    if len(set(actual)) != len(actual) or any(letter not in {"A", "B", "C", "D"} for letter in actual):
        for option, letter in zip(options, expected):
            option.letter = letter
    return options


def _rebalance_image_only_options_from_stem_assets(
    stem_assets: list[AssetRef],
    options: list[OptionNode],
) -> tuple[list[AssetRef], list[OptionNode]]:
    if len(options) != 4 or not stem_assets:
        return stem_assets, options
    if any((option.text or "").strip() for option in options):
        return stem_assets, options

    option_images = [
        (option.image_path, option.source_page, option.page_region)
        for option in options
        if option.image_path
    ]
    if len(option_images) != len(options) - 1:
        return stem_assets, options

    if len(stem_assets) + len(option_images) != len(options):
        return stem_assets, options

    ordered_images = [
        (asset.path, asset.source_page, asset.page_region) for asset in stem_assets
    ] + option_images
    for option, (image_path, source_page, page_region) in zip(options, ordered_images):
        option.image_path = image_path
        option.source_page = source_page
        option.page_region = page_region
    return [], options


def _extract_inline_options_from_stem_text(stem: str) -> tuple[str, list[OptionNode]]:
    value = (stem or "").strip()
    if not value:
        return value, []
    matches = list(_INLINE_OPTION_MARKER.finditer(value))
    fragments: list[tuple[str, str]] = []
    prefix = value
    if len(matches) >= 4:
        start = _tail_option_cluster_start(matches, value)
        prefix = value[:start].rstrip()
        fragments = _split_option_fragments(value[start:], marker_pattern=_INLINE_OPTION_MARKER)
    nonempty_fragment_count = sum(1 for _letter, body in fragments if body.strip())
    if len(fragments) != 4 or (nonempty_fragment_count < 2 and nonempty_fragment_count != 0):
        bare_matches = list(_INLINE_BARE_OPTION_MARKER.finditer(value))
        if len(bare_matches) >= 4:
            start = _tail_option_cluster_start(bare_matches, value)
            prefix = value[:start].rstrip()
            fragments = _split_option_fragments(value[start:], marker_pattern=_INLINE_BARE_OPTION_MARKER)
            nonempty_fragment_count = sum(1 for _letter, body in fragments if body.strip())
        if len(fragments) != 4 or (nonempty_fragment_count < 2 and nonempty_fragment_count != 0):
            prefix, options = _extract_tail_combination_options_from_stem_text(value)
            if options:
                return prefix, options
            prefix, options = _extract_two_column_combination_options_from_stem_text(value)
            if options:
                return prefix, options
            return _extract_tail_multiline_options_from_stem_text(value)
    options = [OptionNode(letter=letter, text=body.strip()) for letter, body in fragments]
    return prefix, options


def _extract_tail_combination_options_from_stem_text(stem: str) -> tuple[str, list[OptionNode]]:
    value = (stem or "").strip()
    if not value:
        return value, []

    matches = [
        match
        for match in _INLINE_OPTION_MARKER.finditer(value)
        if match.group(1).upper() in {"A", "B", "C", "D"}
    ]
    if len(matches) < 4:
        return value, []

    for start_index in range(len(matches) - 4, -1, -1):
        tail = matches[start_index : start_index + 4]
        if tail[0].start() < len(value) * 0.35:
            continue
        bodies: list[str] = []
        valid = True
        for index, match in enumerate(tail):
            body_start = match.end()
            body_end = tail[index + 1].start() if index + 1 < len(tail) else len(value)
            body = re.sub(r"\s+", "", value[body_start:body_end].strip())
            if not _looks_like_combination_option_body(body):
                valid = False
                break
            bodies.append(body)
        if not valid:
            continue
        prefix = value[: tail[0].start()].rstrip()
        options = [
            OptionNode(letter=chr(ord("A") + index), text=body)
            for index, body in enumerate(bodies)
        ]
        return prefix, options
    return value, []


def _extract_two_column_combination_options_from_stem_text(stem: str) -> tuple[str, list[OptionNode]]:
    value = (stem or "").strip()
    if not value:
        return value, []

    lines = [line.rstrip() for line in value.splitlines()]
    for start in range(len(lines) - 1):
        rows: list[list[str]] = []
        for line in lines[start:]:
            stripped = line.strip()
            if not stripped:
                if rows:
                    break
                continue
            columns = [part.strip() for part in _TWO_COLUMN_OPTION_SPLIT.split(stripped) if part.strip()]
            if len(columns) < 2:
                if rows:
                    break
                continue
            parsed_columns: list[str] = []
            for column in columns[:2]:
                fragments = _split_option_fragments(column, marker_pattern=_INLINE_OPTION_MARKER)
                if len(fragments) != 1:
                    parsed_columns = []
                    break
                _letter, body = fragments[0]
                if not _looks_like_combination_option_body(body):
                    parsed_columns = []
                    break
                parsed_columns.append(re.sub(r"\s+", "", body))
            if len(parsed_columns) != 2:
                if rows:
                    break
                continue
            rows.append(parsed_columns)
            if len(rows) == 2:
                prefix = "\n".join(line for line in lines[:start] if line.strip()).strip()
                options = [
                    OptionNode(letter="A", text=rows[0][0]),
                    OptionNode(letter="B", text=rows[0][1]),
                    OptionNode(letter="C", text=rows[1][0]),
                    OptionNode(letter="D", text=rows[1][1]),
                ]
                return prefix, options
        if rows:
            break
    return value, []


def _tail_option_cluster_start(matches: list[re.Match[str]], text: str) -> int:
    for index in range(len(matches) - 4, -1, -1):
        letters = [matches[index + offset].group(1).upper() for offset in range(4)]
        if letters != ["A", "B", "C", "D"]:
            continue
        start = matches[index].start()
        if start >= len(text) * 0.35:
            return start
    return matches[0].start()


def _extract_tail_multiline_options_from_stem_text(stem: str) -> tuple[str, list[OptionNode]]:
    value = (stem or "").strip()
    if not value:
        return value, []

    lines = value.splitlines()
    first_option_index: int | None = None
    for index, raw in enumerate(lines):
        if _TAIL_OPTION_LINE.match((raw or "").strip()):
            first_option_index = index
            break
    if first_option_index is None:
        return value, []

    option_lines = lines[first_option_index:]
    options: list[OptionNode] = []
    current_letter: str | None = None
    current_body: list[str] = []

    def flush_current() -> None:
        nonlocal current_letter, current_body
        if current_letter is None:
            return
        options.append(OptionNode(letter=current_letter, text="\n".join(part for part in current_body if part).strip()))
        current_letter = None
        current_body = []

    for raw in option_lines:
        line = (raw or "").strip()
        if not line:
            continue
        direct = _TAIL_OPTION_LINE.match(line)
        if direct:
            letter = direct.group(1).upper()
            if current_letter is not None and letter != chr(ord(current_letter) + 1):
                return value, []
            flush_current()
            current_letter = letter
            current_body = [direct.group(2).strip()]
            continue

        if current_letter is None:
            return value, []

        embedded_markers = list(_INLINE_OPTION_MARKER.finditer(line))
        if embedded_markers:
            first_marker = embedded_markers[0]
            next_letter = first_marker.group(1).upper()
            if next_letter == chr(ord(current_letter) + 1):
                prefix = line[: first_marker.start()].strip()
                if prefix:
                    current_body.append(prefix)
                flush_current()
                fragments = _split_option_fragments(line[first_marker.start() :], marker_pattern=_INLINE_OPTION_MARKER)
                expected = "ABCD"[len(options)]
                if not fragments or fragments[0][0] != expected:
                    return value, []
                for letter, body in fragments:
                    options.append(OptionNode(letter=letter, text=body.strip()))
                current_letter = None
                current_body = []
                break
        current_body.append(line)

    flush_current()
    letters = [option.letter for option in options]
    if letters != ["A", "B", "C", "D"]:
        return value, []
    prefix = "\n".join(line for line in lines[:first_option_index] if line.strip()).strip()
    return prefix, options


def _content_spans(values: list[int], *, threshold: int, merge_gap: int = 0) -> list[tuple[int, int]]:
    spans: list[tuple[int, int]] = []
    start: int | None = None
    for index, count in enumerate(values):
        if count > threshold:
            if start is None:
                start = index
            continue
        if start is not None:
            spans.append((start, index - 1))
            start = None
    if start is not None:
        spans.append((start, len(values) - 1))

    if merge_gap <= 0 or not spans:
        return spans

    merged = [spans[0]]
    for start, end in spans[1:]:
        prev_start, prev_end = merged[-1]
        if start - prev_end - 1 <= merge_gap:
            merged[-1] = (prev_start, end)
        else:
            merged.append((start, end))
    return merged


def _content_row_spans(image: Image.Image, *, dilate_size: int, threshold_ratio: float, merge_gap: int) -> list[tuple[int, int]]:
    grayscale = image.convert("L")
    binary = grayscale.point(lambda pixel: 255 if pixel < 235 else 0)
    if dilate_size > 1:
        binary = binary.filter(ImageFilter.MaxFilter(dilate_size))
    width, height = binary.size
    pixels = binary.load()
    row_counts = [
        sum(1 for x in range(width) if pixels[x, y] > 0)
        for y in range(height)
    ]
    threshold = max(2, int(width * threshold_ratio))
    return _content_spans(row_counts, threshold=threshold, merge_gap=merge_gap)


def _content_col_spans(
    image: Image.Image,
    *,
    row_start: int,
    row_end: int,
    dilate_size: int,
    threshold_ratio: float,
    merge_gap: int,
) -> list[tuple[int, int]]:
    grayscale = image.convert("L")
    binary = grayscale.point(lambda pixel: 255 if pixel < 235 else 0)
    if dilate_size > 1:
        binary = binary.filter(ImageFilter.MaxFilter(dilate_size))
    width, _height = binary.size
    row_start = max(0, row_start)
    row_end = min(binary.size[1] - 1, row_end)
    pixels = binary.load()
    col_counts = [
        sum(1 for y in range(row_start, row_end + 1) if pixels[x, y] > 0)
        for x in range(width)
    ]
    threshold = max(2, int((row_end - row_start + 1) * threshold_ratio))
    return _content_spans(col_counts, threshold=threshold, merge_gap=merge_gap)


def _content_bbox(image: Image.Image, *, row_start: int, row_end: int) -> tuple[int, int, int, int] | None:
    grayscale = image.convert("L")
    width, _height = grayscale.size
    pixels = grayscale.load()
    coords: list[tuple[int, int]] = []
    for y in range(row_start, row_end + 1):
        for x in range(width):
            if pixels[x, y] < 235:
                coords.append((x, y))
    if not coords:
        return None
    xs = [x for x, _y in coords]
    ys = [y for _x, y in coords]
    return min(xs), min(ys), max(xs), max(ys)


def _padded_crop_box(box: tuple[int, int, int, int], image_size: tuple[int, int], *, margin: int = 4) -> tuple[int, int, int, int]:
    width, height = image_size
    x0, y0, x1, y1 = box
    return (
        max(0, x0 - margin),
        max(0, y0 - margin),
        min(width - 1, x1 + margin),
        min(height - 1, y1 + margin),
    )


def _save_cropped_choice_asset(source_path: str, image: Image.Image, box: tuple[int, int, int, int], suffix: str) -> str:
    stem = os.path.splitext(source_path)[0]
    target_path = f"{stem}_{suffix}.png"
    crop = image.crop((box[0], box[1], box[2] + 1, box[3] + 1))
    crop.save(target_path)
    return target_path


def _scaled_page_region(
    base_region: PageRegion | None,
    crop_box: tuple[int, int, int, int],
    *,
    image_size: tuple[int, int],
) -> PageRegion | None:
    if base_region is None:
        return None
    width = max(float(image_size[0]), 1.0)
    height = max(float(image_size[1]), 1.0)
    x_scale = (base_region.x1 - base_region.x0) / width
    y_scale = (base_region.y1 - base_region.y0) / height
    return PageRegion(
        page_number=base_region.page_number,
        x0=base_region.x0 + crop_box[0] * x_scale,
        y0=base_region.y0 + crop_box[1] * y_scale,
        x1=base_region.x0 + (crop_box[2] + 1) * x_scale,
        y1=base_region.y0 + (crop_box[3] + 1) * y_scale,
    )


def _split_choice_image_asset(asset: AssetRef, *, stem_text_present: bool) -> tuple[list[AssetRef], list[OptionNode]] | None:
    if not asset.path or not os.path.exists(asset.path):
        return None
    try:
        image = Image.open(asset.path).convert("RGB")
    except OSError:
        return None

    width, height = image.size
    coarse_rows = _content_row_spans(image, dilate_size=5, threshold_ratio=0.02, merge_gap=6)
    fine_rows = _content_row_spans(image, dilate_size=1, threshold_ratio=0.02, merge_gap=4)
    top_band: tuple[int, int] | None = None
    option_bands: list[tuple[int, int]] = []
    orientation: str | None = None

    if len(coarse_rows) == 4:
        option_bands = coarse_rows
        orientation = "vertical"
    elif len(coarse_rows) == 2:
        top_band = coarse_rows[0]
        option_bands = [coarse_rows[1]]
        orientation = "horizontal"
    elif len(coarse_rows) >= 5 and coarse_rows[-4][0] > int(height * 0.3):
        top_band = (coarse_rows[0][0], coarse_rows[-5][1])
        option_bands = coarse_rows[-4:]
        orientation = "vertical"
    elif len(fine_rows) == 2 and fine_rows[0][1] < int(height * 0.35):
        option_bbox = _content_bbox(image, row_start=fine_rows[1][0], row_end=fine_rows[1][1])
        if option_bbox is not None:
            option_width = option_bbox[2] - option_bbox[0] + 1
            option_height = option_bbox[3] - option_bbox[1] + 1
            top_band = fine_rows[0]
            option_bands = [fine_rows[1]]
            orientation = "vertical" if option_height >= option_width else "horizontal"
    elif len(fine_rows) >= 3 and fine_rows[0][1] < int(height * 0.45):
        top_band = fine_rows[0]
        option_bands = [(fine_rows[1][0], fine_rows[-1][1])]
        orientation = "horizontal" if width >= height else "vertical"
    elif stem_text_present:
        option_bands = [coarse_rows[0]] if coarse_rows else [(0, height - 1)]
        orientation = "horizontal" if width >= height else "vertical"
    else:
        image.close()
        return None

    stem_assets: list[AssetRef] = []
    if top_band is not None:
        top_box = _content_bbox(image, row_start=top_band[0], row_end=top_band[1])
        if top_box is not None:
            top_box = _padded_crop_box(top_box, image.size)
            top_path = _save_cropped_choice_asset(asset.path, image, top_box, "stem")
            stem_assets.append(
                AssetRef(
                    kind=asset.kind,
                    path=top_path,
                    source_page=asset.source_page,
                    page_region=_scaled_page_region(asset.page_region, top_box, image_size=image.size),
                )
            )

    option_boxes: list[tuple[int, int, int, int]] = []
    if orientation == "vertical" and len(option_bands) == 4:
        for index, (row_start, row_end) in enumerate(option_bands):
            box = _content_bbox(image, row_start=row_start, row_end=row_end)
            if box is None:
                continue
            option_boxes.append(_padded_crop_box(box, image.size))
    else:
        option_start = option_bands[0][0]
        option_end = option_bands[-1][1]
        option_bbox = _content_bbox(image, row_start=option_start, row_end=option_end)
        if option_bbox is None:
            image.close()
            return None
        option_bbox = _padded_crop_box(option_bbox, image.size)
        x0, y0, x1, y1 = option_bbox
        if orientation == "vertical":
            step = max((y1 - y0 + 1) / 4.0, 1.0)
            for index in range(4):
                start = int(round(y0 + index * step))
                end = int(round(y0 + (index + 1) * step)) - 1
                if index == 3:
                    end = y1
                option_boxes.append((x0, start, x1, max(start, end)))
        else:
            step = max((x1 - x0 + 1) / 4.0, 1.0)
            for index in range(4):
                start = int(round(x0 + index * step))
                end = int(round(x0 + (index + 1) * step)) - 1
                if index == 3:
                    end = x1
                option_boxes.append((start, y0, max(start, end), y1))

    if len(option_boxes) != 4:
        image.close()
        return None

    options: list[OptionNode] = []
    for index, box in enumerate(option_boxes):
        option_path = _save_cropped_choice_asset(asset.path, image, box, f"opt_{chr(ord('A') + index)}")
        options.append(
            OptionNode(
                letter=chr(ord("A") + index),
                text="",
                image_path=option_path,
                source_page=asset.source_page,
                page_region=_scaled_page_region(asset.page_region, box, image_size=image.size),
            )
        )
    image.close()
    return stem_assets, options


def _split_two_option_image_asset(asset: AssetRef) -> list[OptionNode] | None:
    if not asset.path or not os.path.exists(asset.path):
        return None
    try:
        image = Image.open(asset.path).convert("RGB")
    except OSError:
        return None

    coarse_rows = _content_row_spans(image, dilate_size=9, threshold_ratio=0.08, merge_gap=max(18, image.height // 48))
    if len(coarse_rows) != 2:
        image.close()
        return None

    options: list[OptionNode] = []
    for index, (row_start, row_end) in enumerate(coarse_rows):
        box = _content_bbox(image, row_start=row_start, row_end=row_end)
        if box is None:
            image.close()
            return None
        box = _padded_crop_box(box, image.size)
        option_path = _save_cropped_choice_asset(asset.path, image, box, f"optfrag_{index}")
        options.append(
            OptionNode(
                letter="",
                text="",
                image_path=option_path,
                source_page=asset.source_page,
                page_region=_scaled_page_region(asset.page_region, box, image_size=image.size),
            )
        )
    image.close()
    return options


def _build_option_nodes_from_boxes(
    asset: AssetRef,
    image: Image.Image,
    boxes: list[tuple[int, int, int, int]],
    *,
    letters: list[str],
    suffix_prefix: str,
) -> list[OptionNode]:
    options: list[OptionNode] = []
    for index, (box, letter) in enumerate(zip(boxes, letters)):
        option_path = _save_cropped_choice_asset(asset.path, image, box, f"{suffix_prefix}_{index}")
        options.append(
            OptionNode(
                letter=letter,
                text="",
                image_path=option_path,
                source_page=asset.source_page,
                page_region=_scaled_page_region(asset.page_region, box, image_size=image.size),
            )
        )
    return options


def _split_equal_vertical_boxes(
    image: Image.Image,
    *,
    expected_count: int,
) -> list[tuple[int, int, int, int]] | None:
    bbox = _content_bbox(image, row_start=0, row_end=image.size[1] - 1)
    if bbox is None:
        return None
    bbox = _padded_crop_box(bbox, image.size)
    x0, y0, x1, y1 = bbox
    step = max((y1 - y0 + 1) / float(expected_count), 1.0)
    boxes: list[tuple[int, int, int, int]] = []
    for index in range(expected_count):
        start = int(round(y0 + index * step))
        end = int(round(y0 + (index + 1) * step)) - 1
        if index == expected_count - 1:
            end = y1
        boxes.append((x0, start, x1, max(start, end)))
    return boxes


def _compress_row_spans(
    row_spans: list[tuple[int, int]],
    *,
    target_count: int,
) -> list[tuple[int, int]] | None:
    if len(row_spans) < target_count or target_count <= 0:
        return None
    if len(row_spans) == target_count:
        return list(row_spans)
    if len(row_spans) > target_count * 2 + 1:
        return None

    groups: list[tuple[int, int]] = []
    step = len(row_spans) / float(target_count)
    for index in range(target_count):
        start_index = int(round(index * step))
        end_index = int(round((index + 1) * step)) - 1
        if index == target_count - 1:
            end_index = len(row_spans) - 1
        start_index = min(start_index, len(row_spans) - 1)
        end_index = max(start_index, min(end_index, len(row_spans) - 1))
        groups.append((row_spans[start_index][0], row_spans[end_index][1]))
    return groups


def _split_option_rows_asset(
    asset: AssetRef,
    *,
    letters: list[str],
    allow_stem_prefix: bool = False,
) -> tuple[list[AssetRef], list[OptionNode]] | None:
    if not asset.path or not os.path.exists(asset.path):
        return None
    try:
        image = Image.open(asset.path).convert("RGB")
    except OSError:
        return None

    expected_count = len(letters)
    coarse_rows = _content_row_spans(image, dilate_size=5, threshold_ratio=0.02, merge_gap=max(6, image.height // 48))
    fine_rows = _content_row_spans(image, dilate_size=1, threshold_ratio=0.02, merge_gap=max(4, image.height // 64))
    sensitive_rows = _content_row_spans(
        image,
        dilate_size=1,
        threshold_ratio=0.01,
        merge_gap=max(1, image.height // 120),
    )

    row_spans: list[tuple[int, int]] | None = None
    stem_row: tuple[int, int] | None = None
    for candidate_rows in (coarse_rows, fine_rows, sensitive_rows):
        if allow_stem_prefix and len(candidate_rows) == expected_count + 1:
            stem_row = candidate_rows[0]
            row_spans = candidate_rows[1:]
            break
        if len(candidate_rows) == expected_count:
            row_spans = candidate_rows
            break
        if allow_stem_prefix:
            compressed_with_stem = _compress_row_spans(candidate_rows, target_count=expected_count + 1)
            if compressed_with_stem is not None:
                stem_row = compressed_with_stem[0]
                row_spans = compressed_with_stem[1:]
                break
        compressed_rows = _compress_row_spans(candidate_rows, target_count=expected_count)
        if compressed_rows is not None:
            row_spans = compressed_rows
            break

    if row_spans is None:
        if image.size[1] < image.size[0] * 1.4:
            image.close()
            return None
        equal_boxes = _split_equal_vertical_boxes(image, expected_count=expected_count)
        if equal_boxes is None:
            image.close()
            return None
        options = _build_option_nodes_from_boxes(asset, image, equal_boxes, letters=letters, suffix_prefix="optrow")
        image.close()
        return [], options

    stem_assets: list[AssetRef] = []
    if stem_row is not None:
        stem_box = _content_bbox(image, row_start=stem_row[0], row_end=stem_row[1])
        if stem_box is not None:
            stem_box = _padded_crop_box(stem_box, image.size)
            stem_path = _save_cropped_choice_asset(asset.path, image, stem_box, "stem")
            stem_assets.append(
                AssetRef(
                    kind=asset.kind,
                    path=stem_path,
                    source_page=asset.source_page,
                    page_region=_scaled_page_region(asset.page_region, stem_box, image_size=image.size),
                )
            )

    boxes: list[tuple[int, int, int, int]] = []
    for row_start, row_end in row_spans:
        box = _content_bbox(image, row_start=row_start, row_end=row_end)
        if box is None:
            image.close()
            return None
        boxes.append(_padded_crop_box(box, image.size))

    options = _build_option_nodes_from_boxes(asset, image, boxes, letters=letters, suffix_prefix="optrow")
    image.close()
    return stem_assets, options


def _split_grid_option_image_asset(asset: AssetRef) -> list[OptionNode] | None:
    if not asset.path or not os.path.exists(asset.path):
        return None
    try:
        image = Image.open(asset.path).convert("RGB")
    except OSError:
        return None

    row_spans = _content_row_spans(image, dilate_size=5, threshold_ratio=0.02, merge_gap=max(6, image.height // 48))
    if len(row_spans) != 2:
        image.close()
        return None

    boxes: list[tuple[int, int, int, int]] = []
    for row_index, (row_start, row_end) in enumerate(row_spans):
        row_box = _content_bbox(image, row_start=row_start, row_end=row_end)
        if row_box is None:
            image.close()
            return None
        col_spans = _content_col_spans(
            image,
            row_start=row_start,
            row_end=row_end,
            dilate_size=3,
            threshold_ratio=0.02,
            merge_gap=max(6, image.size[0] // 48),
        )
        if len(col_spans) != 2:
            image.close()
            return None
        for col_start, col_end in col_spans:
            x0 = max(col_start, row_box[0])
            x1 = min(col_end, row_box[2])
            if x1 < x0:
                image.close()
                return None
            box = (x0, row_box[1], x1, row_box[3])
            boxes.append(_padded_crop_box(box, image.size))

    options = _build_option_nodes_from_boxes(
        asset,
        image,
        boxes,
        letters=["A", "B", "C", "D"],
        suffix_prefix="optgrid",
    )
    image.close()
    return options


def _split_last_asset_into_all_options(question: QuestionNode) -> tuple[list[AssetRef], list[OptionNode]] | None:
    if len(question.stem_assets) < 2:
        return None
    last_asset = question.stem_assets[-1]

    grid_options = _split_grid_option_image_asset(last_asset)
    if grid_options is not None:
        return list(question.stem_assets[:-1]), grid_options

    row_split = _split_option_rows_asset(last_asset, letters=["A", "B", "C", "D"])
    if row_split is not None:
        extra_stem_assets, options = row_split
        if extra_stem_assets:
            return None
        return list(question.stem_assets[:-1]), options
    return None


def _split_last_two_assets_into_partial_options(question: QuestionNode) -> tuple[list[AssetRef], list[OptionNode]] | None:
    if len(question.stem_assets) < 2:
        return None
    prefix_assets = list(question.stem_assets[:-2])
    first_asset = question.stem_assets[-2]
    second_asset = question.stem_assets[-1]

    first_split = _split_option_rows_asset(
        first_asset,
        letters=["A", "B"],
        allow_stem_prefix=not bool(question.stem.strip()),
    )
    if first_split is None:
        return None
    first_stem_assets, first_options = first_split

    second_split = _split_option_rows_asset(second_asset, letters=["C", "D"])
    if second_split is None:
        return None
    second_stem_assets, second_options = second_split
    if second_stem_assets:
        return None

    return prefix_assets + first_stem_assets, first_options + second_options


def _split_all_option_image_asset(asset: AssetRef) -> list[OptionNode] | None:
    if not asset.path or not os.path.exists(asset.path):
        return None
    try:
        image = Image.open(asset.path).convert("RGB")
    except OSError:
        return None

    fine_rows = _content_row_spans(image, dilate_size=1, threshold_ratio=0.02, merge_gap=max(4, image.height // 64))
    coarse_rows = _content_row_spans(image, dilate_size=5, threshold_ratio=0.02, merge_gap=max(6, image.height // 48))
    option_rows = fine_rows if len(fine_rows) == 4 else coarse_rows if len(coarse_rows) == 4 else []
    if len(option_rows) != 4:
        image.close()
        return None

    options: list[OptionNode] = []
    for index, (row_start, row_end) in enumerate(option_rows):
        box = _content_bbox(image, row_start=row_start, row_end=row_end)
        if box is None:
            image.close()
            return None
        box = _padded_crop_box(box, image.size)
        option_path = _save_cropped_choice_asset(asset.path, image, box, f"optimg_{chr(ord('A') + index)}")
        options.append(
            OptionNode(
                letter=chr(ord("A") + index),
                text="",
                image_path=option_path,
                source_page=asset.source_page,
                page_region=_scaled_page_region(asset.page_region, box, image_size=image.size),
            )
        )
    image.close()
    return options


def _extract_duplicate_stem_asset_options(question: QuestionNode) -> tuple[list[AssetRef], list[OptionNode]] | None:
    indexed_assets = [
        (index, asset)
        for index, asset in enumerate(question.stem_assets)
        if asset.path and asset.page_region is not None
    ]
    if len(indexed_assets) < 2:
        return None

    best_candidate: tuple[float, tuple[list[AssetRef], list[OptionNode]]] | None = None
    for left_index in range(len(indexed_assets) - 1):
        left_raw_index, left_asset = indexed_assets[left_index]
        left_region = left_asset.page_region
        if left_region is None:
            continue
        for right_raw_index, right_asset in indexed_assets[left_index + 1 :]:
            right_region = right_asset.page_region
            if right_region is None or left_region.page_number != right_region.page_number:
                continue
            overlap_ratio = _region_overlap_ratio(left_region, right_region)
            if overlap_ratio < 0.92:
                continue

            preferred_index, preferred_asset = (
                (left_raw_index, left_asset)
                if (left_region.x1 - left_region.x0) * (left_region.y1 - left_region.y0)
                >= (right_region.x1 - right_region.x0) * (right_region.y1 - right_region.y0)
                else (right_raw_index, right_asset)
            )
            split = _split_choice_image_asset(preferred_asset, stem_text_present=bool(question.stem.strip()))
            if split is None:
                continue
            extra_stem_assets, split_options = split
            if len(split_options) != 4:
                continue

            drop_indices = {
                raw_index
                for raw_index, asset in indexed_assets
                if asset.page_region is not None
                and preferred_asset.page_region is not None
                and asset.page_region.page_number == preferred_asset.page_region.page_number
                and _region_overlap_ratio(asset.page_region, preferred_asset.page_region) >= 0.92
            }
            remaining = [
                asset
                for raw_index, asset in enumerate(question.stem_assets)
                if raw_index not in drop_indices
            ]
            remaining.extend(extra_stem_assets)
            score = overlap_ratio * max(
                (preferred_asset.page_region.x1 - preferred_asset.page_region.x0)
                * (preferred_asset.page_region.y1 - preferred_asset.page_region.y0),
                1.0,
            )
            candidate = (remaining, split_options)
            if best_candidate is None or score > best_candidate[0]:
                best_candidate = (score, candidate)

    return best_candidate[1] if best_candidate is not None else None


def _extract_option_images_from_stem_assets(question: QuestionNode) -> tuple[list[AssetRef], list[OptionNode]] | None:
    if not question.stem_assets:
        return None

    if len(question.stem_assets) == 1 and not question.stem.strip():
        preferred_split = _split_choice_image_asset(
            question.stem_assets[0],
            stem_text_present=False,
        )
        if preferred_split is not None:
            split_stem_assets, split_options = preferred_split
            if split_options:
                return split_stem_assets, split_options

    duplicate_split = _extract_duplicate_stem_asset_options(question)
    if duplicate_split is not None:
        return duplicate_split

    best_high_confidence: tuple[int, tuple[list[AssetRef], list[OptionNode]]] | None = None
    for index, asset in enumerate(question.stem_assets):
        split_options = _split_all_option_image_asset(asset)
        if split_options is not None:
            remaining = [item for idx, item in enumerate(question.stem_assets) if idx != index]
            score = int((asset.page_region.x1 - asset.page_region.x0) * (asset.page_region.y1 - asset.page_region.y0)) if asset.page_region else 0
            candidate = (remaining, split_options)
            if best_high_confidence is None or score > best_high_confidence[0]:
                best_high_confidence = (score, candidate)
            continue

        row_split = _split_option_rows_asset(asset, letters=["A", "B", "C", "D"])
        if row_split is not None:
            extra_stem_assets, split_options = row_split
            remaining = [item for idx, item in enumerate(question.stem_assets) if idx != index] + list(extra_stem_assets)
            score = int((asset.page_region.x1 - asset.page_region.x0) * (asset.page_region.y1 - asset.page_region.y0)) if asset.page_region else 0
            candidate = (remaining, split_options)
            if best_high_confidence is None or score > best_high_confidence[0]:
                best_high_confidence = (score, candidate)
    if best_high_confidence is not None:
        return best_high_confidence[1]

    if len(question.stem_assets) == 2:
        indexed_assets = list(enumerate(question.stem_assets))
        indexed_assets.sort(
            key=lambda item: (
                item[1].source_page or 0,
                item[1].page_region.y0 if item[1].page_region is not None else 0.0,
            )
        )
        stem_index, stem_asset = indexed_assets[0]
        option_index, option_asset = indexed_assets[-1]
        if option_asset.path and option_asset.page_region is not None:
            region = option_asset.page_region
            width = max(region.x1 - region.x0, 1.0)
            height = max(region.y1 - region.y0, 1.0)
            if width / height >= 2.2 and width * height >= 4000:
                split = _split_choice_image_asset(option_asset, stem_text_present=bool(question.stem.strip()))
                if split is not None:
                    extra_stem_assets, split_options = split
                    remaining = [item for idx, item in enumerate(question.stem_assets) if idx not in {option_index}]
                    if stem_index != option_index:
                        remaining = [item for item in remaining if item.path != stem_asset.path] + [stem_asset]
                    remaining.extend(extra_stem_assets)
                    return remaining, split_options
    return None


def _synthesize_labeled_point_options(question: QuestionNode) -> QuestionNode:
    stem = (question.stem or "").strip()
    if question.options or not question.stem_assets:
        return question
    if not _POINT_DIAGRAM_PROMPT.search(stem):
        return question
    question.options = [
        OptionNode(letter=letter, text=f"{letter}点")
        for letter in "ABCD"
    ]
    return question


def _image_is_nearly_uniform(path: str | None) -> bool:
    if not path or not os.path.exists(path):
        return False
    try:
        image = Image.open(path).convert("L")
    except OSError:
        return False
    extrema = image.getextrema()
    histogram = image.histogram()
    pixel_count = max(sum(histogram), 1)
    near_black = sum(histogram[:8]) / pixel_count
    near_white = sum(histogram[248:]) / pixel_count
    image.close()
    if extrema is None:
        return False
    return (extrema[1] - extrema[0] <= 6) or near_black >= 0.92 or near_white >= 0.98


def _refresh_option_images_from_pdf_crop(
    question: QuestionNode,
    cropper: _PdfRenderCropper | None,
) -> QuestionNode:
    if cropper is None:
        return question
    for option in question.options:
        region = option.page_region
        if region is None or not _image_is_nearly_uniform(option.image_path):
            continue
        cropped = cropper.crop_nonwhite(
            page_number=region.page_number,
            rect=(region.x0 - 4.0, region.y0 - 4.0, region.x1 + 4.0, region.y1 + 4.0),
            suffix=f"q{question.source_number}_{option.letter}_refresh",
        )
        if cropped is None:
            continue
        option.image_path, option.page_region = cropped
        option.source_page = option.page_region.page_number
    return question


def _upgrade_placeholder_text_options_from_stem_asset(question: QuestionNode) -> QuestionNode:
    if len(question.options) != 4 or len(question.stem_assets) != 1:
        return question
    if any(option.image_path for option in question.options):
        return question

    normalized_texts = [
        "".join((option.text or "").split()).strip(".,，。:：;；")
        for option in question.options
    ]
    if not all(normalized_texts):
        return question
    if len(set(normalized_texts)) != 1:
        return question

    placeholder = normalized_texts[0]
    if len(placeholder) > 6 and "图" not in placeholder:
        return question
    if placeholder not in {"时刻", "如图所示", "如下图所示", "如上图所示"} and "图" not in placeholder:
        return question

    split = _split_choice_image_asset(question.stem_assets[0], stem_text_present=bool(question.stem.strip()))
    if split is None:
        return question
    split_stem_assets, split_options = split
    if len(split_options) != 4:
        return question
    question.stem_assets = split_stem_assets
    question.options = split_options
    return question


def _synthesize_missing_options(question: QuestionNode) -> QuestionNode:
    if question.options:
        return question

    updated_stem, inline_options = _extract_inline_options_from_stem_text(question.stem)
    if inline_options:
        question.stem = updated_stem
        question.options = inline_options
        question = _repair_blank_image_options(question)
        return question

    extracted_assets = _extract_option_images_from_stem_assets(question)
    if extracted_assets is not None:
        question.stem_assets, question.options = extracted_assets
        return question

    question = _synthesize_labeled_point_options(question)
    if question.options:
        return question

    split_all_assets = _split_last_asset_into_all_options(question)
    if split_all_assets is not None:
        question.stem_assets, question.options = split_all_assets
        return question

    split_partial_assets = _split_last_two_assets_into_partial_options(question)
    if split_partial_assets is not None:
        question.stem_assets, question.options = split_partial_assets
        return question

    if len(question.stem_assets) == 1 and not question.stem.strip():
        single_asset_split = _split_choice_image_asset(question.stem_assets[0], stem_text_present=True)
        if single_asset_split is not None:
            split_stem_assets, split_options = single_asset_split
            if split_options:
                if split_stem_assets:
                    question.stem_assets = split_stem_assets
                question.options = split_options
                return question

    if len(question.stem_assets) >= 2:
        trailing_options = _split_all_option_image_asset(question.stem_assets[-1])
        if trailing_options is not None:
            question.options = trailing_options
            question.stem_assets = list(question.stem_assets[:-1])
            return question

    if len(question.stem_assets) == 4 and question.stem.strip():
        question.options = [
            OptionNode(
                letter=chr(ord("A") + index),
                text="",
                image_path=asset.path,
                source_page=asset.source_page,
                page_region=asset.page_region,
            )
            for index, asset in enumerate(question.stem_assets[:4])
        ]
        question.stem_assets = []
        return question

    if len(question.stem_assets) == 1:
        split = _split_choice_image_asset(question.stem_assets[0], stem_text_present=bool(question.stem.strip()))
        if split is None:
            return question
        stem_assets, options = split
        question.stem_assets = stem_assets
        question.options = options
        return question

    if len(question.stem_assets) == 2 and question.stem.strip():
        split_options: list[OptionNode] = []
        for asset in question.stem_assets:
            asset_options = _split_two_option_image_asset(asset)
            if asset_options is None:
                split_options = []
                break
            split_options.extend(asset_options)
        if len(split_options) == 4:
            split_options.sort(
                key=lambda option: (
                    option.source_page or 0,
                    option.page_region.y0 if option.page_region else 0.0,
                    option.page_region.x0 if option.page_region else 0.0,
                )
            )
            question.options = [
                OptionNode(
                    letter=chr(ord("A") + index),
                    text="",
                    image_path=option.image_path,
                    source_page=option.source_page,
                    page_region=option.page_region,
                )
                for index, option in enumerate(split_options)
            ]
            question.stem_assets = []
    return question


def _parse_numeric_source_number(value: str) -> int | None:
    raw = (value or "").strip()
    return int(raw) if raw.isdigit() else None


def _question_needs_promoted_stem_asset(question: QuestionNode) -> bool:
    if question.stem.strip() or question.stem_assets or len(question.options) != 4:
        return False
    if all(option.image_path and not (option.text or "").strip() for option in question.options):
        return True
    if all((option.text or "").strip() and not option.image_path for option in question.options):
        return True
    return False


def _question_has_complete_text_options(question: QuestionNode) -> bool:
    if len(question.options) != 4:
        return False
    if [option.letter for option in question.options] != ["A", "B", "C", "D"]:
        return False
    return all((option.text or "").strip() for option in question.options)


def _question_expects_visual_choice_options(question: QuestionNode) -> bool:
    stem = (question.stem or "").strip()
    if not stem:
        return False
    return any(
        marker in stem
        for marker in (
            "图形",
            "下图",
            "图中",
            "图一",
            "图二",
            "示意图",
            "网络",
            "规律性",
        )
    )


def _promote_trailing_option_image_to_next_stem(questions: list[QuestionNode]) -> None:
    for index in range(len(questions) - 1):
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if current_number is None or next_number != current_number + 1:
            continue
        if not _question_has_complete_text_options(current):
            continue
        if not _question_needs_promoted_stem_asset(nxt):
            continue
        if any(option.image_path for option in current.options[:-1]):
            continue

        donor = current.options[-1]
        if not donor.image_path or not (donor.text or "").strip():
            continue

        nxt.stem_assets.insert(
            0,
            AssetRef(
                kind="stem_image",
                path=donor.image_path,
                source_page=donor.source_page,
                page_region=donor.page_region,
            ),
        )
        if donor.source_page is not None:
            nxt.page_numbers = sorted({*nxt.page_numbers, donor.source_page})
        donor.image_path = None
        donor.source_page = None
        donor.page_region = None


def _promote_leading_option_image_to_previous_question(questions: list[QuestionNode]) -> None:
    for index in range(len(questions) - 1):
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if current_number is None or next_number != current_number + 1:
            continue
        if current.options or current.stem_assets:
            continue
        if not _question_expects_visual_choice_options(current):
            continue
        if not _question_has_complete_text_options(nxt):
            continue

        donor = nxt.options[0]
        if not donor.image_path or not (donor.text or "").strip():
            continue
        if any(option.image_path for option in nxt.options[1:]):
            continue

        current.stem_assets.append(
            AssetRef(
                kind="stem_image",
                path=donor.image_path,
                source_page=donor.source_page,
                page_region=donor.page_region,
            ),
        )
        if donor.source_page is not None:
            current.page_numbers = sorted({*current.page_numbers, donor.source_page})
        donor.image_path = None
        donor.source_page = None
        donor.page_region = None
        repaired = _synthesize_missing_options(current)
        current.stem_assets = repaired.stem_assets
        current.options = repaired.options


def _promote_trailing_stem_asset_to_previous_question(questions: list[QuestionNode]) -> None:
    for index in range(len(questions) - 1):
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if current_number is None or next_number != current_number + 1:
            continue
        if current.options or current.stem_assets:
            continue
        if not _question_expects_visual_choice_options(current):
            continue
        if not _question_has_complete_text_options(nxt):
            continue
        if len(nxt.stem_assets) < 2:
            continue

        indexed_assets = list(enumerate(nxt.stem_assets))
        indexed_assets.sort(
            key=lambda item: (
                item[1].source_page or 0,
                item[1].page_region.y0 if item[1].page_region is not None else 0.0,
                item[1].page_region.x0 if item[1].page_region is not None else 0.0,
            )
        )
        donor_index, donor = indexed_assets[-1]
        original_pages = list(current.page_numbers)
        current.stem_assets.append(
            AssetRef(
                kind=donor.kind,
                path=donor.path,
                source_page=donor.source_page,
                page_region=donor.page_region,
            )
        )
        if donor.source_page is not None:
            current.page_numbers = sorted({*current.page_numbers, donor.source_page})
        repaired = _synthesize_missing_options(current)
        if len(repaired.options) == 4:
            del nxt.stem_assets[donor_index]
            continue

        current.stem_assets = []
        current.options = []
        current.page_numbers = original_pages


def _repair_blank_image_options(question: QuestionNode) -> QuestionNode:
    if len(question.options) != 4:
        return question

    blank_indices = [index for index, option in enumerate(question.options) if not (option.text or "").strip()]
    if not blank_indices:
        return question

    if len(question.stem_assets) in {len(blank_indices), len(blank_indices) + 1}:
        transferable_assets = list(question.stem_assets[-len(blank_indices) :])
        preserved_assets = list(question.stem_assets[:-len(blank_indices)])
        for target_index, asset in zip(blank_indices, transferable_assets):
            question.options[target_index].image_path = asset.path
            question.options[target_index].source_page = asset.source_page
            question.options[target_index].page_region = asset.page_region
        question.stem_assets = preserved_assets
        return question

    first_blank = blank_indices[0]
    candidate_slots = [
        index
        for index, option in enumerate(question.options)
        if index >= max(0, first_blank - 1) and option.image_path
    ]
    if len(candidate_slots) != len(blank_indices):
        return question

    candidate_images = [
        (
            question.options[index].image_path,
            question.options[index].source_page,
            question.options[index].page_region,
        )
        for index in candidate_slots
    ]
    for index in candidate_slots:
        question.options[index].image_path = None
        question.options[index].source_page = None
        question.options[index].page_region = None

    for target_index, (image_path, source_page, page_region) in zip(blank_indices, candidate_images):
        question.options[target_index].image_path = image_path
        question.options[target_index].source_page = source_page
        question.options[target_index].page_region = page_region
    return question


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
            key = unicodedata.normalize("NFKC", (text or "").strip())
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

    def find_near(
        self,
        text: str,
        *,
        page_numbers: set[int] | None = None,
        anchor_y: float | None = None,
    ) -> PageTextLine | None:
        key = unicodedata.normalize("NFKC", (text or "").strip())
        if not key:
            return None
        candidates = self._index.get(key, [])
        best: PageTextLine | None = None
        best_score: tuple[float, float, float] | None = None
        for idx in candidates:
            line = self._lines[idx]
            page_penalty = 0.0 if not page_numbers or line.page_number in page_numbers else 10000.0
            cursor_penalty = 0.0 if idx > self._cursor else 500.0
            y_penalty = abs(line.y0 - anchor_y) if anchor_y is not None else 0.0
            score = (page_penalty, y_penalty, cursor_penalty)
            if best_score is None or score < best_score:
                best = line
                best_score = score
        return best


class _PdfRenderCropper:
    _ZOOM = 4.0

    def __init__(self, pdf_path: str | None):
        self._pdf_path = pdf_path
        self._doc = None
        self._temp_dir = tempfile.mkdtemp(prefix="pptconvert_page_crops_") if pdf_path else None
        self._counter = 0

    def _ensure_doc(self):
        if not self._pdf_path:
            return None
        if self._doc is None:
            pdf_extract.require_fitz()
            if pdf_extract.fitz is None:  # pragma: no cover
                return None
            self._doc = pdf_extract.fitz.open(self._pdf_path)
        return self._doc

    def crop_nonwhite(
        self,
        *,
        page_number: int,
        rect: tuple[float, float, float, float],
        suffix: str,
    ) -> tuple[str, PageRegion] | None:
        doc = self._ensure_doc()
        if doc is None or self._temp_dir is None or page_number < 1 or page_number > len(doc):
            return None
        clip = pdf_extract.fitz.Rect(*rect)
        if clip.width <= 1 or clip.height <= 1:
            return None
        page = doc[page_number - 1]
        pix = page.get_pixmap(
            matrix=pdf_extract.fitz.Matrix(self._ZOOM, self._ZOOM),
            clip=clip,
            alpha=False,
        )
        filename = os.path.join(self._temp_dir, f"p{page_number}_{suffix}_{self._counter}.png")
        self._counter += 1
        pix.save(filename)

        image = Image.open(filename).convert("RGB")
        grayscale = image.convert("L")
        mask = grayscale.point(lambda value: 255 if value < 245 else 0)
        bbox = mask.getbbox()
        if bbox is None:
            image.close()
            os.remove(filename)
            return None

        cropped = image.crop(bbox)
        cropped.save(filename)
        image.close()
        x_scale = clip.width / max(float(pix.width), 1.0)
        y_scale = clip.height / max(float(pix.height), 1.0)
        region = PageRegion(
            page_number=page_number,
            x0=clip.x0 + bbox[0] * x_scale,
            y0=clip.y0 + bbox[1] * y_scale,
            x1=clip.x0 + bbox[2] * x_scale,
            y1=clip.y0 + bbox[3] * y_scale,
        )
        return filename, region

    def close(self) -> None:
        if self._doc is not None:
            self._doc.close()
            self._doc = None


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


def _normalize_stem_line(line: str, *, first: bool) -> str:
    normalized = (line or "").strip()
    if first and _LEADING_BLANK_PUNCT.match(normalized):
        return "____" + normalized
    return normalized


def _merge_stem_fragment(left: str, right: str) -> str:
    left = (left or "").rstrip()
    right = (right or "").lstrip()
    if not left:
        return right
    if not right:
        return left

    tail = _TRAILING_SHORT_NUMBER.search(left)
    head = _LEADING_SHORT_NUMBER.match(right)
    if tail and head:
        suffix = head.group(1)
        rest = head.group(2)
        if len(suffix) >= 2 and suffix[0] == tail.group(1)[-1]:
            suffix = suffix[1:]
        return left + "/" + suffix + rest

    if right[:1] in "-—－/" or left.endswith(("-","—","－","/")):
        return left + right
    return left + (" " if _needs_space_between(left, right) else "") + right


def _join_stem_lines(lines: list[str]) -> str:
    merged = ""
    saw_enumeration = False
    for index, raw_line in enumerate(item for item in lines if item and item.strip()):
        line = _normalize_stem_line(raw_line, first=index == 0)
        if not merged:
            merged = line
            saw_enumeration = bool(_ENUMERATED_STEM_LINE.match(line))
            continue
        if _TRAILING_SHORT_NUMBER.search(merged.rstrip()) and _LEADING_SHORT_NUMBER.match(line):
            merged = _merge_stem_fragment(merged, line)
            continue
        if _ENUMERATED_STEM_LINE.match(line):
            merged += "\n" + line
            saw_enumeration = True
            continue
        if saw_enumeration and _PROMPT_AFTER_ENUMERATION.match(line):
            merged += "\n" + line
            continue
        merged = _merge_stem_fragment(merged, line)
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
    for page_number in sorted(grouped):
        page_regions = grouped[page_number]
        merged.append(
            PageRegion(
                page_number=page_number,
                x0=min(region.x0 for region in page_regions),
                y0=min(region.y0 for region in page_regions),
                x1=max(region.x1 for region in page_regions),
                y1=max(region.y1 for region in page_regions),
            )
        )
    return merged


def _blank_option_marker_lines(
    option_lines: list[RichLine],
    locator: _LayoutLocator,
    *,
    page_numbers: set[int] | None = None,
    anchor_y: float | None = None,
) -> dict[str, PageTextLine]:
    option_texts = [_rich_line_text(line) for line in option_lines if _rich_line_text(line)]
    markers: dict[str, PageTextLine] = {}
    for text in option_texts:
        fragments = _split_option_fragments(text)
        if len(fragments) != 1:
            continue
        letter, body = fragments[0]
        line = locator.find_near(text, page_numbers=page_numbers, anchor_y=anchor_y)
        if line is None:
            continue
        if body.strip():
            markers.setdefault(letter, line)
            continue
        markers[letter] = line
    return markers


def _assign_option_row_stem_assets(
    question: QuestionNode,
    marker_lines: Mapping[str, PageTextLine],
) -> QuestionNode:
    if not question.stem_assets or not marker_lines:
        return question

    blank_options = {
        option.letter: option
        for option in question.options
        if option.letter in marker_lines and not option.image_path and not (option.text or "").strip()
    }
    if not blank_options:
        return question

    remaining_assets: list[AssetRef] = []
    for asset in question.stem_assets:
        region = asset.page_region
        if region is None:
            remaining_assets.append(asset)
            continue
        best_letter: str | None = None
        best_distance: float | None = None
        asset_center_x = (region.x0 + region.x1) / 2.0
        for letter, marker in marker_lines.items():
            option = blank_options.get(letter)
            if option is None:
                continue
            if region.page_number != marker.page_number:
                continue
            if region.y1 < marker.y0 - 18 or region.y0 > marker.y1 + 18:
                continue
            marker_center_x = (marker.x0 + marker.x1) / 2.0
            distance = abs(asset_center_x - marker_center_x)
            if distance > 48:
                continue
            if best_distance is None or distance < best_distance:
                best_letter = letter
                best_distance = distance
        if best_letter is None:
            remaining_assets.append(asset)
            continue
        option = blank_options.pop(best_letter)
        option.image_path = asset.path
        option.source_page = asset.source_page
        option.page_region = asset.page_region
    question.stem_assets = remaining_assets
    return question


def _fill_blank_option_images_from_pdf_crops(
    question: QuestionNode,
    marker_lines: Mapping[str, PageTextLine],
    cropper: _PdfRenderCropper | None,
) -> QuestionNode:
    if cropper is None or not marker_lines:
        return question

    row_markers: dict[tuple[int, int], list[tuple[str, PageTextLine]]] = defaultdict(list)
    for letter, line in marker_lines.items():
        row_key = (line.page_number, int(round(line.y0 / 8.0)))
        row_markers[row_key].append((letter, line))

    for markers in row_markers.values():
        markers.sort(key=lambda item: item[1].x0)
        for index, (letter, line) in enumerate(markers):
            option = next((item for item in question.options if item.letter == letter), None)
            if option is None or option.image_path or (option.text or "").strip():
                continue
            next_line = markers[index + 1][1] if index + 1 < len(markers) else None
            left = line.x1 + 2.0
            right = (next_line.x0 - 2.0) if next_line is not None else max(line.block_x1 or line.x1, line.x1 + 90.0)
            top = min(line.block_y0 or line.y0, line.y0) - 8.0
            bottom = max(line.block_y1 or line.y1, line.y1) + 18.0
            if right <= left + 4:
                continue
            cropped = cropper.crop_nonwhite(
                page_number=line.page_number,
                rect=(left, top, right, bottom),
                suffix=f"q{question.source_number}_{letter}",
            )
            if cropped is None:
                continue
            image_path, region = cropped
            option.image_path = image_path
            option.source_page = line.page_number
            option.page_region = region
    return question


def _fill_missing_stem_from_pdf_crop(
    question: QuestionNode,
    *,
    source_number: str,
    locator: _LayoutLocator,
    marker_lines: Mapping[str, PageTextLine],
    cropper: _PdfRenderCropper | None,
) -> QuestionNode:
    if cropper is None or question.stem.strip() or question.stem_assets or not marker_lines:
        return question

    pages = {line.page_number for line in marker_lines.values()}
    top_marker = min(marker_lines.values(), key=lambda line: (line.page_number, line.y0))
    number_line = locator.find_near(
        f"{source_number}.",
        page_numbers=pages or None,
        anchor_y=top_marker.y0 - 32.0,
    )
    if number_line is None or number_line.page_number != top_marker.page_number:
        return question

    right_edge = max(
        line.block_x1 if line.block_x1 is not None else line.x1
        for line in marker_lines.values()
        if line.page_number == top_marker.page_number
    )
    top = min(number_line.block_y0 or number_line.y0, number_line.y0) - 4.0
    bottom = top_marker.y0 - 6.0
    left = number_line.x1 + 4.0
    right = max(right_edge + 80.0, left + 40.0)
    if bottom <= top + 6.0:
        return question

    cropped = cropper.crop_nonwhite(
        page_number=number_line.page_number,
        rect=(left, top, right, bottom),
        suffix=f"q{source_number}_stem",
    )
    if cropped is None:
        return question

    image_path, region = cropped
    question.stem_assets.append(
        AssetRef(
            kind="stem_image",
            path=image_path,
            source_page=number_line.page_number,
            page_region=region,
        )
    )
    if number_line.page_number not in question.page_numbers:
        question.page_numbers.append(number_line.page_number)
        question.page_numbers.sort()
    return question


def _merge_adjacent_page_regions(regions: list[PageRegion]) -> list[PageRegion]:
    """合并相邻页面且纵向区域接近连续的 PageRegion（跨页表格/材料正文）。"""
    if len(regions) <= 1:
        return list(regions)
    ordered = sorted(regions, key=lambda r: (r.page_number, r.y0))
    merged: list[PageRegion] = [ordered[0]]
    for region in ordered[1:]:
        prev = merged[-1]
        if (
            region.page_number == prev.page_number + 1
            and abs(region.x0 - prev.x0) < 60.0
            and abs(region.x1 - prev.x1) < 60.0
        ):
            merged[-1] = PageRegion(
                page_number=prev.page_number,
                x0=min(prev.x0, region.x0),
                y0=prev.y0,
                x1=max(prev.x1, region.x1),
                y1=prev.y1,
            )
            merged.append(region)
        else:
            merged.append(region)
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


def _assigned_question_image_paths(question: QuestionNode) -> set[str]:
    paths = {asset.path for asset in question.stem_assets if asset.path}
    paths.update(option.image_path for option in question.options if option.image_path)
    return {path for path in paths if path}


def _x_overlap_ratio(a: PageRegion, b: PageRegion) -> float:
    overlap = min(a.x1, b.x1) - max(a.x0, b.x0)
    if overlap <= 0:
        return 0.0
    width = min(max(a.x1 - a.x0, 1.0), max(b.x1 - b.x0, 1.0))
    return float(overlap) / float(width)


def _region_overlap_ratio(a: PageRegion, b: PageRegion) -> float:
    overlap_width = min(a.x1, b.x1) - max(a.x0, b.x0)
    overlap_height = min(a.y1, b.y1) - max(a.y0, b.y0)
    if overlap_width <= 0 or overlap_height <= 0:
        return 0.0
    overlap_area = overlap_width * overlap_height
    area_a = max((a.x1 - a.x0) * (a.y1 - a.y0), 1.0)
    area_b = max((b.x1 - b.x0) * (b.y1 - b.y0), 1.0)
    return float(overlap_area) / float(min(area_a, area_b))


def _material_question_page_starts(question_line_groups: list[list[PageTextLine]]) -> dict[int, float]:
    starts: dict[int, float] = {}
    for lines in question_line_groups:
        for line in lines:
            current = starts.get(line.page_number)
            if current is None or line.y0 < current:
                starts[line.page_number] = float(line.y0)
    return starts


def _attach_orphan_material_images(
    material: MaterialSet,
    *,
    question_line_groups: list[list[PageTextLine]],
    image_regions: Mapping[str, ExtractedImageRegion] | None = None,
) -> None:
    if not material.body_regions or not image_regions:
        return

    assigned_paths = {asset.path for asset in material.body_assets if asset.path}
    for question in material.questions:
        assigned_paths.update(_assigned_question_image_paths(question))

    page_regions = {region.page_number: region for region in material.body_regions}
    question_page_starts = _material_question_page_starts(question_line_groups)
    extra_assets: list[AssetRef] = []
    extra_regions: list[PageRegion] = []

    for info in sorted(image_regions.values(), key=lambda item: (item.page_number, item.y0, item.x0)):
        if not info.path or info.path in assigned_paths:
            continue
        page_region = page_regions.get(info.page_number)
        if page_region is None:
            continue
        image_region = _region_from_extracted_image(info)
        if image_region is None:
            continue
        question_start = question_page_starts.get(info.page_number)
        if question_start is not None and info.y0 >= question_start - 24.0:
            continue
        overlap_ratio = _x_overlap_ratio(page_region, image_region)
        within_column = (
            info.x0 >= page_region.x0 - 36.0
            and info.x1 <= page_region.x1 + 36.0
        )
        if overlap_ratio < 0.35 and not within_column:
            continue
        upper_bound = question_start - 12.0 if question_start is not None else page_region.y1 + 80.0
        if info.y0 > upper_bound:
            continue
        if info.y1 < page_region.y0 - 280.0:
            continue
        assigned_paths.add(info.path)
        extra_assets.append(
            AssetRef(
                kind="material_inline_image",
                path=info.path,
                source_page=info.page_number,
                page_region=image_region,
            )
        )
        extra_regions.append(image_region)

    if extra_assets:
        material.body_assets.extend(extra_assets)
        material.body_regions = _coalesce_regions([*material.body_regions, *extra_regions])


def _promote_leading_question_assets_to_material(material: MaterialSet) -> None:
    if (material.body or "").strip() or material.body_assets or not material.questions:
        return
    first_question = material.questions[0]
    if not first_question.stem_assets or len(first_question.options) < 4:
        return
    if any(option.image_path for option in first_question.options):
        return
    promoted_assets = list(first_question.stem_assets)
    material.body_assets.extend(promoted_assets)
    material.body_regions = _coalesce_regions(
        [
            *material.body_regions,
            *[
                asset.page_region
                for asset in promoted_assets
                if asset.page_region is not None
            ],
        ]
    )
    first_question.stem_assets = []


def _copy_asset(path: str, asset_dir: str, seen: dict[str, str]) -> str:
    source = os.path.abspath(path)
    target_dir = os.path.abspath(asset_dir)
    cached = seen.get(source)
    if cached:
        return cached

    source_dir = os.path.dirname(source)
    try:
        already_materialized = os.path.samefile(source_dir, target_dir)
    except OSError:
        already_materialized = os.path.normcase(os.path.normpath(source_dir)) == os.path.normcase(
            os.path.normpath(target_dir)
        )
    if already_materialized:
        seen[source] = source
        return source

    name = os.path.basename(source)
    stem, ext = os.path.splitext(name)
    candidate = os.path.join(target_dir, name)
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
    *,
    located_lines: list[PageTextLine] | None = None,
    page_cropper: _PdfRenderCropper | None = None,
) -> QuestionNode:
    stem_text_lines = [_rich_line_text(line) for line in stem_lines if _rich_line_text(line)]
    located_lines = list(located_lines) if located_lines is not None else locator.consume(stem_text_lines)
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

    marker_page_numbers = set(page_numbers)
    marker_page_numbers.update(
        asset.source_page for asset in stem_assets if asset.source_page is not None
    )
    anchor_y_candidates = [line.y1 for line in located_lines]
    anchor_y_candidates.extend(
        asset.page_region.y1
        for asset in stem_assets
        if asset.page_region is not None
    )
    marker_lines = _blank_option_marker_lines(
        option_lines,
        locator,
        page_numbers=marker_page_numbers or None,
        anchor_y=max(anchor_y_candidates) if anchor_y_candidates else None,
    )

    options: list[OptionNode] = []
    current: Optional[OptionNode] = None
    pending_option_images: list[tuple[str, PageRegion | None]] = []

    def attach_pending_image(option: Optional[OptionNode]) -> None:
        if option is None or option.image_path or not pending_option_images:
            return
        image_path, region = pending_option_images.pop(0)
        option.image_path = image_path
        option.source_page = region.page_number if region else None
        option.page_region = region

    for rich_line in option_lines:
        text = _rich_line_text(rich_line)
        images = [
            (
                image_path,
                _region_from_extracted_image((image_regions or {}).get(image_path)),
            )
            for image_path in _rich_line_images(rich_line)
        ]
        if images:
            if text:
                pending_option_images.extend(images)
            elif current is not None and not current.image_path and not pending_option_images:
                image_path, region = images[0]
                current.image_path = image_path
                current.source_page = region.page_number if region else None
                current.page_region = region
                pending_option_images.extend(images[1:])
            else:
                pending_option_images.extend(images)
        if text:
            fragments = _split_option_fragments(text)
            if fragments:
                attach_pending_image(current)
                if current is not None:
                    options.append(current)
                for letter, body in fragments[:-1]:
                    option = OptionNode(letter=letter, text=body)
                    attach_pending_image(option)
                    options.append(option)
                last_letter, last_body = fragments[-1]
                current = OptionNode(letter=last_letter, text=last_body)
                attach_pending_image(current)
            elif current is not None:
                if (
                    current.letter == "D"
                    and len(options) >= 3
                    and _looks_like_following_passage_line(text)
                ):
                    break
                current.text = _join_wrapped_lines([current.text, text])
                attach_pending_image(current)
    attach_pending_image(current)
    if current is not None:
        options.append(current)
    options = _normalize_option_letters(options)
    stem_assets, options = _rebalance_image_only_options_from_stem_assets(stem_assets, options)

    stem = _join_stem_lines(stem_text_lines)
    question = QuestionNode(
        source_number=(source_number or "").strip(),
        stem=stem,
        options=options,
        stem_assets=stem_assets,
        page_numbers=page_numbers,
    )
    question = _repair_blank_image_options(question)
    question = _assign_option_row_stem_assets(question, marker_lines)
    question = _fill_blank_option_images_from_pdf_crops(question, marker_lines, page_cropper)
    question = _fill_missing_stem_from_pdf_crop(
        question,
        source_number=source_number,
        locator=locator,
        marker_lines=marker_lines,
        cropper=page_cropper,
    )
    question = _synthesize_missing_options(question)
    question = _upgrade_placeholder_text_options_from_stem_asset(question)
    question = _refresh_option_images_from_pdf_crop(question, page_cropper)
    return question


def _append_objective_project_section(
    project: ExamProject,
    parsed_section: ObjectiveSection,
    locator: _LayoutLocator,
    image_regions: Mapping[str, ExtractedImageRegion] | None = None,
    page_cropper: _PdfRenderCropper | None = None,
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
                page_cropper=page_cropper,
            )
        )
    _promote_trailing_option_image_to_next_stem(section.questions)
    _promote_leading_option_image_to_previous_question(section.questions)
    _promote_trailing_stem_asset_to_previous_question(section.questions)
    if not section.questions:
        return
    last_section = project.sections[-1] if project.sections else None
    if (
        last_section is not None
        and last_section.kind == section.kind
        and should_merge_subject_sections(section.kind, last_section.title, section.title)
    ):
        last_section.title = preferred_subject_title(section.kind, last_section.title, section.title)
        last_section.questions.extend(section.questions)
        return
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
    page_cropper = _PdfRenderCropper(source_pdf_path)
    project = ExamProject(
        title=project_title,
        source=PaperSource(pdf_path=source_pdf_path),
    )
    try:
        for parsed_section in exam.iter_objective_sections():
            _append_objective_project_section(
                project,
                parsed_section,
                locator,
                image_regions=image_regions,
                page_cropper=page_cropper,
            )

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
                question_line_groups: list[list[PageTextLine]] = []
                for parsed_question in parsed_material.questions:
                    question_stem_lines = [_rich_line_text(line) for line in parsed_question.stem_lines if _rich_line_text(line)]
                    question_located_lines = locator.consume(question_stem_lines)
                    question_line_groups.append(question_located_lines)
                    material.questions.append(
                        _question_from_rich(
                            parsed_question.source_number,
                            parsed_question.stem_lines,
                            parsed_question.option_lines,
                            locator,
                            image_regions=image_regions,
                            located_lines=question_located_lines,
                            page_cropper=page_cropper,
                        )
                    )
                _promote_trailing_option_image_to_next_stem(material.questions)
                _promote_leading_option_image_to_previous_question(material.questions)
                _promote_trailing_stem_asset_to_previous_question(material.questions)
                _promote_leading_question_assets_to_material(material)
                _attach_orphan_material_images(
                    material,
                    question_line_groups=question_line_groups,
                    image_regions=image_regions,
                )
                if material.questions:
                    section.material_sets.append(material)
            if section.material_sets:
                project.sections.append(section)

        annotate_project_quality(project)
        return project
    finally:
        page_cropper.close()


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
            source_name=os.path.basename(pdf_path),
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
