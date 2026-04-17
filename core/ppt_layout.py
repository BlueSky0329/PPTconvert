from __future__ import annotations

from typing import Mapping

SLIDE_WIDTH_IN = 13.333
SLIDE_HEIGHT_IN = 7.5
LAYOUT_BLOCK_ORDER = ("stem", "image", "options")
OPTION_BLOCK_PREFIX = "option_"
_DEFAULT_BOTTOM_PADDING_IN = 0.18
_MIN_LAYOUT_BLOCK_SIZE = {
    "stem": {"w": 0.18, "h": 0.10},
    "image": {"w": 0.12, "h": 0.10},
    "options": {"w": 0.20, "h": 0.16},
    "option_item": {"w": 0.08, "h": 0.06},
}


def _clamp(value: float, lo: float, hi: float) -> float:
    return max(lo, min(hi, value))


def _as_float(value: object, default: float = 0.0) -> float:
    try:
        return float(value)
    except (TypeError, ValueError):
        return default


def _inches(value: object, default: float = 0.0) -> float:
    inches = getattr(value, "inches", None)
    if inches is not None:
        return _as_float(inches, default)
    return _as_float(value, default)


def option_layout_block_key(letter: str) -> str:
    normalized = (letter or "").strip().upper() or "A"
    return f"{OPTION_BLOCK_PREFIX}{normalized.lower()}"


def is_option_layout_block(block: str | None) -> bool:
    normalized = (block or "").strip().lower()
    return normalized.startswith(OPTION_BLOCK_PREFIX) and len(normalized) > len(OPTION_BLOCK_PREFIX)


def option_layout_block_letter(block: str | None) -> str:
    normalized = (block or "").strip()
    if not is_option_layout_block(normalized):
        return ""
    return normalized[len(OPTION_BLOCK_PREFIX) :].upper()


def _rect_min_size(block: str) -> dict[str, float]:
    normalized = (block or "").strip().lower()
    if normalized in _MIN_LAYOUT_BLOCK_SIZE:
        return _MIN_LAYOUT_BLOCK_SIZE[normalized]
    if is_option_layout_block(normalized):
        return _MIN_LAYOUT_BLOCK_SIZE["option_item"]
    return _MIN_LAYOUT_BLOCK_SIZE["option_item"]


def _sorted_layout_blocks(layout: Mapping[str, Mapping[str, object]]) -> list[str]:
    option_blocks = sorted(
        block for block in layout.keys() if is_option_layout_block(block)
    )
    ordered = [block for block in LAYOUT_BLOCK_ORDER if block in layout]
    ordered.extend(option_blocks)
    return ordered


def sanitize_ppt_layout(
    layout: Mapping[str, Mapping[str, object]] | None,
) -> dict[str, dict[str, float]]:
    if not layout:
        return {}

    normalized: dict[str, dict[str, float]] = {}
    for block in _sorted_layout_blocks(layout):
        raw_rect = layout.get(block)
        if not raw_rect:
            continue
        normalized_block = (block or "").strip().lower()
        min_size = _rect_min_size(normalized_block)
        raw_x = _clamp(_as_float(raw_rect.get("x"), 0.0), 0.0, 1.0)
        raw_y = _clamp(_as_float(raw_rect.get("y"), 0.0), 0.0, 1.0)
        max_w = max(min_size["w"], 1.0 - raw_x)
        max_h = max(min_size["h"], 1.0 - raw_y)
        raw_w = _clamp(_as_float(raw_rect.get("w"), min_size["w"]), min_size["w"], max_w)
        raw_h = _clamp(_as_float(raw_rect.get("h"), min_size["h"]), min_size["h"], max_h)
        normalized[normalized_block] = {
            "x": round(raw_x, 6),
            "y": round(raw_y, 6),
            "w": round(raw_w, 6),
            "h": round(raw_h, 6),
        }
    return normalized


def merge_ppt_layout(
    base_layout: Mapping[str, Mapping[str, object]] | None,
    overrides: Mapping[str, Mapping[str, object]] | None,
) -> dict[str, dict[str, float]]:
    merged = sanitize_ppt_layout(base_layout)
    for block, rect in sanitize_ppt_layout(overrides).items():
        merged[block] = rect
    return merged


def _layout_kind(question, config) -> str:
    layout = (getattr(question, "option_layout", None) or getattr(config, "option_layout", "grid") or "grid").strip().lower()
    return layout if layout in {"grid", "list", "one_row"} else "grid"


def _default_option_item_layouts(question, config, options_rect: Mapping[str, float]) -> dict[str, dict[str, float]]:
    option_count = min(4, len(getattr(question, "options", []) or []))
    if option_count <= 0:
        return {}

    x = _as_float(options_rect.get("x"), 0.0) * SLIDE_WIDTH_IN
    y = _as_float(options_rect.get("y"), 0.0) * SLIDE_HEIGHT_IN
    w = _as_float(options_rect.get("w"), 0.0) * SLIDE_WIDTH_IN
    h = _as_float(options_rect.get("h"), 0.0) * SLIDE_HEIGHT_IN
    layout_kind = _layout_kind(question, config)
    option_rects: dict[str, dict[str, float]] = {}
    option_letters = [
        (getattr(option, "letter", "") or chr(ord("A") + index)).strip().upper() or chr(ord("A") + index)
        for index, option in enumerate((getattr(question, "options", []) or [])[:option_count])
    ]

    if layout_kind == "one_row":
        gap = max(0.0, _as_float(getattr(config, "one_row_gap_in", 0.06), 0.06))
        row_height = min(h, max(0.36, _as_float(getattr(config, "one_row_height_in", 0.55), 0.55)))
        cell_width = (w - gap * (option_count - 1)) / max(1, option_count)
        row_top = y + max(0.0, (h - row_height) / 2.0)
        for index, letter in enumerate(option_letters):
            item_x = x + index * (cell_width + gap)
            option_rects[option_layout_block_key(letter)] = normalize_layout_rect(
                item_x,
                row_top,
                item_x + max(0.2, cell_width - 0.04),
                row_top + max(0.2, row_height - 0.04),
                width=SLIDE_WIDTH_IN,
                height=SLIDE_HEIGHT_IN,
            )
        return option_rects

    if layout_kind == "list":
        gap = max(0.0, 0.06)
        row_height = max(0.4, _as_float(getattr(config, "list_row_height_in", 0.7), 0.7))
        needed_height = row_height * option_count + gap * max(0, option_count - 1)
        shrink = min(1.0, h / needed_height) if needed_height else 1.0
        row_height *= shrink
        gap *= shrink
        for index, letter in enumerate(option_letters):
            item_top = y + index * (row_height + gap)
            option_rects[option_layout_block_key(letter)] = normalize_layout_rect(
                x,
                item_top,
                x + w,
                item_top + max(0.2, row_height - 0.1),
                width=SLIDE_WIDTH_IN,
                height=SLIDE_HEIGHT_IN,
            )
        return option_rects

    gap = max(0.0, _as_float(getattr(config, "grid_col_gap_in", 0.15), 0.15))
    col_width = (w - gap) / 2.0
    row_height = max(0.45, _as_float(getattr(config, "grid_row_height_in", 0.9), 0.9))
    needed_height = row_height * 2.0 + gap
    shrink = min(1.0, h / needed_height) if needed_height else 1.0
    row_height *= shrink
    gap *= shrink
    grid_layout = (getattr(config, "grid_layout", "ab_cd") or "ab_cd").strip().lower()
    order = (
        [("A", 0, 0), ("C", 0, 1), ("B", 1, 0), ("D", 1, 1)]
        if grid_layout == "ac_bd"
        else [("A", 0, 0), ("B", 1, 0), ("C", 0, 1), ("D", 1, 1)]
    )
    render_set = set(option_letters)
    for letter, grid_x, grid_y in order:
        if letter not in render_set:
            continue
        item_x = x + grid_x * (col_width + gap)
        item_y = y + grid_y * (row_height + gap)
        option_rects[option_layout_block_key(letter)] = normalize_layout_rect(
            item_x,
            item_y,
            item_x + max(0.2, col_width - 0.3),
            item_y + max(0.2, row_height - 0.1),
            width=SLIDE_WIDTH_IN,
            height=SLIDE_HEIGHT_IN,
        )
    return option_rects


def build_default_question_layout(question, config) -> dict[str, dict[str, float]]:
    margin_left = _as_float(getattr(config, "margin_left_in", 0.8), 0.8)
    margin_right = _as_float(getattr(config, "margin_right_in", 0.8), 0.8)
    margin_top = _as_float(getattr(config, "margin_top_in", 0.5), 0.5)
    content_width = max(2.0, SLIDE_WIDTH_IN - margin_left - margin_right)
    has_image = bool(getattr(question, "image_paths", None))
    stem_height = _as_float(
        getattr(config, "stem_height_with_image_in", 1.5)
        if has_image
        else getattr(config, "stem_height_no_image_in", 2.5),
        1.5 if has_image else 2.5,
    )
    stem_rect = {
        "x": margin_left / SLIDE_WIDTH_IN,
        "y": margin_top / SLIDE_HEIGHT_IN,
        "w": content_width / SLIDE_WIDTH_IN,
        "h": stem_height / SLIDE_HEIGHT_IN,
    }

    cursor_y = margin_top + stem_height + _as_float(getattr(config, "gap_after_stem_in", 0.2), 0.2)
    default_layout: dict[str, dict[str, float]] = {"stem": stem_rect}

    if has_image:
        image_width = min(
            content_width,
            _inches(getattr(config, "image_max_width", content_width), content_width) or content_width,
        )
        image_height = min(
            max(0.6, SLIDE_HEIGHT_IN * 0.22),
            _inches(getattr(config, "image_max_height", 2.5), 2.5) or 2.5,
        )
        image_align = (getattr(config, "image_h_align", "center") or "center").strip().lower()
        if image_align == "right":
            image_left = margin_left + content_width - image_width
        elif image_align == "left":
            image_left = margin_left
        else:
            image_left = margin_left + max(0.0, (content_width - image_width) / 2.0)
        default_layout["image"] = {
            "x": image_left / SLIDE_WIDTH_IN,
            "y": cursor_y / SLIDE_HEIGHT_IN,
            "w": image_width / SLIDE_WIDTH_IN,
            "h": image_height / SLIDE_HEIGHT_IN,
        }
        cursor_y += image_height + _as_float(getattr(config, "gap_after_image_in", 0.15), 0.15)

    options_top = cursor_y + _as_float(getattr(config, "gap_before_options_in", 0.2), 0.2)
    options_height = max(0.9, SLIDE_HEIGHT_IN - options_top - _DEFAULT_BOTTOM_PADDING_IN)
    default_layout["options"] = {
        "x": margin_left / SLIDE_WIDTH_IN,
        "y": options_top / SLIDE_HEIGHT_IN,
        "w": content_width / SLIDE_WIDTH_IN,
        "h": options_height / SLIDE_HEIGHT_IN,
    }
    return sanitize_ppt_layout(default_layout)


def build_effective_question_layout(question, config) -> dict[str, dict[str, float]]:
    default_layout = build_default_question_layout(question, config)
    overrides = getattr(question, "ppt_layout", None) or {}
    main_overrides = {
        block: rect
        for block, rect in overrides.items()
        if not is_option_layout_block(block)
    }
    option_overrides = {
        block: rect
        for block, rect in overrides.items()
        if is_option_layout_block(block)
    }
    effective = merge_ppt_layout(default_layout, main_overrides)
    if not bool(getattr(question, "image_paths", None)):
        effective.pop("image", None)
    options_rect = effective.get("options")
    if options_rect:
        effective.update(_default_option_item_layouts(question, config, options_rect))
    effective = merge_ppt_layout(effective, option_overrides)
    return effective


def scale_layout_rect(
    rect: Mapping[str, object],
    width: float,
    height: float,
    *,
    offset_x: float = 0.0,
    offset_y: float = 0.0,
) -> tuple[float, float, float, float]:
    x = offset_x + _as_float(rect.get("x"), 0.0) * width
    y = offset_y + _as_float(rect.get("y"), 0.0) * height
    w = _as_float(rect.get("w"), 0.0) * width
    h = _as_float(rect.get("h"), 0.0) * height
    return x, y, x + w, y + h


def normalize_layout_rect(
    x0: float,
    y0: float,
    x1: float,
    y1: float,
    *,
    origin_x: float = 0.0,
    origin_y: float = 0.0,
    width: float = 1.0,
    height: float = 1.0,
) -> dict[str, float]:
    usable_width = max(1e-6, float(width))
    usable_height = max(1e-6, float(height))
    raw_x = _clamp((x0 - origin_x) / usable_width, 0.0, 1.0)
    raw_y = _clamp((y0 - origin_y) / usable_height, 0.0, 1.0)
    raw_w = _clamp((x1 - x0) / usable_width, 0.0, 1.0 - raw_x)
    raw_h = _clamp((y1 - y0) / usable_height, 0.0, 1.0 - raw_y)
    return {
        "x": round(raw_x, 6),
        "y": round(raw_y, 6),
        "w": round(raw_w, 6),
        "h": round(raw_h, 6),
    }
