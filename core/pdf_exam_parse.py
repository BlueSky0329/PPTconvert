"""将 PDF 抽取的行序列解析为公考六大模块结构。"""

from __future__ import annotations

import re
import unicodedata
from typing import Iterable, Literal

from core.pdf_exam_models import (
    CommonSenseSection,
    DataAnalysisSection,
    ExamOption,
    ExamQuestion,
    MaterialUnit,
    ParsedExam,
    PoliticsSection,
    QuantSection,
    ReasoningSection,
    RichLine,
    VerbalSection,
)
from domain.models import ALL_SUBJECT_KINDS, SubjectKind
from core.word_parser import OPTION_MARKER, _parse_options_from_line

LineItem = tuple[str, str | None]  # (text, image_path)


def _has_option_markers(text: str) -> bool:
    return bool(list(OPTION_MARKER.finditer(text)))


def _normalize_pdf_option_text(text: str) -> str:
    """
    PDF 常提取为全角拉丁字母 Ａ-Ｚ，而 OPTION_MARKER 只认半角 A-Z。
    """
    out: list[str] = []
    for ch in text:
        o = ord(ch)
        if 0xFF21 <= o <= 0xFF3A:  # Ａ-Ｚ
            out.append(chr(o - 0xFF21 + ord("A")))
        elif 0xFF41 <= o <= 0xFF5A:  # ａ-ｚ
            out.append(chr(o - 0xFF41 + ord("a")))
        else:
            out.append(ch)
    return "".join(out)


def _has_option_markers_pdf(text: str) -> bool:
    if not (text or "").strip():
        return False
    return bool(list(OPTION_MARKER.finditer(_normalize_pdf_option_text(text))))


# 竖排选项：每行 A. xxx
_VERTICAL_OPTION = re.compile(
    r"^\s*([ABCDabcd])\s*[\.．、:：]\s*(.+)$",
)
_SINGLE_OPTION = re.compile(r"^\s*([ABCDabcd])\s*[\.．、:：)\uFF09]\s*(.*)$")
_QUESTION_NO_LINE = re.compile(r"^\s*\d{1,3}\s*[\.．、]?\s*$")
_LEADING_QUESTION_LINE = re.compile(
    r"^\s*(?P<number>\d{1,3})\s*[\.．、)\uFF09]\s*(?P<stem>[（(【\[]?[A-Za-z\u4e00-\u9fff].*)$"
)
_LEADING_CN_QUESTION_LINE = re.compile(
    r"^\s*第\s*(?P<number>\d{1,3})\s*题(?:\s*[\uff1a:.\uFF0E\u3001)\uFF09]\s*)?(?P<stem>[（(【\[]?[A-Za-z\u4e00-\u9fff].*)$"
)
_EMBEDDED_QUESTION_TRANSITION = re.compile(
    r"^(?P<head>.*?\S)\s+(?P<number>\d{1,3})\s*[\.．、)\uFF09]\s*(?P<stem>[（(【\[]?[A-Za-z\u4e00-\u9fff].*)$"
)

MATERIAL_HEADER = re.compile(r"^\s*材料[一二三四五六七八九十百千万\d〇零两]+\s*$")


def material_header_line(line: str) -> bool:
    """比严格整行匹配更宽松，兼容「材料一、」「【材料一】」「材料 1」等。"""
    s = _nfkc((line or "").strip())
    if not s:
        return False
    s = s.strip("【】[]［］")
    s = s.strip()
    if re.match(r"^材料\s*[一二三四五六七八九十百千万\d〇零两]+", s):
        return True
    if re.match(r"^材料\s*[1-9１-９]", s):
        return True
    return bool(MATERIAL_HEADER.match(s))

# 题干里常出现的「资料分析」短语，应排除，避免误判为篇题
_TITLE_FALSE_POSITIVE = re.compile(
    r"^(根据|由|从|结合|阅读|下列|关于|对于|若|如|由此|以下|能够|不能|推出|说法)"
)

_SECTION_LABELS: dict[SubjectKind, tuple[str, ...]] = {
    "politics": ("政治理论",),
    "common_sense": ("常识判断",),
    "verbal": ("言语理解与表达", "言语理解和表达"),
    "quant": ("数量关系",),
    "reasoning": ("判断推理",),
    "data": ("资料分析",),
}


def _nfkc(s: str) -> str:
    """兼容区字符（如 ⼀）归一成常规汉字/标点。"""
    return unicodedata.normalize("NFKC", s or "")


def _normalize_digits(s: str) -> str:
    """全角数字转半角，便于匹配年份。"""
    return s.translate(str.maketrans("０１２３４５６７８９", "0123456789"))


_CN_NUM = "一二三四五六七八九十"
_OUTLINE_SECTION_TITLE = re.compile(
    rf"^\s*(?:第[一二三四五六七八九十百\d〇零]+部分|[{_CN_NUM}\d]{{1,3}}[、．。.])\s*[\u4e00-\u9fff]+"
)


def _is_boilerplate_line(line: str) -> bool:
    """篇首说明（非题目），需跳过。"""
    s = (line or "").strip()
    if s in ("案。", "答案。"):
        return True
    hints = (
        "本部分包括表达与理解",
        "在四个选项中选出一个最",
        "恰当的答案",
        "在这部分试题中",
        "所给出的图",
        "表、文字或综合性资料",
        "你应根据资料提供的信息",
        "要求你迅速",
        "准确地计算出答案",
        "最恰当的答案",
        "分析、比较、计算和判断",
    )
    if any(h in s for h in hints):
        return True
    if len(s) < 10:
        return False
    return False


def _skip_section_boilerplate(items: list[LineItem], start: int, end: int, *, kind: SubjectKind) -> int:
    i = start
    saw_boilerplate = False
    while i < end:
        t, img = items[i]
        if img:
            return i
        line = _nfkc((t or "").strip())
        if not line:
            i += 1
            continue
        if _is_boilerplate_line(line):
            saw_boilerplate = True
            i += 1
            continue
        if saw_boilerplate and kind != "data":
            if _starts_new_question_line(line):
                return i
            if _detect_subject_section_kind(line) or _is_other_section_title(line):
                return i
            i += 1
            continue
        return i
    return start


def _has_four_digit_year(s: str) -> bool:
    """行内是否含四位年份（含全角数字归一后）。"""
    s = _normalize_digits(s)
    return bool(re.search(r"(?<!\d)\d{4}(?!\d)", s))


def _is_question_no_line(line: str) -> bool:
    return bool(_QUESTION_NO_LINE.match(_nfkc(_normalize_digits((line or "").strip()))))


def _extract_question_no(line: str) -> str:
    s = _nfkc(_normalize_digits((line or "").strip()))
    m = re.match(r"^\s*(\d{1,3})", s)
    return m.group(1) if m else ""


def _match_leading_question_with_stem(line: str) -> tuple[str, str] | None:
    s = _nfkc(_normalize_digits((line or "").strip()))
    for pattern in (_LEADING_QUESTION_LINE, _LEADING_CN_QUESTION_LINE):
        match = pattern.match(s)
        if match:
            number = (match.group("number") or "").strip()
            stem = (match.group("stem") or "").strip()
            if number and stem:
                return number, stem
    return None


def _starts_new_question_line(line: str) -> bool:
    return _is_question_no_line(line) or _match_leading_question_with_stem(line) is not None


def _is_other_section_title(line: str) -> bool:
    s = _nfkc(_normalize_digits((line or "").strip()))
    if not s:
        return False
    if _detect_subject_section_kind(s):
        return False
    if _TITLE_FALSE_POSITIVE.search(s):
        return False
    if len(s) > 40:
        return False
    return bool(_OUTLINE_SECTION_TITLE.match(s))


def _section_labels(kind: SubjectKind) -> tuple[str, ...]:
    return _SECTION_LABELS.get(kind, ())


def _subject_kind_in_text(text: str) -> SubjectKind | None:
    for kind, labels in _SECTION_LABELS.items():
        if any(label in text for label in labels):
            return kind
    return None


def _is_subject_section_title(line: str, kind: SubjectKind) -> bool:
    s = _nfkc(_normalize_digits((line or "").strip()))
    labels = _section_labels(kind)
    if not s or not any(label in s for label in labels):
        return False
    if _TITLE_FALSE_POSITIVE.search(s):
        return False
    if _has_four_digit_year(s):
        return True
    if re.search(r"第[一二三四五六七八九十百\d〇零]+部分", s):
        return True
    label_pattern = "|".join(re.escape(label) for label in labels)
    if re.match(rf"^[{_CN_NUM}\d]+\s*[、．。.]?\s*(?:{label_pattern})", s):
        return True
    if re.match(rf"^[（(]\s*[一二三四五六七八九十\d]+\s*[）)]\s*(?:{label_pattern})", s):
        return True
    if len(s) <= 36:
        core = re.sub(r"[\s「」〖〗\[\]（）():：共\d题项道\s]+", "", s)
        if any(core in (label, label + "题") for label in labels):
            return True
    return False


def _detect_subject_section_kind(line: str) -> SubjectKind | None:
    for kind in ALL_SUBJECT_KINDS:
        if _is_subject_section_title(line, kind):
            return kind
    return None


def _is_data_section_title(line: str) -> bool:
    return _is_subject_section_title(line, "data")


def _is_quant_section_title(line: str) -> bool:
    return _is_subject_section_title(line, "quant")


def _pair_cn_section(items: list[LineItem], i: int) -> tuple[SubjectKind, str, int] | None:
    """
    识别「四.」与「数量关系：…」被拆成两行的篇题（无年份的公考大纲式）。
    """
    if i + 1 >= len(items):
        return None
    t0, im0 = items[i]
    t1, im1 = items[i + 1]
    if im0 or im1:
        return None
    a = _nfkc(_normalize_digits((t0 or "").strip()))
    b = _nfkc(_normalize_digits((t1 or "").strip()))
    if not a or not b:
        return None
    if not re.match(rf"^[{_CN_NUM}]+[\.．、]?\s*$", a) or len(a) > 8:
        return None
    kind = _subject_kind_in_text(b)
    if kind:
        return (kind, a + b, i + 2)
    return None


def _pair_section_title(items: list[LineItem], i: int) -> tuple[SubjectKind, str, int] | None:
    """
    识别「年份/地区行」与下一行「资料分析/数量关系」被 PDF 拆成两行的篇题。
    返回 (kind, merged_title, end_index)。
    """
    if i + 1 >= len(items):
        return None
    t0, im0 = items[i]
    t1, im1 = items[i + 1]
    if im0 or im1:
        return None
    a = _normalize_digits((t0 or "").strip())
    b = (t1 or "").strip()
    if not a or not b:
        return None
    if not re.match(r"^\d{4}", a):
        return None
    if len(a) > 56 or len(b) > 48:
        return None
    if _subject_kind_in_text(a):
        return None
    kind = _subject_kind_in_text(b)
    if kind:
        return (kind, a + b, i + 2)
    return None


def _rich_text(s: str) -> RichLine:
    return RichLine(parts=[(s, None)])


def _rich_img(path: str) -> RichLine:
    return RichLine(parts=[("", path)])


def _line_has_text(rich: RichLine) -> bool:
    return any((text or "").strip() for text, _img in rich.parts)


def _rich_line_text_value(rich: RichLine) -> str:
    return "".join((text or "") for text, _img in rich.parts).strip()


def _rich_line_has_image(rich: RichLine) -> bool:
    return any(img for _text, img in rich.parts)


def _rich_option_letter(rich: RichLine) -> str | None:
    text = _normalize_pdf_option_text(_rich_line_text_value(rich))
    if not text:
        return None
    match = _SINGLE_OPTION.match(text)
    if not match:
        return None
    return match.group(1).upper()


def _looks_like_material_intro_lines(lines: list[RichLine]) -> bool:
    meaningful = [line for line in lines if _line_has_text(line) or _rich_line_has_image(line)]
    if not meaningful:
        return False

    texts = [_rich_line_text_value(line) for line in meaningful if _line_has_text(line)]
    image_count = sum(1 for line in meaningful if _rich_line_has_image(line))
    total_chars = sum(len(text) for text in texts)
    long_lines = sum(1 for text in texts if len(text) >= 18)

    if image_count >= 2:
        return True
    if image_count >= 1 and total_chars >= 18:
        return True
    if len(texts) >= 3:
        return True
    if len(texts) >= 2 and (total_chars >= 28 or long_lines >= 1):
        return True
    return False


def _split_material_intro_from_option_lines(lines: list[RichLine]) -> tuple[list[RichLine], list[RichLine]]:
    letters = "ABCD"
    expected_index = 0
    last_option_index = -1

    for idx, line in enumerate(lines):
        letter = _rich_option_letter(line)
        if expected_index < len(letters) and letter == letters[expected_index]:
            last_option_index = idx
            expected_index += 1

    if expected_index < len(letters) or last_option_index < 0 or last_option_index >= len(lines) - 1:
        return list(lines), []

    spill = list(lines[last_option_index + 1 :])
    if not _looks_like_material_intro_lines(spill):
        return list(lines), []
    return list(lines[: last_option_index + 1]), spill


def _split_rich_intro_stem(lines: list[RichLine]) -> tuple[list[RichLine], list[RichLine]]:
    text_positions = [i for i, line in enumerate(lines) if _line_has_text(line)]
    if len(text_positions) <= 1:
        return [], list(lines)
    stem_start = text_positions[-1]
    return list(lines[:stem_start]), list(lines[stem_start:])


def _extract_question_number_and_strip(
    rich_lines: list[RichLine],
) -> tuple[str, list[RichLine]]:
    number = ""
    cleaned: list[RichLine] = []
    for line in rich_lines:
        text = "".join(part or "" for part, _img in line.parts).strip()
        if not number and text and _is_question_no_line(text):
            number = _extract_question_no(text)
            continue
        cleaned.append(line)
    return number, cleaned


def _parse_options_line(text: str) -> list[ExamOption] | None:
    raw = _parse_options_from_line(_normalize_pdf_option_text(text))
    if len(raw) < 2:
        return None
    return [ExamOption(letter=o.letter.upper(), text=o.text) for o in raw]


def _try_vertical_four_options(
    items: list[LineItem],
    start: int,
    end: int,
) -> tuple[int, list[RichLine]] | None:
    """每行仅 A. xxx 形式、共四行时的选项块。"""

    def text_at(idx: int) -> str:
        t, img = items[idx]
        if img:
            return ""
        return (t or "").strip()

    if start + 3 >= end:
        return None
    merged_parts: list[str] = []
    for k, want in enumerate("ABCD"):
        raw = text_at(start + k)
        if not raw:
            return None
        n = _normalize_pdf_option_text(raw)
        m = _VERTICAL_OPTION.match(n)
        if not m:
            return None
        if m.group(1).upper() != want:
            return None
        merged_parts.append(n)
    merged = "\t".join(merged_parts)
    opts = _parse_options_line(merged)
    if opts and len(opts) >= 4:
        return start + 4, _options_to_rich_lines(opts[:4])
    return None


def _collect_sequential_option_lines(
    items: list[LineItem],
    start: int,
    end: int,
) -> tuple[int, list[RichLine]] | None:
    i = start
    expected_index = 0
    out: list[RichLine] = []
    letters = "ABCD"

    while i < end and expected_index < len(letters):
        t, img = items[i]
        if img:
            if not out:
                return None
            out.append(_rich_img(img))
            i += 1
            continue

        raw = (t or "").strip()
        if not raw:
            i += 1
            continue

        normalized = _normalize_pdf_option_text(raw)
        match = _SINGLE_OPTION.match(normalized)
        want = letters[expected_index]
        if not match or match.group(1).upper() != want:
            return None

        body = match.group(2).strip()
        out.append(_rich_text(f"{want}．{body}" if body else f"{want}．"))
        expected_index += 1
        i += 1

        while i < end:
            next_text, next_img = items[i]
            if next_img:
                out.append(_rich_img(next_img))
                i += 1
                continue

            next_raw = (next_text or "").strip()
            if not next_raw:
                i += 1
                continue

            next_normalized = _normalize_pdf_option_text(next_raw)
            next_match = _SINGLE_OPTION.match(next_normalized)
            if expected_index < len(letters) and next_match and next_match.group(1).upper() == letters[expected_index]:
                break
            if _starts_new_question_line(next_normalized) or material_header_line(next_normalized):
                break
            if _detect_subject_section_kind(next_normalized) or _is_other_section_title(next_normalized):
                break
            out.append(_rich_text(next_raw))
            i += 1

    if expected_index == len(letters):
        return i, out
    return None


def _collect_accumulated_option_lines(
    items: list[LineItem],
    start: int,
    end: int,
) -> tuple[int, list[RichLine]] | None:
    i = start
    parts: list[str] = []
    consumed_any = False

    while i < end:
        text, image_path = items[i]
        if image_path:
            break

        raw = (text or "").strip()
        if not raw:
            i += 1
            continue

        normalized = _normalize_pdf_option_text(raw)
        if consumed_any and (
            _starts_new_question_line(normalized)
            or material_header_line(normalized)
            or _detect_subject_section_kind(normalized)
            or _is_other_section_title(normalized)
        ):
            break
        if consumed_any and not _has_option_markers_pdf(normalized):
            break
        if not consumed_any and not _has_option_markers_pdf(normalized):
            return None

        parts.append(normalized)
        consumed_any = True
        combined = "\t".join(parts)
        options = _parse_options_line(combined)
        if options and len(options) >= 4:
            return i + 1, _options_to_rich_lines(options[:4])
        i += 1

    return None


def _option_cluster_end(
    items: list[LineItem],
    start: int,
    end: int,
) -> tuple[int, list[RichLine]] | None:
    if start >= end or start >= len(items):
        return None

    def text_at(idx: int) -> str:
        t, img = items[idx]
        if img:
            return ""
        return (t or "").strip()

    first = text_at(start)
    if not first:
        return None

    # 避免把题干与下一行选项拼在一起误判为「四个选项」
    if not _has_option_markers_pdf(first):
        return None

    opts = _parse_options_line(first)
    if opts and len(opts) >= 4:
        return start + 1, _options_to_rich_lines(opts[:4])

    if opts and len(opts) == 2 and start + 1 < end:
        second = text_at(start + 1)
        if not second or not _has_option_markers_pdf(second):
            return None
        merged = _normalize_pdf_option_text(first) + "\t" + _normalize_pdf_option_text(second)
        opts2 = _parse_options_line(merged)
        if opts2 and len(opts2) >= 4:
            return start + 2, _options_to_rich_lines(opts2[:4])

    seq = _collect_sequential_option_lines(items, start, end)
    if seq:
        return seq

    if opts and 1 < len(opts) < 4:
        accumulated = _collect_accumulated_option_lines(items, start, end)
        if accumulated:
            return accumulated

    return None


def _split_intro_stem(segment: list[LineItem]) -> tuple[list[RichLine], list[RichLine]]:
    """材料下第一小题：前文为材料，最后一行文字为题干。"""
    text_positions = [
        i
        for i, it in enumerate(segment)
        if (it[0] or "").strip() and not it[1] and not _is_question_no_line(it[0] or "")
    ]
    if not text_positions:
        rich: list[RichLine] = []
        for it in segment:
            if it[1]:
                rich.append(_rich_img(it[1]))
        return rich, []

    last_t = text_positions[-1]
    intro: list[RichLine] = []
    stem: list[RichLine] = []

    for i, it in enumerate(segment):
        t, img = it
        if img:
            if i < last_t:
                intro.append(_rich_img(img))
            else:
                stem.append(_rich_img(img))
            continue
        if not (t or "").strip() or _is_question_no_line(t):
            continue
        if i < last_t:
            intro.append(_rich_text(t))
        elif i == last_t:
            stem.append(_rich_text(t))
        else:
            stem.append(_rich_text(t))

    return intro, stem


def _segment_to_stem_only(segment: list[LineItem]) -> list[RichLine]:
    out: list[RichLine] = []
    for t, img in segment:
        if img:
            out.append(_rich_img(img))
            continue
        stripped = (t or "").strip()
        if not stripped:
            continue
        matched = _match_leading_question_with_stem(stripped)
        if matched:
            _number, stem = matched
            out.append(_rich_text(stem))
        elif not _is_question_no_line(stripped):
            out.append(_rich_text(stripped))
    return out


def _extract_source_number_from_segment(segment: list[LineItem]) -> str:
    for text, image_path in segment:
        if image_path:
            continue
        stripped = (text or "").strip()
        if not stripped:
            continue
        if _is_question_no_line(stripped):
            return _extract_question_no(stripped)
        matched = _match_leading_question_with_stem(stripped)
        if matched:
            return matched[0]
    return ""


def _options_to_rich_lines(options: list[ExamOption]) -> list[RichLine]:
    if not options:
        return []
    letters = [o.letter for o in options]
    texts = [o.text for o in options]
    sep = "\t\t"
    if len(options) == 4:
        line1 = f"{letters[0]}．{texts[0]}{sep}{letters[1]}．{texts[1]}"
        line2 = f"{letters[2]}．{texts[2]}{sep}{letters[3]}．{texts[3]}"
        return [_rich_text(line1), _rich_text(line2)]
    line = "\t".join(f"{letters[i]}．{texts[i]}" for i in range(len(options)))
    return [_rich_text(line)]


def _split_embedded_question_transition(text: str) -> list[str]:
    normalized = _nfkc(_normalize_digits((text or "").strip()))
    if not normalized:
        return []

    pieces = [normalized]
    while True:
        updated: list[str] = []
        changed = False
        for piece in pieces:
            match = _EMBEDDED_QUESTION_TRANSITION.match(piece)
            if not match:
                updated.append(piece)
                continue
            head = (match.group("head") or "").strip()
            number = (match.group("number") or "").strip()
            stem = (match.group("stem") or "").strip()
            if head:
                updated.append(head)
            updated.append(f"{number}.")
            if stem:
                updated.append(stem)
            changed = True
        pieces = updated
        if not changed:
            break

    expanded: list[str] = []
    for piece in pieces:
        matched = _match_leading_question_with_stem(piece)
        if not matched:
            expanded.append(piece)
            continue
        number, stem = matched
        expanded.append(f"{number}.")
        expanded.append(stem)
    return [piece for piece in expanded if piece]


def _preprocess_line_items(items: list[LineItem]) -> list[LineItem]:
    normalized_items: list[LineItem] = []
    for text, image_path in items:
        if image_path:
            normalized_items.append((text, image_path))
            continue
        parts = _split_embedded_question_transition(text or "")
        if not parts:
            continue
        normalized_items.extend((part, None) for part in parts)
    return normalized_items


def _collect_option_spans(items: list[LineItem], a: int, b: int) -> list[tuple[int, int, list[RichLine]]]:
    """在 [a,b) 内找出所有选项块 (start, end, options)。"""
    spans: list[tuple[int, int, list[RichLine]]] = []
    i = a
    while i < b:
        oc = _option_cluster_end(items, i, b)
        if not oc:
            i += 1
            continue
        opt_end, opts = oc
        spans.append((i, opt_end, opts))
        i = opt_end
    return spans


def parse_material_body(
    items: list[LineItem],
    body_start: int,
    body_end: int,
    header: str,
) -> MaterialUnit | None:
    """解析 [body_start, body_end) 正文区间（无「材料X」标题行）。"""
    spans = _collect_option_spans(items, body_start, body_end)
    if not spans:
        return None

    questions: list[ExamQuestion] = []
    first_start, _first_end, first_opts = spans[0]
    first_seg = items[body_start:first_start]
    first_number = _extract_source_number_from_segment(first_seg)
    intro, stem = _split_intro_stem(first_seg)
    questions.append(
        ExamQuestion(stem_lines=stem, option_lines=first_opts, source_number=first_number)
    )

    for k in range(1, len(spans)):
        prev_end = spans[k - 1][1]
        cur_start, _cur_end, cur_opts = spans[k]
        seg = items[prev_end:cur_start]
        number = _extract_source_number_from_segment(seg)
        questions.append(
            ExamQuestion(
                stem_lines=_segment_to_stem_only(seg),
                option_lines=cur_opts,
                source_number=number,
            )
        )

    return MaterialUnit(header=header, intro_lines=intro, questions=questions)


def _split_into_material_units(unit: MaterialUnit) -> list[MaterialUnit]:
    """资料分析 20 题常为四组×5 题；无「材料」标记时按 5 题一组拆分。"""
    qs = unit.questions
    n = len(qs)
    group_size = 5
    if n < group_size or n % group_size != 0:
        return [unit]
    labels = "一二三四五六七八九十"
    out: list[MaterialUnit] = []
    for g in range(n // group_size):
        label = labels[g] if g < len(labels) else str(g + 1)
        chunk = qs[g * group_size : (g + 1) * group_size]
        chunk_questions = [
            ExamQuestion(
                stem_lines=list(q.stem_lines),
                option_lines=list(q.option_lines),
                source_number=q.source_number,
            )
            for q in chunk
        ]
        if g == 0:
            intro = list(unit.intro_lines)
        else:
            spill_intro: list[RichLine] = []
            if out and out[-1].questions:
                prev_last = out[-1].questions[-1]
                clean_options, spill_intro = _split_material_intro_from_option_lines(prev_last.option_lines)
                prev_last.option_lines = clean_options
            intro, stem = _split_rich_intro_stem(chunk_questions[0].stem_lines)
            if spill_intro:
                intro = spill_intro + intro
            chunk_questions[0].stem_lines = stem
        out.append(MaterialUnit(header=f"材料{label}", intro_lines=intro, questions=chunk_questions))
    return out


def parse_material_block(items: list[LineItem], header_idx: int, block_end: int) -> MaterialUnit | None:
    """解析 [header_idx, block_end) 材料块（首行为 材料X）。"""
    header = items[header_idx][0].strip()
    return parse_material_body(items, header_idx + 1, block_end, header)


def parse_quant_block(items: list[LineItem], a: int, b: int) -> list[ExamQuestion]:
    spans = _collect_option_spans(items, a, b)
    if not spans:
        return []

    questions: list[ExamQuestion] = []
    first_start, first_end, first_opts = spans[0]
    first_seg = items[a:first_start]
    first_number = _extract_source_number_from_segment(first_seg)
    questions.append(
        ExamQuestion(
            stem_lines=_segment_to_stem_only(first_seg),
            option_lines=first_opts,
            source_number=first_number,
        )
    )
    for k in range(1, len(spans)):
        prev_end = spans[k - 1][1]
        cur_start, cur_end, cur_opts = spans[k]
        seg = items[prev_end:cur_start]
        number = _extract_source_number_from_segment(seg)
        questions.append(
            ExamQuestion(
                stem_lines=_segment_to_stem_only(seg),
                option_lines=cur_opts,
                source_number=number,
            )
        )
    return questions


def _normalize_subject_selection(mode: str | Iterable[str]) -> set[SubjectKind]:
    if isinstance(mode, str):
        raw_parts = [
            part.strip()
            for chunk in mode.replace("，", ",").replace("、", ",").split(",")
            for part in [chunk]
            if part.strip()
        ]
    else:
        raw_parts = [str(part).strip() for part in mode if str(part).strip()]

    if not raw_parts:
        return set(ALL_SUBJECT_KINDS)

    selected: set[SubjectKind] = set()
    for raw in raw_parts:
        token = raw.lower()
        if token in ("all", "*"):
            return set(ALL_SUBJECT_KINDS)
        if token == "both":
            selected.update(("quant", "data"))
            continue
        if token in ALL_SUBJECT_KINDS:
            selected.add(token)  # type: ignore[arg-type]
            continue
        for kind, labels in _SECTION_LABELS.items():
            if raw in labels:
                selected.add(kind)
                break
    return selected or set(ALL_SUBJECT_KINDS)


def _append_objective_section(exam: ParsedExam, kind: SubjectKind, title: str, questions: list[ExamQuestion]) -> None:
    if not questions:
        return
    if kind == "politics":
        exam.politics_sections.append(PoliticsSection(title=title, questions=questions))
    elif kind == "common_sense":
        exam.common_sense_sections.append(CommonSenseSection(title=title, questions=questions))
    elif kind == "verbal":
        exam.verbal_sections.append(VerbalSection(title=title, questions=questions))
    elif kind == "quant":
        exam.quant_sections.append(QuantSection(title=title, questions=questions))
    elif kind == "reasoning":
        exam.reasoning_sections.append(ReasoningSection(title=title, questions=questions))


def parse_line_items(
    items: list[LineItem],
    mode: Literal["data", "quant", "both", "all"] | str = "all",
) -> ParsedExam:
    items = _preprocess_line_items(items)
    exam = ParsedExam()
    n = len(items)
    selected_subjects = _normalize_subject_selection(mode)
    # (篇类, 合并后的篇题, 篇题起始行下标, 正文起始行下标, 是否需要解析)
    title_entries: list[tuple[str, str, int, int, bool]] = []

    i = 0
    while i < n:
        t, img = items[i]
        if img:
            i += 1
            continue
        line = _nfkc((t or "").strip())
        if not line:
            i += 1
            continue

        paired_cn = _pair_cn_section(items, i)
        if paired_cn:
            kind, merged, end_i = paired_cn
            should_parse = kind in selected_subjects
            title_entries.append((kind, merged, i, end_i, should_parse))
            i = end_i
            continue

        paired = _pair_section_title(items, i)
        if paired:
            kind, merged, end_i = paired
            should_parse = kind in selected_subjects
            title_entries.append((kind, merged, i, end_i, should_parse))
            i = end_i
            continue

        detected_kind = _detect_subject_section_kind(line)
        if detected_kind:
            title_entries.append((detected_kind, line, i, i + 1, detected_kind in selected_subjects))
        elif _is_other_section_title(line):
            title_entries.append(("other", line, i, i + 1, False))

        i += 1

    for j, (kind, title, _start_idx, body_start, should_parse) in enumerate(title_entries):
        if not should_parse:
            continue
        next_start = title_entries[j + 1][2] if j + 1 < len(title_entries) else n
        body_end = next_start
        body_start = _skip_section_boilerplate(items, body_start, body_end, kind=kind)  # type: ignore[arg-type]

        if kind == "data":
            sec = DataAnalysisSection(title=title, materials=[])
            mat_positions = [
                i
                for i in range(body_start, body_end)
                if not items[i][1] and material_header_line((items[i][0] or "").strip())
            ]
            if not mat_positions:
                unit = parse_material_body(items, body_start, body_end, "材料一")
                if unit:
                    for u in _split_into_material_units(unit):
                        sec.materials.append(u)
            else:
                for mi, m_start in enumerate(mat_positions):
                    m_next = mat_positions[mi + 1] if mi + 1 < len(mat_positions) else body_end
                    unit = parse_material_block(items, m_start, m_next)
                    if unit:
                        sec.materials.append(unit)
            exam.data_sections.append(sec)
        else:
            _append_objective_section(
                exam,
                kind,  # type: ignore[arg-type]
                title,
                parse_quant_block(items, body_start, body_end),
            )

    return exam
