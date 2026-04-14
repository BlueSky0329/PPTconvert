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
    UnknownSection,
    VerbalSection,
)
from domain.models import ALL_SUBJECT_KINDS, SubjectKind
from core.subject_inference import (
    default_subject_title,
    infer_document_subject,
    infer_pdf_filename_profile,
    infer_subject_from_content,
    resolve_objective_section_kinds,
)
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
    normalized = _normalize_pdf_option_text(text)
    if bool(list(OPTION_MARKER.finditer(normalized))):
        return True
    if _match_single_option_line(normalized):
        return True
    return len(_find_spaced_bare_option_markers(normalized)) >= 2


# 竖排选项：每行 A. xxx
_VERTICAL_OPTION = re.compile(
    r"^\s*([ABCDabcd])\s*[\.．、:：]\s*(.+)$",
)
_SINGLE_OPTION = re.compile(r"^\s*([ABCDabcd])\s*[\.．、:：)\uFF09]\s*(.*)$")
_BARE_SINGLE_OPTION = re.compile(r"^\s*([ABCDabcd])\s+(.+)$")
_QUESTION_NO_LINE = re.compile(r"^\s*\d{1,3}\s*[\.．、]?\s*$")
_LEADING_QUESTION_LINE = re.compile(
    r"^\s*(?P<number>\d{1,3})\s*(?P<delimiter>[\.．、)\uFF09])(?P<gap>\s*)(?P<stem>.+)$"
)
_LEADING_PLAIN_QUESTION_LINE = re.compile(r"^\s*(?P<number>\d{2,3})\s+(?P<stem>.+)$")
_LEADING_CN_QUESTION_LINE = re.compile(
    r"^\s*第\s*(?P<number>\d{1,3})\s*题(?:\s*(?P<delimiter>[\uff1a:.\uFF0E\u3001)\uFF09])(?P<gap>\s*))?(?P<stem>.+)$"
)
_EMBEDDED_QUESTION_TRANSITION = re.compile(
    r"^(?P<head>.*?\S)\s+(?P<number>\d{1,3})\s*[\.．、)\uFF09]\s*(?P<stem>.+)$"
)
_NUMERIC_SEQUENCE_STEM = re.compile(
    r"^[\d\.\-＋\+×xX÷/\\,，:：;；%=％()\[\]（）<>≤≥≈~～mM米千百十万亿个只张年月日]+$"
)
_SHORT_OPTION_FRAGMENT = re.compile(
    r"^[\d\.\-＋\+×xX÷/\\,，:：;；%=％()\[\]（）<>≤≥≈~～mM米千百十万亿个只张]{1,8}$"
)
_ISOLATED_NUMBER_LINE = re.compile(r"^\s*(\d{2,3})\s*[\.．、]?\s*$")
_NUMERIC_CONTINUATION_PREFIX = re.compile(r"(?:用近|近|约|共|历时|耗时|长达|超过|超|不足|不到|将近|近乎)$")
_NUMERIC_CONTINUATION_NEXT = re.compile(
    r"^(?:天|年|月|日|周|小时|分钟|秒|人|名|个|家|项|亩|米|公里|千米|亿元|万元|元|倍|%|％)"
)
_LEADING_QUANTITY_MEASURE_STEM = re.compile(
    r"^(?:余|多)(?:颗|个|人|名|家|项|亩|米|公里|千米|亿元|万元|元|倍|天|年|月|日|周|小时|分钟|秒|台|架|艘|只|张|条|次)"
)

MATERIAL_HEADER = re.compile(r"^\s*材料[一二三四五六七八九十百千万\d〇零两]+\s*$")
GENERIC_MATERIAL_HEADER = re.compile(r"^\s*[【\[]?\s*材料\s*[】\]]?\s*[：:]?\s*$")
_SCAN_AD_LINE = re.compile(
    r"^\s*各种考试资料购买\s*[,，]\s*请加微信\s*[:：]\s*行测资料库\s*$"
)
_TOPIC_HEADER_LINE = re.compile(
    r"^\s*行测\s*[—\-–―]+\s*(数量关系|判断推理|言语理解|资料分析).*$"
)
_SECTION_PART_MARKER = re.compile(r"^\s*[（(]?\s*[一二三四五六七八九十\d]+\s*[)）]?\s*$")


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


def generic_material_header_line(line: str) -> bool:
    s = _nfkc((line or "").strip())
    if not s:
        return False
    return bool(GENERIC_MATERIAL_HEADER.match(s))


def _material_header_label(index: int) -> str:
    labels = "一二三四五六七八九十"
    return f"材料{labels[index - 1]}" if 1 <= index <= len(labels) else f"材料{index}"


def _collect_material_blocks(
    items: list[LineItem],
    body_start: int,
    body_end: int,
) -> list[tuple[int, int, str]]:
    numbered_positions = [
        i
        for i in range(body_start, body_end)
        if not items[i][1] and material_header_line((items[i][0] or "").strip())
    ]
    if numbered_positions:
        blocks: list[tuple[int, int, str]] = []
        for idx, start in enumerate(numbered_positions, start=1):
            end = numbered_positions[idx] if idx < len(numbered_positions) else body_end
            blocks.append((start, end, (items[start][0] or "").strip() or _material_header_label(idx)))
        return blocks

    generic_positions = [
        i
        for i in range(body_start, body_end)
        if not items[i][1] and generic_material_header_line((items[i][0] or "").strip())
    ]
    if not generic_positions:
        return []
    if len(generic_positions) == 1 and generic_positions[0] != body_start:
        return []

    blocks = []
    for idx, start in enumerate(generic_positions, start=1):
        end = generic_positions[idx] if idx < len(generic_positions) else body_end
        blocks.append((start, end, _material_header_label(idx)))
    return blocks

# 题干里常出现的「资料分析」短语，应排除，避免误判为篇题
_TITLE_FALSE_POSITIVE = re.compile(
    r"^(根据|由|从|结合|阅读|下列|关于|对于|若|如|由此|以下|能够|不能|推出|说法)"
)
_BARE_OPTION_MARKER = re.compile(r"(?<![A-Za-z])([A-D])\s+(?=\S)", re.IGNORECASE)
_QUESTION_PROMPT_HINT = re.compile(
    r"^(?:根据上述|根据下列|根据材料|根据题意|以下哪项|以下哪个|下列哪项|下列哪个|下列最|由此可以推出|从所给的|问哪项|问哪个|问下列|问)"
)
_QUESTION_STEM_HINT = re.compile(
    r"(?:下列|哪项|哪个|最能|最不|属于|不属于|符合|不符合|正确|错误|能够|不能|可以|不可能|规律|关系|的是|原因|结果|削弱|支持|推出)"
)

_SECTION_LABELS: dict[SubjectKind, tuple[str, ...]] = {
    "politics": ("政治理论",),
    "common_sense": ("常识判断",),
    "verbal": ("言语理解与表达", "言语理解和表达"),
    "quant": ("数量关系",),
    "reasoning": ("判断推理",),
    "data": ("资料分析",),
}

_LEADING_STEM_FILLERS = " \t　_＿·….,，、:：;；-—―\"'“”‘’()（）[]【】"


def _nfkc(s: str) -> str:
    """兼容区字符（如 ⼀）归一成常规汉字/标点。"""
    return unicodedata.normalize("NFKC", s or "")


def _normalize_digits(s: str) -> str:
    """全角数字转半角，便于匹配年份。"""
    return s.translate(str.maketrans("０１２３４５６７８９", "0123456789"))


_CN_NUM = "一二三四五六七八九十"
_OUTLINE_SECTION_TITLE = re.compile(
    rf"^\s*(?:第[一二三四五六七八九十百\d〇零]+部分|[{_CN_NUM}]{{1,3}}[、．。.])\s*[\u4e00-\u9fff]+"
)
_REASONING_SUBSECTION_TITLE = re.compile(
    r"^\s*[一二三四]\s*[、．。.]\s*(图形推理|定义判断|类比推理|逻辑判断)"
)


def _is_boilerplate_line(line: str) -> bool:
    """篇首说明（非题目），需跳过。"""
    s = (line or "").strip()
    if s in ("案。", "答案。"):
        return True
    if _REASONING_SUBSECTION_TITLE.match(s):
        return True
    if "参考时限" in s and "共" in s and "题" in s:
        return True
    if s.startswith("请开始答题"):
        return True
    hints = (
        "本部分包括表达与理解",
        "根据题目要求，在四个选项中选出一个正确答案",
        "根据题目要求,在四个选项中选出一个正确答案",
        "根据题目要求，在四个选项中选出一个最恰当的答案",
        "根据题目要求,在四个选项中选出一个最恰当的答案",
        "在四个选项中选出一个最",
        "恰当的答案",
        "在这部分试题中",
        "所给出的图",
        "表、文字或综合性资料",
        "图形推理。请按每道题的答题要求作答",
        "每道题先给出定义",
        "最符合或最不符合该定义的答案",
        "每道题先给出一组相关的词",
        "在逻辑关系上最为贴近",
        "这段陈述被假设是正确的",
        "不容置疑",
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


def _strip_stem_leading_fillers(stem: str) -> str:
    return _nfkc(_normalize_digits((stem or "").strip())).lstrip(_LEADING_STEM_FILLERS)


def _find_spaced_bare_option_markers(text: str) -> list[re.Match[str]]:
    matches: list[re.Match[str]] = []
    for match in _BARE_OPTION_MARKER.finditer(text):
        if match.start() == 0 or text[match.start() - 1].isspace():
            matches.append(match)
    return matches


def _match_single_option_line(text: str) -> re.Match[str] | None:
    punctuated = _SINGLE_OPTION.match(text)
    if punctuated:
        return punctuated

    bare = _BARE_SINGLE_OPTION.match(text)
    if not bare:
        return None

    trailing_markers = [
        match for match in _BARE_OPTION_MARKER.finditer(text) if match.start() > bare.start()
    ]
    if trailing_markers:
        return None
    return bare


def _looks_like_year_leading_stem(stem: str) -> bool:
    s = _nfkc(_normalize_digits((stem or "").strip()))
    return bool(re.match(r"^(?:19|20)\d{2}(?:\D|$)", s))


def _looks_like_numbered_intro_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    return bool(re.match(r"^[1-9][\u4e00-\u9fff]", s))


def _looks_like_numeric_sequence_stem(stem: str) -> bool:
    s = _nfkc(_normalize_digits((stem or "").strip()))
    if not s or len(s) > 48:
        return False
    if not _NUMERIC_SEQUENCE_STEM.match(s):
        return False
    digit_count = sum(ch.isdigit() for ch in s)
    punctuation_count = sum(ch in ",，.．-+＋×xX÷/\\()[]（）:：;；%=％<>≤≥≈~～" for ch in s)
    return digit_count >= 4 and punctuation_count >= 2


def _looks_like_quantity_measure_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    if not s:
        return False
    return bool(_LEADING_QUANTITY_MEASURE_STEM.match(s))


def _is_parser_noise_text(text: str) -> bool:
    s = _nfkc((text or "").strip())
    if not s:
        return True
    if _SCAN_AD_LINE.match(s):
        return True
    if _TOPIC_HEADER_LINE.match(s):
        return True
    return False


def _is_short_option_fragment(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s or len(s) > 8:
        return False
    return bool(_SHORT_OPTION_FRAGMENT.match(s))


def _looks_like_question_stem_text(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s:
        return False
    if _QUESTION_PROMPT_HINT.match(s):
        return True
    if s.endswith(("?", "？", ":", "：")):
        return True
    if len(s) >= 10 and _QUESTION_STEM_HINT.search(s):
        return True
    return False


def _merge_fragmented_option_line(
    items: list[LineItem],
    index: int,
) -> tuple[str, int] | None:
    if index + 2 >= len(items):
        return None

    text, image_path = items[index]
    fragment_text, fragment_image = items[index + 1]
    next_text, next_image = items[index + 2]
    if image_path or fragment_image or next_image:
        return None

    normalized = _normalize_pdf_option_text((text or "").strip())
    match = _match_single_option_line(normalized)
    if not match:
        return None

    current_letter = match.group(1).upper()
    current_body = (match.group(2) or "").strip()
    if current_letter not in "ABCD":
        return None

    fragment = _nfkc(_normalize_digits((fragment_text or "").strip()))
    if not _is_short_option_fragment(fragment):
        return None

    upcoming = _normalize_pdf_option_text((next_text or "").strip())
    next_match = _match_single_option_line(upcoming)
    expected_next = "ABCD"["ABCD".index(current_letter) + 1] if current_letter != "D" else None
    if not expected_next or not next_match or next_match.group(1).upper() != expected_next:
        return None

    merged_body = f"{current_body}{fragment}".strip()
    return f"{current_letter}. {merged_body}", 1


def _match_leading_question_with_stem(line: str) -> tuple[str, str] | None:
    s = _nfkc(_normalize_digits((line or "").strip()))
    for pattern in (_LEADING_QUESTION_LINE, _LEADING_CN_QUESTION_LINE, _LEADING_PLAIN_QUESTION_LINE):
        match = pattern.match(s)
        if match:
            number = (match.group("number") or "").strip()
            gap = match.groupdict().get("gap", " ")
            stem = (match.group("stem") or "").strip()
            head = _strip_stem_leading_fillers(stem)
            if not head:
                continue
            if head[0].isdigit() and not gap and not (
                _looks_like_year_leading_stem(head)
                or _looks_like_numbered_intro_stem(head)
                or _looks_like_numeric_sequence_stem(head)
            ):
                continue
            if pattern is _LEADING_PLAIN_QUESTION_LINE and _looks_like_quantity_measure_stem(head):
                continue
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
    if _REASONING_SUBSECTION_TITLE.match(s):
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
    match = _match_single_option_line(text)
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
    normalized = _normalize_pdf_option_text(text)
    raw = _parse_options_from_line(normalized)
    if len(raw) < 2:
        matches = _find_spaced_bare_option_markers(normalized)
        if len(matches) >= 2:
            raw = []
            for idx, match in enumerate(matches):
                letter = match.group(1).upper()
                body_start = match.end()
                body_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(normalized)
                raw.append(ExamOption(letter=letter, text=normalized[body_start:body_end].strip()))
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
    trailing_after_complete: list[RichLine] = []

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
        if _is_parser_noise_text(raw):
            i += 1
            continue

        normalized = _normalize_pdf_option_text(raw)
        match = _match_single_option_line(normalized)
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
                if expected_index >= len(letters):
                    trailing_after_complete.append(_rich_img(next_img))
                else:
                    out.append(_rich_img(next_img))
                i += 1
                continue

            next_raw = (next_text or "").strip()
            if not next_raw:
                i += 1
                continue
            if _is_parser_noise_text(next_raw):
                i += 1
                continue

            next_normalized = _normalize_pdf_option_text(next_raw)
            next_match = _match_single_option_line(next_normalized)
            if expected_index < len(letters) and next_match and next_match.group(1).upper() == letters[expected_index]:
                break
            if expected_index >= len(letters) and next_match and next_match.group(1).upper() in letters:
                break
            if _starts_new_question_line(next_normalized) or material_header_line(next_normalized):
                break
            if _detect_subject_section_kind(next_normalized) or _is_other_section_title(next_normalized):
                break
            if expected_index >= len(letters) and (
                _SECTION_PART_MARKER.match(next_normalized)
                or _is_parser_noise_text(next_normalized)
                or
                _REASONING_SUBSECTION_TITLE.match(next_normalized)
                or _QUESTION_PROMPT_HINT.match(next_normalized)
            ):
                break
            if expected_index >= len(letters) and trailing_after_complete:
                out.extend(trailing_after_complete)
                trailing_after_complete = []
            out.append(_rich_text(next_raw))
            i += 1

    if expected_index == len(letters) and trailing_after_complete:
        out.extend(trailing_after_complete)
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
        if _is_parser_noise_text(raw):
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
        rich_lines = _options_to_rich_lines(opts[:4])
        if _option_cluster_has_substance(rich_lines) or _cluster_can_rebalance_from_preceding_images(
            items, start, rich_lines
        ):
            return start + 1, rich_lines

    if opts and len(opts) == 2 and start + 1 < end:
        second = text_at(start + 1)
        if not second or not _has_option_markers_pdf(second):
            return None
        merged = _normalize_pdf_option_text(first) + "\t" + _normalize_pdf_option_text(second)
        opts2 = _parse_options_line(merged)
        if opts2 and len(opts2) >= 4:
            rich_lines = _options_to_rich_lines(opts2[:4])
            if _option_cluster_has_substance(rich_lines) or _cluster_can_rebalance_from_preceding_images(
                items, start, rich_lines
            ):
                return start + 2, rich_lines

    seq = _collect_sequential_option_lines(items, start, end)
    if seq:
        seq_end, seq_lines = seq
        if _option_cluster_has_substance(seq_lines) or _cluster_can_rebalance_from_preceding_images(
            items, start, seq_lines
        ):
            return seq_end, seq_lines

    if opts and 1 < len(opts) < 4:
        accumulated = _collect_accumulated_option_lines(items, start, end)
        if accumulated:
            acc_end, acc_lines = accumulated
            if _option_cluster_has_substance(acc_lines) or _cluster_can_rebalance_from_preceding_images(
                items, start, acc_lines
            ):
                return acc_end, acc_lines

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


def _is_blank_option_placeholder(rich: RichLine) -> bool:
    text = _normalize_pdf_option_text(_rich_line_text_value(rich))
    if not text or _rich_line_has_image(rich):
        return False
    match = _match_single_option_line(text)
    if not match:
        return False
    return not (match.group(2) or "").strip()


def _option_cluster_has_substance(option_lines: list[RichLine]) -> bool:
    letters_seen: set[str] = set()
    nonempty_text_count = 0
    image_count = 0

    for line in option_lines:
        if _rich_line_has_image(line):
            image_count += 1
        text = _normalize_pdf_option_text(_rich_line_text_value(line))
        if not text:
            continue
        parsed = _parse_options_line(text)
        if parsed:
            for option in parsed:
                letters_seen.add(option.letter.upper())
                if option.text.strip():
                    nonempty_text_count += 1
            continue
        matched = _match_single_option_line(text)
        if matched:
            letters_seen.add(matched.group(1).upper())
            if (matched.group(2) or "").strip():
                nonempty_text_count += 1

    if not letters_seen:
        return False
    if nonempty_text_count > 0:
        return True
    return image_count >= len(letters_seen)


def _blank_option_placeholder_letters(option_lines: list[RichLine]) -> list[str]:
    letters: list[str] = []
    for line in option_lines:
        if _rich_line_has_image(line):
            continue
        if not _is_blank_option_placeholder(line):
            return []
        letter = _rich_option_letter(line)
        if not letter:
            return []
        letters.append(letter)
    return letters


def _cluster_can_rebalance_from_preceding_images(
    items: list[LineItem],
    start: int,
    option_lines: list[RichLine],
) -> bool:
    if _blank_option_placeholder_letters(option_lines) != ["A", "B", "C", "D"]:
        return False

    image_count = 0
    index = start - 1
    while index >= 0:
        text, image_path = items[index]
        if image_path:
            image_count += 1
            index -= 1
            continue
        if (text or "").strip():
            break
        index -= 1
    return image_count >= 4


def _rebalance_trailing_option_images(
    segment: list[LineItem],
    option_lines: list[RichLine],
) -> tuple[list[LineItem], list[RichLine]]:
    if not segment or not option_lines:
        return segment, option_lines
    if not all(_is_blank_option_placeholder(line) for line in option_lines):
        return segment, option_lines

    letters = [_rich_option_letter(line) or "" for line in option_lines]
    if letters != [chr(ord("A") + index) for index in range(len(option_lines))]:
        return segment, option_lines

    trailing_images: list[str] = []
    split_index = len(segment)
    for index in range(len(segment) - 1, -1, -1):
        text, image_path = segment[index]
        if image_path:
            trailing_images.append(image_path)
            split_index = index
            continue
        if (text or "").strip():
            break
    trailing_images.reverse()
    if len(trailing_images) != len(option_lines):
        return segment, option_lines

    rebalanced_option_lines: list[RichLine] = []
    for image_path, option_line in zip(trailing_images, option_lines):
        rebalanced_option_lines.append(_rich_img(image_path))
        rebalanced_option_lines.append(option_line)
    return segment[:split_index], rebalanced_option_lines


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
            normalized_stem = _strip_stem_leading_fillers(stem)
            if normalized_stem and normalized_stem[0].isdigit() and not (
                _looks_like_year_leading_stem(normalized_stem)
                or _looks_like_numbered_intro_stem(normalized_stem)
                or _looks_like_numeric_sequence_stem(normalized_stem)
            ):
                updated.append(piece)
                continue
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
    i = 0
    while i < len(items):
        text, image_path = items[i]
        if image_path:
            normalized_items.append((text, image_path))
            i += 1
            continue
        normalized_text = _nfkc((text or "").strip())
        if _is_parser_noise_text(normalized_text):
            i += 1
            continue
        skip_count = 0
        merged_option = _merge_fragmented_option_line(items, i)
        if merged_option is not None:
            text, skip_count = merged_option
        if i + 4 < len(items):
            char_window = [
                _nfkc((items[i + offset][0] or "").strip())
                for offset in range(4)
            ]
            tail_text, tail_image = items[i + 4]
            if (
                all(items[i + offset][1] is None for offset in range(4))
                and char_window == ["扫", "码", "关", "注"]
                and tail_image is None
                and _SCAN_AD_LINE.match(_nfkc((tail_text or "").strip()))
            ):
                i += 5
                continue
        parts = _split_embedded_question_transition(text or "")
        if not parts:
            i += 1 + skip_count
            continue
        normalized_items.extend((part, None) for part in parts)
        i += 1 + skip_count
    return _merge_numeric_continuation_fragments(normalized_items)


def _merge_numeric_continuation_fragments(items: list[LineItem]) -> list[LineItem]:
    merged: list[LineItem] = []
    i = 0
    while i < len(items):
        text, image_path = items[i]
        if (
            image_path is None
            and i + 1 < len(items)
            and merged
            and merged[-1][1] is None
            and items[i + 1][1] is None
        ):
            prev_text = (merged[-1][0] or "").strip()
            current_text = (text or "").strip()
            next_text = (items[i + 1][0] or "").strip()
            number_match = _ISOLATED_NUMBER_LINE.match(current_text)
            if (
                number_match
                and prev_text
                and next_text
                and _NUMERIC_CONTINUATION_PREFIX.search(_nfkc(prev_text))
                and _NUMERIC_CONTINUATION_NEXT.match(_nfkc(next_text))
            ):
                merged[-1] = (f"{prev_text}{number_match.group(1)}{next_text}", None)
                i += 2
                continue
        merged.append((text, image_path))
        i += 1
    return merged


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


def _collect_question_markers(items: list[LineItem], a: int, b: int) -> list[int]:
    markers: list[int] = []
    for i in range(a, b):
        text, image_path = items[i]
        if image_path:
            continue
        raw = (text or "").strip()
        if not raw or _is_parser_noise_text(raw):
            continue
        normalized = _normalize_pdf_option_text(raw)
        if _starts_new_question_line(normalized):
            marker_start = i
            if i - 1 >= a:
                prev_text, prev_image = items[i - 1]
                prev_raw = (prev_text or "").strip()
                prev_normalized = _normalize_pdf_option_text(prev_raw)
                if (
                    not prev_image
                    and prev_raw
                    and not _is_parser_noise_text(prev_raw)
                    and _looks_like_question_stem_text(prev_raw)
                    and not _starts_new_question_line(prev_normalized)
                    and not _has_option_markers_pdf(prev_normalized)
                    and not material_header_line(prev_normalized)
                    and not _detect_subject_section_kind(prev_normalized)
                    and not _is_other_section_title(prev_normalized)
                ):
                    marker_start = i - 1
            if not markers or markers[-1] != marker_start:
                markers.append(marker_start)
    return markers


def _append_option_tail_continuation(
    option_lines: list[RichLine],
    tail_items: list[LineItem],
) -> list[RichLine]:
    if not option_lines or not tail_items:
        return option_lines

    tail_texts: list[str] = []
    tail_images: list[str] = []
    for text, image_path in tail_items:
        if image_path:
            tail_images.append(image_path)
            continue
        raw = (text or "").strip()
        if not raw or _is_parser_noise_text(raw):
            continue
        normalized = _normalize_pdf_option_text(raw)
        if (
            _starts_new_question_line(normalized)
            or _has_option_markers_pdf(normalized)
            or material_header_line(normalized)
            or _detect_subject_section_kind(normalized)
            or _is_other_section_title(normalized)
            or _looks_like_question_stem_text(normalized)
        ):
            return option_lines
        tail_texts.append(raw)

    if not tail_texts and not tail_images:
        return option_lines

    updated = list(option_lines)
    if tail_texts:
        last = updated[-1]
        last_text = _rich_line_text_value(last)
        merged_text = "\n".join(part for part in [last_text, *tail_texts] if part)
        new_parts: list[tuple[str, str | None]] = []
        if merged_text:
            new_parts.append((merged_text, None))
        new_parts.extend([("", image_path) for image_path in tail_images])
        updated[-1] = RichLine(parts=new_parts)
    elif tail_images:
        for image_path in tail_images:
            updated.append(_rich_img(image_path))
    return updated


def _build_questions_from_option_spans(
    items: list[LineItem],
    start: int,
    spans: list[tuple[int, int, list[RichLine]]],
) -> list[ExamQuestion]:
    questions: list[ExamQuestion] = []
    first_start, _first_end, first_opts = spans[0]
    first_seg = items[start:first_start]
    first_seg, first_opts = _rebalance_trailing_option_images(first_seg, first_opts)
    questions.append(
        ExamQuestion(
            stem_lines=_segment_to_stem_only(first_seg),
            option_lines=first_opts,
            source_number=_extract_source_number_from_segment(first_seg),
        )
    )
    for k in range(1, len(spans)):
        prev_end = spans[k - 1][1]
        cur_start, _cur_end, cur_opts = spans[k]
        seg = items[prev_end:cur_start]
        seg, cur_opts = _rebalance_trailing_option_images(seg, cur_opts)
        questions.append(
            ExamQuestion(
                stem_lines=_segment_to_stem_only(seg),
                option_lines=cur_opts,
                source_number=_extract_source_number_from_segment(seg),
            )
        )
    return questions


def _build_question_from_segment(
    items: list[LineItem],
    start: int,
    end: int,
) -> ExamQuestion | None:
    if start >= end:
        return None

    segment = items[start:end]
    if not any(((text or "").strip() or image_path) for text, image_path in segment):
        return None

    spans = _collect_option_spans(items, start, end)
    if spans:
        opt_start, opt_end, opts = spans[0]
        stem_seg = items[start:opt_start]
        stem_seg, opts = _rebalance_trailing_option_images(stem_seg, opts)
        opts = _append_option_tail_continuation(opts, items[opt_end:end])
        source_number = _extract_source_number_from_segment(segment)
        stem_lines = _segment_to_stem_only(stem_seg)
        if stem_lines or opts or source_number:
            return ExamQuestion(
                stem_lines=stem_lines,
                option_lines=opts,
                source_number=source_number,
            )

    source_number = _extract_source_number_from_segment(segment)
    stem_lines = _segment_to_stem_only(segment)
    if stem_lines or source_number:
        return ExamQuestion(
            stem_lines=stem_lines,
            option_lines=[],
            source_number=source_number,
        )
    return None


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
    first_seg, first_opts = _rebalance_trailing_option_images(first_seg, first_opts)
    first_number = _extract_source_number_from_segment(first_seg)
    intro, stem = _split_intro_stem(first_seg)
    questions.append(
        ExamQuestion(stem_lines=stem, option_lines=first_opts, source_number=first_number)
    )

    for k in range(1, len(spans)):
        prev_end = spans[k - 1][1]
        cur_start, _cur_end, cur_opts = spans[k]
        seg = items[prev_end:cur_start]
        seg, cur_opts = _rebalance_trailing_option_images(seg, cur_opts)
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


def parse_material_block(
    items: list[LineItem],
    header_idx: int,
    block_end: int,
    header: str | None = None,
) -> MaterialUnit | None:
    """解析 [header_idx, block_end) 材料块（首行为 材料X）。"""
    header = header or items[header_idx][0].strip()
    return parse_material_body(items, header_idx + 1, block_end, header)


def parse_quant_block(items: list[LineItem], a: int, b: int) -> list[ExamQuestion]:
    questions: list[ExamQuestion] = []
    markers = _collect_question_markers(items, a, b)
    if markers:
        cursor = a
        for segment_end in markers[1:] + [b]:
            segment_spans = _collect_option_spans(items, cursor, segment_end)
            if len(segment_spans) > 1:
                questions.extend(_build_questions_from_option_spans(items, cursor, segment_spans))
            else:
                question = _build_question_from_segment(items, cursor, segment_end)
                if question is not None:
                    questions.append(question)
            cursor = segment_end
        if questions:
            return questions

    spans = _collect_option_spans(items, a, b)
    if not spans:
        return []
    return _build_questions_from_option_spans(items, a, spans)


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
    else:
        exam.unknown_sections.append(UnknownSection(title=title, questions=questions))


def _rich_lines_text(lines: list[RichLine]) -> str:
    return "\n".join(_rich_line_text_value(line) for line in lines if _rich_line_text_value(line))


def _question_option_texts(question: ExamQuestion) -> list[str]:
    option_texts: list[str] = []
    for line in question.option_lines:
        text = _rich_line_text_value(line)
        if text:
            option_texts.append(text)
    return option_texts


def _infer_question_subject(question: ExamQuestion, *, allow_data: bool = False) -> tuple[SubjectKind, float]:
    kind, confidence = infer_subject_from_content(
        stem=_rich_lines_text(question.stem_lines),
        options=_question_option_texts(question),
        image_count=sum(1 for line in question.stem_lines + question.option_lines if _rich_line_has_image(line)),
        allow_data=allow_data,
    )
    return kind, confidence


def _infer_dominant_question_subject(
    questions: list[ExamQuestion],
    *,
    limit: int = 48,
) -> tuple[SubjectKind | None, float]:
    scores: dict[SubjectKind, float] = {
        "politics": 0.0,
        "common_sense": 0.0,
        "verbal": 0.0,
        "quant": 0.0,
        "reasoning": 0.0,
        "data": 0.0,
    }
    sample_size = 0
    for question in questions[:limit]:
        kind, confidence = _infer_question_subject(question)
        if kind == "unknown":
            continue
        scores[kind] += max(confidence, 0.2)
        sample_size += 1

    if sample_size == 0:
        return None, 0.0

    ranked = sorted(scores.items(), key=lambda item: item[1], reverse=True)
    best_kind, best_score = ranked[0]
    second_score = ranked[1][1] if len(ranked) > 1 else 0.0
    if best_score <= 0.0:
        return None, 0.0

    confidence = min(1.6, best_score / max(sample_size, 1))
    if best_score < second_score + 1.2:
        confidence = min(confidence, 0.45)
    return best_kind, confidence


def _group_questions_by_subject(
    questions: list[ExamQuestion],
    *,
    default_kind: SubjectKind = "unknown",
    strict_default: bool = False,
) -> list[tuple[SubjectKind, list[ExamQuestion]]]:
    inferred_pairs, strong_text_signals = _infer_question_pairs(questions)
    resolved_kinds = resolve_objective_section_kinds(
        default_kind=default_kind,
        inferred_pairs=inferred_pairs,
        source_numbers=[question.source_number for question in questions],
        strong_text_signals=strong_text_signals,
        strict_default=strict_default,
    )
    groups: list[tuple[SubjectKind, list[ExamQuestion]]] = []
    for kind, question in zip(resolved_kinds, questions):
        if groups and groups[-1][0] == kind:
            groups[-1][1].append(question)
        else:
            groups.append((kind, [question]))
    return groups


def _infer_question_pairs(
    questions: list[ExamQuestion],
) -> tuple[list[tuple[SubjectKind, float]], list[bool]]:
    inferred_pairs: list[tuple[SubjectKind, float]] = []
    strong_text_signals: list[bool] = []
    for question in questions:
        inferred_kind, confidence = _infer_question_subject(question)
        stem_text = _rich_lines_text(question.stem_lines)
        option_text_blob = " ".join(_question_option_texts(question))
        inferred_pairs.append((inferred_kind, confidence))
        strong_text_signals.append(len(stem_text) >= 10 or len(option_text_blob) >= 18)
    return inferred_pairs, strong_text_signals


def _append_grouped_objective_sections(
    exam: ParsedExam,
    *,
    default_kind: SubjectKind,
    default_title: str,
    questions: list[ExamQuestion],
    selected_subjects: set[SubjectKind],
    strict_default: bool = False,
) -> None:
    grouped = _group_questions_by_subject(
        questions,
        default_kind=default_kind,
        strict_default=strict_default,
    )
    for index, (kind, grouped_questions) in enumerate(grouped):
        if kind != "unknown" and kind not in selected_subjects:
            continue
        title = default_title if kind == default_kind else default_subject_title(kind)
        _append_objective_section(exam, kind, title, grouped_questions)


def _parse_numeric_source_number(source_number: str) -> int | None:
    value = (source_number or "").strip()
    return int(value) if value.isdigit() else None


_SET_PAPER_CANONICAL_ORDER: tuple[SubjectKind, ...] = (
    "politics",
    "common_sense",
    "verbal",
    "quant",
    "reasoning",
)


def _set_paper_order_candidates() -> list[tuple[SubjectKind, ...]]:
    order = list(_SET_PAPER_CANONICAL_ORDER)
    candidates: list[tuple[SubjectKind, ...]] = []
    for start in range(len(order) - 1):
        candidates.append(tuple(order[start:]))
    return candidates


def _set_paper_emission_score(
    *,
    state_kind: SubjectKind,
    inferred_kind: SubjectKind,
    confidence: float,
    strong_text_signal: bool,
) -> float:
    if inferred_kind == state_kind:
        return 2.1 * max(confidence, 0.2) + 0.35
    if inferred_kind == "unknown":
        return 0.15 if strong_text_signal else 0.25
    return -1.45 * max(confidence, 0.25)


def _resolve_set_paper_objective_kinds(
    questions: list[ExamQuestion],
) -> list[SubjectKind] | None:
    if len(questions) < 5:
        return None

    inferred_pairs, strong_text_signals = _infer_question_pairs(questions)
    distinct_inferred = {
        kind
        for kind, confidence in inferred_pairs
        if kind in _SET_PAPER_CANONICAL_ORDER and confidence >= 0.55
    }
    if len(distinct_inferred) < 2:
        return None

    best_score = float("-inf")
    best_path: list[SubjectKind] | None = None

    for order in _set_paper_order_candidates():
        if len(order) < 2:
            continue
        state_count = len(order)
        question_count = len(questions)
        scores = [[float("-inf")] * state_count for _ in range(question_count)]
        parents = [[-1] * state_count for _ in range(question_count)]

        for state_index, state_kind in enumerate(order):
            if state_index > 0:
                continue
            inferred_kind, confidence = inferred_pairs[0]
            scores[0][state_index] = _set_paper_emission_score(
                state_kind=state_kind,
                inferred_kind=inferred_kind,
                confidence=confidence,
                strong_text_signal=strong_text_signals[0],
            )

        for question_index in range(1, question_count):
            inferred_kind, confidence = inferred_pairs[question_index]
            for state_index, state_kind in enumerate(order):
                emission = _set_paper_emission_score(
                    state_kind=state_kind,
                    inferred_kind=inferred_kind,
                    confidence=confidence,
                    strong_text_signal=strong_text_signals[question_index],
                )
                best_prev_score = float("-inf")
                best_prev_index = -1
                for prev_index in range(state_index + 1):
                    transition = 0.12 if prev_index == state_index else -0.08 * (state_index - prev_index)
                    candidate_score = scores[question_index - 1][prev_index] + transition
                    if candidate_score > best_prev_score:
                        best_prev_score = candidate_score
                        best_prev_index = prev_index
                scores[question_index][state_index] = best_prev_score + emission
                parents[question_index][state_index] = best_prev_index

        final_state = max(range(state_count), key=lambda idx: scores[-1][idx])
        final_score = scores[-1][final_state]
        if final_score <= best_score:
            continue

        path_indices = [final_state]
        cursor = final_state
        for question_index in range(question_count - 1, 0, -1):
            cursor = parents[question_index][cursor]
            path_indices.append(cursor)
        path_indices.reverse()
        best_score = final_score
        best_path = [order[index] for index in path_indices]

    if not best_path:
        return None

    return best_path


def _looks_like_standard_set_paper_sequence(questions: list[ExamQuestion]) -> bool:
    numeric_numbers = [
        _parse_numeric_source_number(question.source_number)
        for question in questions
    ]
    numeric_numbers = [number for number in numeric_numbers if number is not None]
    if len(numeric_numbers) < 5:
        return False
    min_number = min(numeric_numbers)
    max_number = max(numeric_numbers)
    if min_number > 5 or max_number < 60:
        return False
    return _resolve_set_paper_objective_kinds(questions) is not None


def _append_adaptive_set_paper_objective_sections(
    exam: ParsedExam,
    *,
    questions: list[ExamQuestion],
    selected_subjects: set[SubjectKind],
) -> bool:
    resolved_kinds = _resolve_set_paper_objective_kinds(questions)
    if not resolved_kinds:
        return False

    groups: list[tuple[SubjectKind, list[ExamQuestion]]] = []
    for kind, question in zip(resolved_kinds, questions):
        if groups and groups[-1][0] == kind:
            groups[-1][1].append(question)
        else:
            groups.append((kind, [question]))

    for kind, grouped_questions in groups:
        if kind != "unknown" and kind not in selected_subjects:
            continue
        _append_objective_section(exam, kind, default_subject_title(kind), grouped_questions)
    return True


def _first_material_header_index(items: list[LineItem], start: int) -> int | None:
    for index in range(start, len(items)):
        text, image_path = items[index]
        if image_path:
            continue
        line = (text or "").strip()
        if material_header_line(line) or generic_material_header_line(line):
            return index
    return None


def _parse_set_paper_without_titles(
    items: list[LineItem],
    exam: ParsedExam,
    *,
    selected_subjects: set[SubjectKind],
) -> bool:
    start = _skip_section_boilerplate(items, 0, len(items), kind="quant")
    first_material = _first_material_header_index(items, start)
    objective_end = first_material if first_material is not None else len(items)
    objective_questions = parse_quant_block(items, start, objective_end)
    if not _looks_like_standard_set_paper_sequence(objective_questions):
        return False

    if not _append_adaptive_set_paper_objective_sections(
        exam,
        questions=objective_questions,
        selected_subjects=selected_subjects,
    ):
        return False

    if first_material is not None and "data" in selected_subjects:
        _append_data_section_from_body(
            exam,
            items,
            title=default_subject_title("data"),
            body_start=first_material,
            body_end=len(items),
        )
    return True


def _parse_whole_document_as_subject(
    items: list[LineItem],
    exam: ParsedExam,
    kind: SubjectKind,
    *,
    title: str | None = None,
) -> None:
    body_start = _skip_section_boilerplate(items, 0, len(items), kind="data" if kind == "data" else kind)
    section_title = title or default_subject_title(kind)
    if kind == "data":
        sec = DataAnalysisSection(title=section_title, materials=[])
        material_blocks = _collect_material_blocks(items, body_start, len(items))
        if not material_blocks:
            unit = parse_material_body(items, body_start, len(items), "材料一")
            if unit:
                for split_unit in _split_into_material_units(unit):
                    sec.materials.append(split_unit)
        else:
            for m_start, m_next, header in material_blocks:
                unit = parse_material_block(items, m_start, m_next, header=header)
                if unit:
                    sec.materials.append(unit)
        if sec.materials:
            exam.data_sections.append(sec)
        return

    questions = parse_quant_block(items, body_start, len(items))
    _append_objective_section(exam, kind, section_title, questions)


def _parse_without_titles(
    items: list[LineItem],
    exam: ParsedExam,
    selected_subjects: set[SubjectKind],
    *,
    source_name: str | None = None,
) -> None:
    filename_profile = infer_pdf_filename_profile(source_name or "")
    texts = [text for text, image_path in items if not image_path and (text or "").strip()]
    image_count = sum(1 for _text, image_path in items if image_path)
    material_header_count = sum(
        1
        for text, image_path in items
        if not image_path and (
            material_header_line((text or "").strip())
            or generic_material_header_line((text or "").strip())
        )
    )
    inferred_kind, confidence = infer_document_subject(
        texts,
        image_count=image_count,
        material_header_count=material_header_count,
    )
    start = _skip_section_boilerplate(items, 0, len(items), kind="quant")
    objective_questions = parse_quant_block(items, start, len(items))
    dominant_question_kind, dominant_question_confidence = _infer_dominant_question_subject(objective_questions)
    filename_subject_hint = (
        filename_profile.subject_hint
        if filename_profile.form == "single_subject_book"
        else None
    )
    if filename_subject_hint == "data":
        if inferred_kind in {None, "unknown", "data"} or confidence < 0.72:
            inferred_kind = "data"
            confidence = max(confidence, filename_profile.confidence)
    elif filename_subject_hint in {"politics", "common_sense", "verbal", "quant", "reasoning"}:
        if (
            inferred_kind is None
            or inferred_kind == "unknown"
            or inferred_kind == "data"
            or confidence < 0.72
        ):
            inferred_kind = filename_subject_hint
            confidence = max(confidence, filename_profile.confidence)
    if (
        dominant_question_kind in {"politics", "common_sense", "verbal", "quant", "reasoning"}
        and not filename_subject_hint
        and (
            inferred_kind is None
            or inferred_kind == "unknown"
            or (inferred_kind == "data" and material_header_count <= 2)
            or confidence < 0.55
        )
    ):
        inferred_kind = dominant_question_kind
        confidence = max(confidence, min(dominant_question_confidence, 0.95))
    elif filename_subject_hint and dominant_question_kind == filename_subject_hint:
        inferred_kind = filename_subject_hint
        confidence = max(
            confidence,
            min(max(dominant_question_confidence, filename_profile.confidence), 0.98),
        )

    if filename_profile.form == "single_subject_book" and filename_subject_hint in selected_subjects:
        _parse_whole_document_as_subject(
            items,
            exam,
            filename_subject_hint,
            title=default_subject_title(filename_subject_hint),
        )
        return

    if filename_profile.form == "set_paper":
        if _parse_set_paper_without_titles(
            items,
            exam,
            selected_subjects=selected_subjects,
        ):
            return
        strong_data_profile = (
            inferred_kind == "data"
            and confidence >= 0.78
            and material_header_count >= 3
        )
        if strong_data_profile and "data" in selected_subjects:
            _parse_whole_document_as_subject(
                items,
                exam,
                "data",
                title=default_subject_title("data"),
            )
            return
        _append_grouped_objective_sections(
            exam,
            default_kind="unknown",
            default_title=default_subject_title("unknown"),
            questions=objective_questions,
            selected_subjects=selected_subjects,
        )
        return

    sparse_material_headers_in_objective_book = (
        material_header_count > 0
        and material_header_count <= 2
        and len(objective_questions) >= 8
        and inferred_kind != "data"
    )
    if sparse_material_headers_in_objective_book:
        default_kind: SubjectKind = (
            inferred_kind
            if inferred_kind in {"politics", "common_sense", "verbal", "quant", "reasoning"}
            else "unknown"
        )
        if default_kind != "unknown" and confidence >= 0.75 and default_kind in selected_subjects:
            _parse_whole_document_as_subject(
                items,
                exam,
                default_kind,
                title=default_subject_title(default_kind),
            )
            return
        _append_grouped_objective_sections(
            exam,
            default_kind=default_kind,
            default_title=default_subject_title(default_kind),
            questions=objective_questions,
            selected_subjects=selected_subjects,
            strict_default=default_kind != "unknown",
        )
        return
    if material_header_count or inferred_kind == "data":
        if "data" not in selected_subjects:
            _append_grouped_objective_sections(
                exam,
                default_kind="unknown",
                default_title=default_subject_title("unknown"),
                questions=objective_questions,
                selected_subjects=selected_subjects,
            )
            return
        _parse_whole_document_as_subject(
            items,
            exam,
            "data",
            title=default_subject_title("data") if confidence >= 0.4 else default_subject_title("unknown"),
        )
        return
    if (
        inferred_kind in {"politics", "common_sense", "verbal", "quant", "reasoning"}
        and confidence >= 0.55
        and inferred_kind in selected_subjects
    ):
        _parse_whole_document_as_subject(items, exam, inferred_kind, title=default_subject_title(inferred_kind))
        return

    if not objective_questions:
        return
    _append_grouped_objective_sections(
        exam,
        default_kind="unknown",
        default_title=default_subject_title("unknown"),
        questions=objective_questions,
        selected_subjects=selected_subjects,
    )


def _append_data_section_from_body(
    exam: ParsedExam,
    items: list[LineItem],
    *,
    title: str,
    body_start: int,
    body_end: int,
) -> None:
    sec = DataAnalysisSection(title=title, materials=[])
    material_blocks = _collect_material_blocks(items, body_start, body_end)
    if not material_blocks:
        unit = parse_material_body(items, body_start, body_end, "材料一")
        if unit:
            for split_unit in _split_into_material_units(unit):
                sec.materials.append(split_unit)
    else:
        for m_start, m_next, header in material_blocks:
            unit = parse_material_block(items, m_start, m_next, header=header)
            if unit:
                sec.materials.append(unit)
    if sec.materials:
        exam.data_sections.append(sec)


def parse_line_items(
    items: list[LineItem],
    mode: Literal["data", "quant", "both", "all"] | str = "all",
    document_subject_hint: SubjectKind | None = None,
    source_name: str | None = None,
) -> ParsedExam:
    items = _preprocess_line_items(items)
    exam = ParsedExam()
    n = len(items)
    selected_subjects = _normalize_subject_selection(mode)

    if document_subject_hint and document_subject_hint != "unknown":
        _parse_whole_document_as_subject(items, exam, document_subject_hint)
        return exam
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

    if not title_entries:
        _parse_without_titles(items, exam, selected_subjects, source_name=source_name)
        return exam

    for j, (kind, title, _start_idx, body_start, should_parse) in enumerate(title_entries):
        if not should_parse:
            continue
        next_start = title_entries[j + 1][2] if j + 1 < len(title_entries) else n
        body_end = next_start
        body_start = _skip_section_boilerplate(items, body_start, body_end, kind=kind)  # type: ignore[arg-type]

        if kind == "data":
            _append_data_section_from_body(
                exam,
                items,
                title=title,
                body_start=body_start,
                body_end=body_end,
            )
        else:
            material_positions = [
                i
                for i in range(body_start, body_end)
                if not items[i][1] and material_header_line((items[i][0] or "").strip())
            ]
            if material_positions:
                first_material = material_positions[0]
                objective_questions = parse_quant_block(items, body_start, first_material)
                _append_grouped_objective_sections(
                    exam,
                    default_kind=kind,  # type: ignore[arg-type]
                    default_title=title,
                    questions=objective_questions,
                    selected_subjects=selected_subjects,
                    strict_default=True,
                )
                if "data" in selected_subjects:
                    _append_data_section_from_body(
                        exam,
                        items,
                        title=default_subject_title("data"),
                        body_start=first_material,
                        body_end=body_end,
                    )
            else:
                _append_grouped_objective_sections(
                    exam,
                    default_kind=kind,  # type: ignore[arg-type]
                    default_title=title,
                    questions=parse_quant_block(items, body_start, body_end),
                    selected_subjects=selected_subjects,
                    strict_default=True,
                )

    return exam
