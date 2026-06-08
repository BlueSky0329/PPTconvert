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
_GLUED_SINGLE_OPTION = re.compile(r"^\s*([ABCDabcd])(?=\d)(.+)$")
_GLUED_QUOTED_OPTION = re.compile(r"^\s*([ABCDabcd])(?=[《\"“'‘])(.*)$")
_QUESTION_NO_LINE = re.compile(r"^\s*[1-9]\d{0,3}\s*[\.．、]?\s*$")
_QUESTION_NO_OPEN_BRACKET_LINE = re.compile(r"^\s*[1-9]\d{0,3}\s*[\.．、]\s*[（(]\s*$")
_LEADING_QUESTION_LINE = re.compile(
    r"^\s*(?P<number>[1-9]\d{0,3})\s*(?P<delimiter>[\.．、)\uFF09])(?P<gap>\s*)(?P<stem>.+)$"
)
_LEADING_PLAIN_QUESTION_LINE = re.compile(r"^\s*(?P<number>\d{2,3})\s+(?P<stem>.+)$")
_LEADING_CN_QUESTION_LINE = re.compile(
    r"^\s*第\s*(?P<number>[1-9]\d{0,3})\s*题(?:\s*(?P<delimiter>[\uff1a:.\uFF0E\u3001)\uFF09])(?P<gap>\s*))?(?P<stem>.+)$"
)
_EMBEDDED_QUESTION_TRANSITION = re.compile(
    r"^(?P<head>.*?\S)\s+(?P<number>\d{1,4})\s*[\.．、)\uFF09]\s*(?P<stem>.+)$"
)
_GLUED_YEAR_QUESTION_TRANSITION = re.compile(
    r"^\s*(?P<prefix>\d{3})(?P<tail>\d)\s*[\.．、]\s*(?P<stem>(?:19|20)\d{2}.*)$"
)
_MALFORMED_NEXT_QUESTION_TAIL = re.compile(
    r"^\s*(?P<number>\d{4,5})\s*[\.．、)\uFF09]\s*(?P<stem>.+)$"
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
_ANSWER_OVERVIEW_HEADER_LINE = re.compile(r"^\s*答案速览\s*$")
_ANSWER_OVERVIEW_RANGE_LINE = re.compile(r"^\s*\d{1,3}\s*-\s*\d{1,3}\s*$")
_ANSWER_OVERVIEW_KEY_LINE = re.compile(r"^\s*[A-D]{1,8}\s*$", re.IGNORECASE)
_PAGE_NUMBER_ONLY_LINE = re.compile(r"^\s*\d{2,3}\s*$")
_SECTION_PART_MARKER = re.compile(r"^\s*[（(]?\s*[一二三四五六七八九十\d]+\s*[)）]?\s*$")
_PAREN_SECTION_PART_MARKER = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十\d〇零两]+)\s*[)）]?\s*$")
_OPEN_PAREN_SECTION_PART_MARKER = re.compile(r"^\s*[（(]\s*([一二三四五六七八九十\d〇零两]+)\s*$")
_CLOSE_PAREN_SECTION_PART_MARKER = re.compile(r"^\s*[)）]\s*$")


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


def _material_part_header_from_token(token: str) -> str | None:
    value = _nfkc((token or "").strip())
    if not value:
        return None
    if value.isdigit():
        return _material_header_label(int(value))
    if re.match(r"^[一二三四五六七八九十〇零两]+$", value):
        return f"材料{value}"
    return None


def _material_part_header_at(items: list[LineItem], index: int) -> tuple[int, str] | None:
    text, image_path = items[index]
    if image_path:
        return None
    line = _nfkc((text or "").strip())
    if not line or material_header_line(line) or generic_material_header_line(line):
        return None

    direct_match = _PAREN_SECTION_PART_MARKER.match(line)
    if direct_match:
        header = _material_part_header_from_token(direct_match.group(1))
        if header:
            return index, header

    open_match = _OPEN_PAREN_SECTION_PART_MARKER.match(line)
    if not open_match or index + 1 >= len(items):
        return None

    next_text, next_image = items[index + 1]
    if next_image:
        return None
    next_line = _nfkc((next_text or "").strip())
    if not _CLOSE_PAREN_SECTION_PART_MARKER.match(next_line):
        return None

    header = _material_part_header_from_token(open_match.group(1))
    if not header:
        return None
    return index + 1, header


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
    if generic_positions:
        if len(generic_positions) != 1 or generic_positions[0] == body_start:
            blocks = []
            for idx, start in enumerate(generic_positions, start=1):
                end = generic_positions[idx] if idx < len(generic_positions) else body_end
                blocks.append((start, end, _material_header_label(idx)))
            if blocks:
                return blocks

    part_positions: list[tuple[int, str]] = []
    i = body_start
    while i < body_end:
        marker = _material_part_header_at(items, i)
        if marker is None:
            i += 1
            continue
        marker_end_index, header = marker
        part_positions.append((marker_end_index, header))
        i = max(i + 1, marker_end_index + 1)

    if not part_positions:
        return []
    if len(part_positions) == 1 and part_positions[0][0] > body_start + 1:
        return []

    blocks = []
    for idx, (start, header) in enumerate(part_positions):
        end = part_positions[idx + 1][0] if idx + 1 < len(part_positions) else body_end
        blocks.append((start, end, header))
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
_QUESTION_QUERY_HINT = re.compile(
    r"(?:多少|几个|几类|几种|几家|几条|几项|哪年|约为|约是|可推出|说法|不正确|正确|错误|范围内)"
)
_LOCAL_NUMBER_REPAIR_MAX_SPAN = 12
_OPTION_LEFT_CONTEXT_CHARS = set(":：?？(（[【")
_OPTION_TAIL_PASSAGE_MARKER = re.compile(
    r"(?:以下是[^。\n]{0,80}?阅读之后回答\d{1,4}[—\-]\d{1,4}\s*题|阅读之后回答\d{1,4}[—\-]\d{1,4}\s*题)"
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
    # 「数量关系(共15 题，参考时限…)」带科目名的是篇题，不能当 boilerplate 跳过。
    if "参考时限" in s and "共" in s and "题" in s and not _detect_subject_section_kind(s):
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
    normalized = _nfkc(_normalize_digits((line or "").strip()))
    return bool(
        _QUESTION_NO_LINE.match(normalized)
        or _QUESTION_NO_OPEN_BRACKET_LINE.match(normalized)
    )


def _extract_question_no(line: str) -> str:
    s = _nfkc(_normalize_digits((line or "").strip()))
    m = re.match(r"^\s*(\d{1,4})", s)
    return m.group(1) if m else ""


def _strip_stem_leading_fillers(stem: str) -> str:
    return _nfkc(_normalize_digits((stem or "").strip())).lstrip(_LEADING_STEM_FILLERS)


def _find_spaced_bare_option_markers(text: str) -> list[re.Match[str]]:
    matches: list[re.Match[str]] = []
    for match in _BARE_OPTION_MARKER.finditer(text):
        if match.start() == 0 or text[match.start() - 1].isspace():
            matches.append(match)
    return matches


def _option_marker_has_valid_left_context(text: str, start: int) -> bool:
    if start <= 0:
        return True
    previous = text[start - 1]
    return previous.isspace() or previous in _OPTION_LEFT_CONTEXT_CHARS


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
        options = _options_from_marker_matches(text, chosen)
        embedded_marker_penalty = 0
        short_body_penalty = 0
        leading_punct_penalty = 0
        body_length_score = 0
        for option in options:
            body = (option.text or "").strip()
            body_length_score += min(len(body), 40)
            if body and len(body) <= 3:
                short_body_penalty += 1
            if body[:1] in ".,，。:：;；)）]】":
                leading_punct_penalty += 1
            embedded_marker_penalty += sum(
                1
                for marker in OPTION_MARKER.finditer(body)
                if marker.group(1).upper() in {"A", "B", "C", "D"}
            )
        last_body = (options[-1].text or "").strip()
        last_body_clean = 0 if any(
            marker.group(1).upper() in {"A", "B", "C", "D"} for marker in OPTION_MARKER.finditer(last_body)
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


def _options_from_marker_matches(text: str, matches: list[re.Match[str]]) -> list[ExamOption]:
    options: list[ExamOption] = []
    for index, match in enumerate(matches):
        body_start = match.end()
        body_end = matches[index + 1].start() if index + 1 < len(matches) else len(text)
        body = text[body_start:body_end].strip()
        marker = _OPTION_TAIL_PASSAGE_MARKER.search(body)
        if marker and body[: marker.start()].strip():
            body = body[: marker.start()].rstrip()
        options.append(ExamOption(letter=match.group(1).upper(), text=body))
    return options


def _match_looks_like_embedded_list_marker(text: str, match: re.Match[str]) -> bool:
    if _option_marker_has_valid_left_context(text, match.start()):
        return False
    token = text[match.start() : match.end()]
    return "、" in token


def _match_single_option_line(text: str) -> re.Match[str] | None:
    normalized_text = re.sub(
        r"^\s*([ABCDabcd])\s*[,，]\s*(?=[^\d\-])",
        r"\1．",
        text,
    )
    punctuated = _SINGLE_OPTION.match(normalized_text)
    if punctuated:
        return punctuated

    bare = _BARE_SINGLE_OPTION.match(normalized_text)
    if not bare:
        glued = _GLUED_SINGLE_OPTION.match(normalized_text)
        if not glued:
            glued = _GLUED_QUOTED_OPTION.match(normalized_text)
        if not glued:
            return None
        bare = glued

    trailing_markers = [
        match for match in _BARE_OPTION_MARKER.finditer(normalized_text) if match.start() > bare.start()
    ]
    if trailing_markers:
        return None
    return bare


def _looks_like_prefixed_entity_stem_line(
    items: list[LineItem],
    start: int,
    end: int,
) -> bool:
    """
    避免把「A、B 两地……」这类题干首句当成 A 选项。

    仅在当前行本身像单选项、正文又以另一个字母实体开头，
    且后文还出现了一组真正的 A/B/C/D 选项时触发。
    """
    if start >= end:
        return False
    text, image_path = items[start]
    if image_path:
        return False
    normalized = _normalize_pdf_option_text((text or "").strip())
    if not normalized:
        return False
    match = _match_single_option_line(normalized)
    if not match:
        return False

    body = (match.group(2) or "").strip()
    leading_letter = match.group(1).upper()
    looks_like_entity_enumeration = bool(re.match(r"^[A-Z](?:\s+|[、,，和及与])", body))
    looks_like_entity_token = bool(re.match(r"^[A-Z][A-Za-z0-9.\-]{2,}", body))
    if not looks_like_entity_enumeration and not looks_like_entity_token:
        return False

    seen_letters: set[str] = set()
    for index in range(start + 1, min(end, start + 12)):
        next_text, next_image = items[index]
        if next_image:
            continue
        next_raw = (next_text or "").strip()
        if not next_raw or _is_parser_noise_text(next_raw):
            continue
        next_match = _match_single_option_line(_normalize_pdf_option_text(next_raw))
        if not next_match:
            continue
        seen_letters.add(next_match.group(1).upper())
    return leading_letter in seen_letters and {"A", "B", "C", "D"}.issubset(seen_letters)


def _looks_like_year_leading_stem(stem: str) -> bool:
    s = _nfkc(_normalize_digits((stem or "").strip()))
    return bool(re.match(r"^(?:19|20)\d{2}(?:\D|$)", s))


def _looks_like_numbered_intro_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    return bool(
        re.match(
            r"^(?:[1-9](?![万千百十亿年月日个家项次位人名天周时分秒米吨亩元票条份])[\u4e00-\u9fff]|[1-9]\s*[)）]\s*\S|[1-9]\s*[\"“'‘《])",
            s,
        )
    )


def _looks_like_numeric_sequence_stem(stem: str) -> bool:
    s = _nfkc(_normalize_digits((stem or "").strip()))
    s = re.sub(r"\s+", "", s)
    if not s or len(s) > 48:
        return False
    if not _NUMERIC_SEQUENCE_STEM.match(s):
        return False
    digit_count = sum(ch.isdigit() for ch in s)
    punctuation_count = sum(ch in ",，.．-+＋×xX÷/\\()[]（）:：;；%=％<>≤≥≈~～" for ch in s)
    return digit_count >= 4 and punctuation_count >= 2


def _looks_like_numeric_sequence_fragment_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    if not s or len(s) > 12:
        return False
    if not re.match(r"^[\d\.\-＋\+×xX÷/\\,，()（）]+$", s):
        return False
    digit_count = sum(ch.isdigit() for ch in s)
    punctuation_count = sum(ch in ",，.．-+＋×xX÷/\\()（）" for ch in s)
    return digit_count >= 1 and punctuation_count >= 1


def _looks_like_embedded_numeric_continuation_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    if not s:
        return False
    if (
        _looks_like_numeric_sequence_stem(s)
        or _looks_like_numeric_sequence_fragment_stem(s)
        or _looks_like_digit_leading_chinese_stem(s)
    ):
        return True
    return bool(re.match(r"^(?:世纪|年代|毫焦耳|焦耳|度(?:[;；,，。.]|$))", s))


def _looks_like_quantity_measure_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    if not s:
        return False
    return bool(_LEADING_QUANTITY_MEASURE_STEM.match(s))


def _looks_like_plain_number_continuation_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    if not s:
        return False
    return bool(
        re.match(
            r"^(?:家|个|人|名|项|户|条|次|年|月|日|周|天|小时|分钟|秒|米|公里|千米|亿元|万元|元|倍|个百分点)(?:[\u4e00-\u9fffA-Za-z]|$)",
            s,
        )
    )


def _looks_like_digit_leading_chinese_stem(stem: str) -> bool:
    s = _strip_stem_leading_fillers(stem)
    return bool(re.match(r"^\d{1,4}\s*[\u4e00-\u9fff]", s))


def _is_parser_noise_text(text: str) -> bool:
    s = _nfkc((text or "").strip())
    if not s:
        return True
    if _SCAN_AD_LINE.match(s):
        return True
    if _TOPIC_HEADER_LINE.match(s):
        return True
    return False


def _is_answer_overview_content_line(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s:
        return True
    compact = re.sub(r"\s+", "", s.upper())
    if _ANSWER_OVERVIEW_RANGE_LINE.match(s):
        return True
    if _ANSWER_OVERVIEW_KEY_LINE.match(compact):
        return True
    return bool(re.fullmatch(r"\d{1,3}", compact))


def _looks_like_section_reset_heading(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s:
        return False
    if s.startswith(("第", "第一节", "第二节", "第三节", "第四节")):
        return True
    return any(token in s for token in ("基础夯实", "进阶提升", "高难突破", "套题特训"))


def _is_probable_page_footer_number(items: list[LineItem], index: int) -> bool:
    text, image_path = items[index]
    if image_path:
        return False
    current = _nfkc(_normalize_digits((text or "").strip()))
    if not _PAGE_NUMBER_ONLY_LINE.match(current):
        return False

    current_number = int(current)
    if current_number < 20:
        return False

    prev_texts: list[str] = []
    prev_index = index - 1
    while prev_index >= 0 and len(prev_texts) < 3:
        probe_text, probe_image = items[prev_index]
        if probe_image:
            prev_index -= 1
            continue
        candidate = _normalize_pdf_option_text((probe_text or "").strip())
        if candidate:
            prev_texts.append(candidate)
        prev_index -= 1

    next_texts: list[str] = []
    next_index = index + 1
    while next_index < len(items) and len(next_texts) < 3:
        probe_text, probe_image = items[next_index]
        if probe_image:
            next_index += 1
            continue
        candidate = _normalize_pdf_option_text((probe_text or "").strip())
        if candidate:
            next_texts.append(candidate)
        next_index += 1

    if not prev_texts or not next_texts:
        return False

    prev_text = prev_texts[0]
    next_text = next_texts[0]
    prev_option = next((match for text in prev_texts if (match := _match_single_option_line(text)) is not None), None)
    next_option = next((match for text in next_texts if (match := _match_single_option_line(text)) is not None), None)
    prev_question = next(
        (
            match
            for text in prev_texts
            if (match := _match_leading_question_with_stem(text)) is not None and match[0].isdigit()
        ),
        None,
    )
    if prev_option and next_option:
        prev_letter = prev_option.group(1).upper()
        next_letter = next_option.group(1).upper()
        if ord(next_letter) == ord(prev_letter) + 1:
            return True
    if prev_question is not None and next_option is not None:
        prev_number = int(prev_question[0])
        if prev_number <= 5 and current_number > prev_number + 10:
            return True

    if any(_looks_like_digit_leading_chinese_stem(text) for text in prev_texts) and (
        _looks_like_digit_leading_chinese_stem(next_text) or next_option is not None
    ):
        return True

    if _is_question_no_line(next_text):
        next_number_raw = _extract_question_no(next_text)
        next_number = int(next_number_raw) if next_number_raw.isdigit() else None
        if next_number is not None and next_number <= 5 and any(
            _looks_like_section_reset_heading(text) for text in prev_texts
        ):
            return True

    next_question = _match_leading_question_with_stem(next_text)
    if next_question is not None and next_question[0].isdigit():
        next_number = int(next_question[0])
        if next_number <= 5 and any(_looks_like_section_reset_heading(text) for text in prev_texts):
            return True
        if current_number > next_number + 5 and (
            prev_option is not None
            or any(_looks_like_question_candidate_text(text) for text in prev_texts)
            or any(_looks_like_digit_leading_chinese_stem(text) for text in prev_texts)
        ):
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


def _looks_like_question_candidate_text(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s or _looks_like_numeric_sequence_stem(s):
        return False
    if s in {"缺失", "题干缺失", "图片缺失"}:
        return True
    if _looks_like_question_stem_text(s):
        return True
    if _QUESTION_QUERY_HINT.search(s):
        return True
    if re.search(r"[（(]\s*[)）]", s):
        return True
    if s.endswith(("?", "？", ":", "：")):
        return True
    if re.search(r"[（(]\s*[）)]?\s*$", s):
        return True
    return False


def _looks_like_split_question_stem_placeholder(text: str) -> bool:
    s = _strip_stem_leading_fillers(text)
    if not s:
        return False
    if any(token in s for token in ("题干", "下一题", "材料说明", "看图", "图形", "排序")):
        return True
    if (
        len(s) <= 12
        and re.search(r"[\u4e00-\u9fff]", s)
        and not s.endswith(("。", "．", ".", "，", ",", "；", ";", "：", ":", "？", "?"))
    ):
        return True
    return False


def _looks_like_number_only_marker_payload_text(text: str) -> bool:
    s = _strip_stem_leading_fillers(text)
    if not s or _starts_new_question_line(s):
        return False
    if _looks_like_numeric_sequence_stem(s) or _looks_like_numeric_sequence_fragment_stem(s):
        return True
    if _looks_like_question_candidate_text(s) or _looks_like_split_question_stem_placeholder(s):
        return True
    return len(s) >= 12 and bool(re.search(r"[\u4e00-\u9fff]", s))


def _number_only_marker_has_question_payload(items: list[LineItem], marker_index: int, upper_bound: int) -> bool:
    saw_stem_like_text = False
    consumed_text_lines = 0
    stem_texts: list[str] = []
    probe = marker_index + 1
    while probe < upper_bound:
        text, image_path = items[probe]
        if image_path:
            return True
        normalized = _normalize_pdf_option_text((text or "").strip())
        if not normalized or _is_parser_noise_text(normalized):
            probe += 1
            continue
        if _starts_new_question_line(normalized):
            return False
        if _has_option_markers_pdf(normalized):
            starts_with_explicit_option = bool(_match_single_option_line(normalized))
            parsed_inline_options = _parse_options_line(normalized)
            has_inline_option_cluster = parsed_inline_options is not None and len(parsed_inline_options) >= 4
            merged_text = _strip_stem_leading_fillers("".join(stem_texts))
            if saw_stem_like_text:
                return True
            if (
                consumed_text_lines > 0
                and _looks_like_number_only_marker_payload_text(merged_text)
                and (starts_with_explicit_option or has_inline_option_cluster)
            ):
                return True
            if starts_with_explicit_option or has_inline_option_cluster:
                return False
        if (
            material_header_line(normalized)
            or generic_material_header_line(normalized)
            or _detect_subject_section_kind(normalized)
            or _is_other_section_title(normalized)
        ):
            return False
        consumed_text_lines += 1
        stem_texts.append(normalized)
        merged_text = _strip_stem_leading_fillers("".join(stem_texts))
        if _looks_like_split_question_stem_placeholder(normalized):
            saw_stem_like_text = True
            probe += 1
            continue
        if _looks_like_number_only_marker_payload_text(merged_text):
            saw_stem_like_text = True
            probe += 1
            continue
        if consumed_text_lines >= 5:
            return False
        probe += 1
    return saw_stem_like_text


def _has_question_prompt_signal(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s:
        return False
    if _QUESTION_PROMPT_HINT.match(s):
        return True
    if _QUESTION_QUERY_HINT.search(s):
        return True
    if re.search(r"[（(]\s*[)）]", s):
        return True
    return s.endswith(("?", "？", ":", "："))


def _has_terminal_question_signal(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s:
        return False
    if s.endswith(("?", "？", ":", "：")):
        return True
    tail = s[-48:]
    return bool(
        re.search(
            r"(?:以下哪项|以下哪个|下列哪项|下列哪个|下列说法|最适合|最能|最不|意在说明|最恰当的一项是|依次填入|可推出|不正确的是|正确的是|错误的是|属于|不属于|有几项|的是)\s*$",
            tail,
        )
    )


def _rich_lines_look_like_question_text(lines: list[RichLine]) -> bool:
    texts = [_rich_line_text_value(line).strip() for line in lines]
    merged = " ".join(text for text in texts if text)
    if not merged:
        return False
    return _looks_like_question_candidate_text(merged)


def _looks_like_material_intro_text(text: str) -> bool:
    s = _nfkc(_normalize_digits((text or "").strip()))
    if not s:
        return False
    if _starts_new_question_line(s) or _has_option_markers_pdf(s):
        return False
    if material_header_line(s) or generic_material_header_line(s):
        return False
    if re.search(r"第\s*[一二三四五六七八九十\d]+\s*组材料", s):
        return True
    if "材料说明" in s or "根据材料" in s:
        return True
    if "材料" in s and len(s) >= 8 and not _looks_like_question_stem_text(s):
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
    if _is_probable_page_footer_number(items, index + 1):
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
            delimiter = match.groupdict().get("delimiter", "")
            gap = match.groupdict().get("gap", " ")
            stem = (match.group("stem") or "").strip()
            head = _strip_stem_leading_fillers(stem)
            if not head:
                continue
            if head[0].isdigit() and not gap and not (
                _looks_like_year_leading_stem(head)
                or _looks_like_numbered_intro_stem(head)
                or _looks_like_numeric_sequence_stem(head)
                or _looks_like_numeric_sequence_fragment_stem(head)
                or _looks_like_digit_leading_chinese_stem(head)
            ):
                continue
            if pattern is _LEADING_PLAIN_QUESTION_LINE and _looks_like_quantity_measure_stem(head):
                continue
            if pattern is _LEADING_PLAIN_QUESTION_LINE and _looks_like_plain_number_continuation_stem(head):
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
    # 「数量关系(共15 题，参考时限15 分钟)」式篇题：科目名后紧跟「(共N题…」。
    if re.match(rf"^(?:{label_pattern})\s*[（(]\s*共\s*\d+\s*[题道]", s):
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
    if any(_looks_like_material_intro_text(text) for text in texts):
        return True
    image_count = sum(1 for line in meaningful if _rich_line_has_image(line))
    total_chars = sum(len(text) for text in texts)
    long_lines = sum(1 for text in texts if len(text) >= 18)

    if image_count >= 1 and not texts:
        return True
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
        text = _normalize_pdf_option_text(_rich_line_text_value(line))
        line_letters: list[str] = []
        if text:
            parsed = _parse_options_line(text)
            if parsed:
                line_letters = [option.letter.upper() for option in parsed]
            else:
                letter = _rich_option_letter(line)
                if letter:
                    line_letters = [letter]
        if (
            line_letters
            and expected_index < len(letters)
            and line_letters == list(letters[expected_index : expected_index + len(line_letters)])
        ):
            last_option_index = idx
            expected_index += len(line_letters)

    if expected_index < len(letters) or last_option_index < 0 or last_option_index >= len(lines) - 1:
        return list(lines), []

    spill = list(lines[last_option_index + 1 :])
    if not _looks_like_material_intro_lines(spill):
        return list(lines), []
    return list(lines[: last_option_index + 1]), spill


def _clone_rich_lines(lines: list[RichLine]) -> list[RichLine]:
    return [RichLine(parts=list(line.parts)) for line in lines]


def _extract_full_option_cluster_from_rich_lines(
    lines: list[RichLine],
    start: int,
) -> tuple[list[RichLine], int] | None:
    if start >= len(lines):
        return None
    text = _normalize_pdf_option_text(_rich_line_text_value(lines[start]))
    if not text or _rich_line_has_image(lines[start]):
        return None

    parsed = _parse_options_line(text)
    if parsed and [option.letter.upper() for option in parsed[:4]] == ["A", "B", "C", "D"]:
        return _clone_rich_lines(_options_to_rich_lines(parsed[:4])), start + 1

    cluster: list[RichLine] = []
    expected_letters = ["A", "B", "C", "D"]
    index = start
    while index < len(lines) and len(cluster) < len(expected_letters):
        line = lines[index]
        if _rich_line_has_image(line):
            return None
        letter = _rich_option_letter(line)
        if letter != expected_letters[len(cluster)]:
            return None
        cluster.append(RichLine(parts=list(line.parts)))
        index += 1
    if len(cluster) != len(expected_letters):
        return None
    return cluster, index


def _extract_option_clusters_from_rich_lines(
    lines: list[RichLine],
) -> tuple[list[RichLine], list[list[RichLine]]]:
    remaining: list[RichLine] = []
    clusters: list[list[RichLine]] = []
    index = 0
    while index < len(lines):
        cluster = _extract_full_option_cluster_from_rich_lines(lines, index)
        if cluster is not None:
            cluster_lines, next_index = cluster
            clusters.append(cluster_lines)
            index = next_index
            continue
        remaining.append(RichLine(parts=list(lines[index].parts)))
        index += 1
    return remaining, clusters


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
    normalized = re.sub(
        r"(?<![A-Za-z])([A-D])\s*[,，]\s*(?=[^\d\-])",
        r"\1．",
        normalized,
    )
    all_matches = [
        match
        for match in OPTION_MARKER.finditer(normalized)
        if match.group(1).upper() in {"A", "B", "C", "D"}
    ]
    marker_matches = _best_ordered_option_match_sequence(normalized, all_matches)
    if len(marker_matches) < 2:
        marker_matches = [
            match for match in all_matches if _option_marker_has_valid_left_context(normalized, match.start())
        ]
    raw = _options_from_marker_matches(normalized, marker_matches)
    if len(raw) >= 2 and raw[0].letter != "A":
        suspicious_tail = any(
            (option.text or "").strip()[:1] in ".,，。:：;；)）]】"
            for option in raw[1:]
        )
        suspicious_embedded_marker = any(
            _match_looks_like_embedded_list_marker(normalized, match)
            for match in marker_matches[1:]
        )
        if suspicious_tail or suspicious_embedded_marker:
            raw = []
    if len(raw) >= 2:
        return raw

    bare_matches = _select_ordered_option_matches(_find_spaced_bare_option_markers(normalized))
    raw = _options_from_marker_matches(normalized, bare_matches)
    if len(raw) >= 2:
        return raw

    single = _match_single_option_line(normalized)
    if not single:
        return None
    return [ExamOption(letter=single.group(1).upper(), text=(single.group(2) or "").strip())]


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
        if expected_index == 0 and _looks_like_prefixed_entity_stem_line(items, i, end):
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

    if _looks_like_prefixed_entity_stem_line(items, start, end):
        return None

    # 避免把题干与下一行选项拼在一起误判为「四个选项」
    if not _has_option_markers_pdf(first):
        return None

    starts_with_option = bool(re.match(r"^\s*[A-Da-d]", _normalize_pdf_option_text(first)))
    opts = _parse_options_line(first) if starts_with_option else None
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


def _segment_to_rich_lines(
    segment: list[LineItem],
    *,
    strip_question_prefix: bool,
) -> list[RichLine]:
    out: list[RichLine] = []
    prefix_consumed = False
    for t, img in segment:
        if img:
            out.append(_rich_img(img))
            continue
        stripped = (t or "").strip()
        if not stripped:
            continue
        if strip_question_prefix:
            if not prefix_consumed and _is_question_no_line(stripped):
                prefix_consumed = True
                continue
            matched = _match_leading_question_with_stem(stripped)
            if matched and not prefix_consumed:
                _number, stem = matched
                out.append(_rich_text(stem))
                prefix_consumed = True
                continue
        out.append(_rich_text(stripped))
    return out


def _segment_to_stem_only(segment: list[LineItem]) -> list[RichLine]:
    return _segment_to_rich_lines(segment, strip_question_prefix=True)


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
    supporting_text_count = 0

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
            continue
        if len(text) <= 8:
            supporting_text_count += 1

    if not letters_seen:
        return False
    if nonempty_text_count > 0:
        return True
    if supporting_text_count > 0:
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


def _collect_blank_option_placeholder_cluster(
    items: list[LineItem],
    start: int,
    end: int,
) -> tuple[int, int, list[RichLine]] | None:
    for index in range(start, end):
        seq = _collect_sequential_option_lines(items, index, end)
        if seq is None:
            continue
        seq_end, seq_lines = seq
        if any(_rich_line_has_image(line) for line in seq_lines):
            continue
        if _blank_option_placeholder_letters(seq_lines) != ["A", "B", "C", "D"]:
            continue
        return index, seq_end, seq_lines
    return None


def _question_number_has_leading_visual_payload(
    items: list[LineItem],
    index: int,
    lower_bound: int,
) -> bool:
    image_count = 0
    probe = index - 1
    while probe >= lower_bound:
        text, image_path = items[probe]
        if image_path:
            image_count += 1
            probe -= 1
            continue
        raw = (text or "").strip()
        if not raw:
            probe -= 1
            continue
        normalized = _normalize_pdf_option_text(raw)
        if _starts_new_question_line(normalized):
            return False
        break
    return image_count > 0


def _cluster_can_rebalance_from_preceding_images(
    items: list[LineItem],
    start: int,
    option_lines: list[RichLine],
) -> bool:
    if _blank_option_placeholder_letters(option_lines) != ["A", "B", "C", "D"]:
        return False

    image_count = sum(1 for line in option_lines if _rich_line_has_image(line))
    recent_question_no = False
    index = start - 1
    while index >= 0:
        text, image_path = items[index]
        if image_path:
            image_count += 1
            index -= 1
            continue
        raw = (text or "").strip()
        if raw:
            recent_question_no = _is_question_no_line(_normalize_pdf_option_text(raw))
            break
        index -= 1
    if image_count >= len(option_lines):
        return True
    return recent_question_no and image_count >= len(option_lines) - 1


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

    glued_year_match = _GLUED_YEAR_QUESTION_TRANSITION.match(normalized)
    if glued_year_match:
        prefix = int(glued_year_match.group("prefix"))
        tail = glued_year_match.group("tail")
        candidate = prefix - 1
        if 1 <= candidate <= 999 and str(candidate).endswith(tail):
            stem = (glued_year_match.group("stem") or "").strip()
            return [f"{candidate}.", stem] if stem else [f"{candidate}."]

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
            if _is_question_no_line(head) and _looks_like_embedded_numeric_continuation_stem(stem):
                updated.append(piece)
                continue
            if re.match(
                r"^\d(?:\s*(?:万人|亿元|万亿元|个百分点|%|％|个|家|条|项|米|吨|户|人次))",
                normalized_stem,
            ):
                updated.append(piece)
                continue
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
        if _ANSWER_OVERVIEW_HEADER_LINE.match(_nfkc(_normalize_digits(normalized_text))):
            i += 1
            while i < len(items):
                next_text, next_image = items[i]
                if next_image:
                    i += 1
                    continue
                next_normalized = _nfkc(_normalize_digits((next_text or "").strip()))
                if not next_normalized:
                    i += 1
                    continue
                if _is_answer_overview_content_line(next_normalized):
                    i += 1
                    continue
                break
            continue
        if _is_probable_page_footer_number(items, i):
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
    return _dedupe_adjacent_question_number_lines(_merge_numeric_continuation_fragments(normalized_items))


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


def _dedupe_adjacent_question_number_lines(items: list[LineItem]) -> list[LineItem]:
    deduped: list[LineItem] = []
    i = 0
    while i < len(items):
        text, image_path = items[i]
        if image_path is None and i + 1 < len(items):
            next_text, next_image = items[i + 1]
            current = _nfkc((text or "").strip())
            nxt = _nfkc((next_text or "").strip())
            if (
                next_image is None
                and _is_question_no_line(current)
                and _is_question_no_line(nxt)
                and _extract_question_no(current) == _extract_question_no(nxt)
            ):
                deduped.append((current, None))
                i += 2
                continue
        deduped.append(items[i])
        i += 1
    return deduped


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
            matched = _match_leading_question_with_stem(normalized)
            prev_nonempty_text = ""
            prev_prev_nonempty_text = ""
            back = i - 1
            while back >= a:
                back_text, back_image = items[back]
                if back_image:
                    back -= 1
                    continue
                candidate = _normalize_pdf_option_text((back_text or "").strip())
                if candidate:
                    if not prev_nonempty_text:
                        prev_nonempty_text = candidate
                    else:
                        prev_prev_nonempty_text = candidate
                        break
                back -= 1
            next_nonempty_text = ""
            probe_next = i + 1
            while probe_next < b:
                next_text, next_image = items[probe_next]
                if next_image:
                    probe_next += 1
                    continue
                candidate = _normalize_pdf_option_text((next_text or "").strip())
                if candidate:
                    next_nonempty_text = candidate
                    break
                probe_next += 1
            if prev_nonempty_text in {"注", "注:", "注："}:
                continue
            if (
                _is_question_no_line(normalized)
                and prev_nonempty_text.endswith(("、", ",", "，", ";", "；"))
                and re.match(r"^\d{1,3}\s*[\u4e00-\u9fff]", _strip_stem_leading_fillers(next_nonempty_text))
            ):
                continue
            if (
                _is_question_no_line(normalized)
                and _has_option_markers_pdf(prev_nonempty_text)
                and not _question_number_has_leading_visual_payload(items, i, a)
                and not (
                    _looks_like_question_candidate_text(_strip_stem_leading_fillers(next_nonempty_text))
                    or _number_only_marker_has_question_payload(items, i, b)
                )
            ):
                continue
            if (
                matched
                and _is_question_no_line(prev_nonempty_text)
                and prev_prev_nonempty_text.endswith(("、", ",", "，", ";", "；"))
                and _looks_like_digit_leading_chinese_stem(matched[1])
            ):
                continue
            current_number = None
            suspicious_numeric_tail = False
            if matched and matched[0].isdigit():
                current_number = int(matched[0])
                suspicious_numeric_tail = (
                    _looks_like_digit_leading_chinese_stem(matched[1])
                    or _looks_like_numeric_sequence_fragment_stem(matched[1])
                ) and not _looks_like_question_candidate_text(matched[1])
            elif _is_question_no_line(normalized):
                value = _extract_question_no(normalized)
                current_number = int(value) if value.isdigit() else None
                next_text = ""
                probe = i + 1
                while probe < b:
                    probe_text, probe_image = items[probe]
                    if probe_image:
                        probe += 1
                        continue
                    next_text = _strip_stem_leading_fillers((probe_text or "").strip())
                    if next_text:
                        break
                    probe += 1
                suspicious_numeric_tail = bool(
                    re.match(
                        r"^\d+(?:\.\d+)?(?:\s*(?:万人|亿元|万亿元|个百分点|%|％|个|家|条|项|米|吨|户|人次))",
                        next_text,
                    )
                ) and not _looks_like_question_candidate_text(next_text)
            if suspicious_numeric_tail and current_number is not None:
                last_marker_number = None
                if markers:
                    probe = markers[-1]
                    while probe < i:
                        probe_text, probe_image = items[probe]
                        if not probe_image:
                            probe_raw = _normalize_pdf_option_text((probe_text or "").strip())
                            if _is_question_no_line(probe_raw):
                                value = _extract_question_no(probe_raw)
                                last_marker_number = int(value) if value.isdigit() else None
                                break
                            probe_matched = _match_leading_question_with_stem(probe_raw)
                            if probe_matched and probe_matched[0].isdigit():
                                last_marker_number = int(probe_matched[0])
                                break
                        probe += 1
                prev_context_has_text = any(
                    idx >= a
                    and items[idx][1] is None
                    and (items[idx][0] or "").strip()
                    for idx in range(max(a, i - 3), i)
                )
                if last_marker_number is None and (current_number > 200 or prev_context_has_text):
                    continue
                if (
                    last_marker_number is not None
                    and current_number > last_marker_number + 20
                ):
                    continue
            prev_is_question_no = False
            prev_question_no = ""
            if i - 1 >= a:
                prev_text, prev_image = items[i - 1]
                prev_raw = (prev_text or "").strip()
                prev_normalized = _normalize_pdf_option_text(prev_raw)
                prev_is_question_no = not prev_image and prev_raw and _is_question_no_line(prev_normalized)
                if prev_is_question_no:
                    prev_question_no = _extract_question_no(prev_normalized)
                if matched and prev_is_question_no and matched[0] != prev_question_no:
                    prev_number = int(prev_question_no) if prev_question_no.isdigit() else None
                    if prev_number is None or current_number is None or current_number != prev_number + 1:
                        continue
                if (
                    matched
                    and not prev_image
                    and prev_raw
                    and prev_is_question_no
                    and prev_question_no == matched[0]
                ):
                    marker_start = i - 1
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
    spill_lines: list[RichLine] = []
    spill_lines.extend(_rich_text(text) for text in tail_texts)
    spill_lines.extend(_rich_img(image_path) for image_path in tail_images)
    if spill_lines and (
        _looks_like_material_intro_lines(spill_lines)
        or any(_looks_like_material_intro_text(text) for text in tail_texts)
    ):
        return updated + spill_lines

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


def _expand_option_rich_line(line: RichLine) -> list[RichLine]:
    text = _normalize_pdf_option_text(_rich_line_text_value(line))
    if not text:
        return []
    parsed = _parse_options_line(text)
    if parsed:
        return [RichLine(parts=[(f"{option.letter}．{option.text}", None)]) for option in parsed]
    if _rich_option_letter(line):
        return [RichLine(parts=list(line.parts))]
    return []


def _split_rich_line_with_embedded_options(
    line: RichLine,
    *,
    allow_single_option_line: bool = False,
) -> tuple[list[RichLine], list[RichLine]] | None:
    if _rich_line_has_image(line):
        return None
    text = _normalize_pdf_option_text(_rich_line_text_value(line))
    if not text:
        return None
    parsed = _parse_options_line(text)
    if not parsed or len(parsed) < 4:
        if not allow_single_option_line:
            return None
        single = _expand_option_rich_line(line)
        return ([], single) if single else None

    matches = list(OPTION_MARKER.finditer(text))
    if len(matches) < 2:
        option_lines = [RichLine(parts=[(f"{option.letter}．{option.text}", None)]) for option in parsed]
        return [], option_lines

    prefix = text[: matches[0].start()].strip()
    stem_lines = [RichLine(parts=[(prefix, None)])] if prefix else []
    option_lines = [RichLine(parts=[(f"{option.letter}．{option.text}", None)]) for option in parsed]
    return stem_lines, option_lines


def _extract_option_like_rich_lines(items: list[LineItem]) -> list[RichLine]:
    lines: list[RichLine] = []
    for text, image_path in items:
        if image_path:
            continue
        raw = (text or "").strip()
        if not raw or _is_parser_noise_text(raw):
            continue
        rich = RichLine(parts=[(raw, None)])
        expanded = _expand_option_rich_line(rich)
        if expanded:
            lines.extend(expanded)
    return lines


def _group_option_rich_line_clusters(lines: list[RichLine]) -> list[list[RichLine]]:
    expanded: list[RichLine] = []
    for line in lines:
        expanded.extend(_expand_option_rich_line(line))

    clusters: list[list[RichLine]] = []
    current: list[RichLine] = []
    for line in expanded:
        letter = _rich_option_letter(line)
        if not letter:
            continue
        if not current:
            current = [line] if letter == "A" else []
            continue
        if letter == "A":
            clusters.append(current)
            current = [line]
            continue
        current.append(line)
    if current:
        clusters.append(current)
    return clusters


def _merge_option_rich_lines(*groups: list[RichLine]) -> list[RichLine]:
    merged_by_letter: dict[str, RichLine] = {}
    for group in groups:
        for line in group:
            letter = _rich_option_letter(line)
            if not letter or letter in merged_by_letter:
                continue
            merged_by_letter[letter] = RichLine(parts=list(line.parts))
    if set(merged_by_letter) != {"A", "B", "C", "D"}:
        return []
    return [merged_by_letter[letter] for letter in "ABCD"]


def _split_embedded_option_lines_from_stem(
    stem_lines: list[RichLine],
    *,
    allow_single_option_lines: bool = False,
) -> tuple[list[RichLine], list[RichLine]]:
    clean_stem: list[RichLine] = []
    embedded_option_lines: list[RichLine] = []
    for line in stem_lines:
        split = _split_rich_line_with_embedded_options(
            line,
            allow_single_option_line=allow_single_option_lines,
        )
        if split is not None:
            stem_part, option_part = split
            clean_stem.extend(stem_part)
            embedded_option_lines.extend(option_part)
            continue
        clean_stem.append(line)
    return clean_stem, embedded_option_lines


def _redistribute_orphan_option_lines(
    questions: list[ExamQuestion],
    orphan_option_lines: list[RichLine],
) -> None:
    if not questions or not orphan_option_lines:
        return
    orphan_clusters = _group_option_rich_line_clusters(orphan_option_lines)
    if not orphan_clusters:
        return

    embedded_by_index: dict[int, list[RichLine]] = {}
    for index, question in enumerate(questions):
        clean_stem, embedded = _split_embedded_option_lines_from_stem(
            question.stem_lines,
            allow_single_option_lines=True,
        )
        question.stem_lines = clean_stem
        if embedded:
            embedded_by_index[index] = embedded

    for index, question in enumerate(questions):
        if question.option_lines:
            continue
        embedded = embedded_by_index.get(index, [])
        if not embedded:
            continue
        merged_embedded = _merge_option_rich_lines(embedded)
        if merged_embedded:
            question.option_lines = merged_embedded
            continue
        embedded_letters = {letter for letter in (_rich_option_letter(line) for line in embedded) if letter}
        best_cluster_index: int | None = None
        best_cluster_size: int | None = None
        for cluster_index, cluster in enumerate(orphan_clusters):
            cluster_letters = {letter for letter in (_rich_option_letter(line) for line in cluster) if letter}
            if cluster_letters & embedded_letters:
                continue
            if cluster_letters | embedded_letters != {"A", "B", "C", "D"}:
                continue
            cluster_size = len(cluster)
            if best_cluster_index is None or cluster_size < best_cluster_size:
                best_cluster_index = cluster_index
                best_cluster_size = cluster_size
        if best_cluster_index is None:
            continue
        merged = _merge_option_rich_lines(embedded, orphan_clusters.pop(best_cluster_index))
        if merged:
            question.option_lines = merged

    for question in questions:
        if question.option_lines:
            continue
        for cluster_index, cluster in enumerate(orphan_clusters):
            cluster_letters = {letter for letter in (_rich_option_letter(line) for line in cluster) if letter}
            if cluster_letters != {"A", "B", "C", "D"}:
                continue
            merged = _merge_option_rich_lines(cluster)
            if merged:
                question.option_lines = merged
                orphan_clusters.pop(cluster_index)
            break


def _repair_page_footer_split_questions(questions: list[ExamQuestion]) -> None:
    index = 1
    while index < len(questions) - 1:
        previous = questions[index - 1]
        current = questions[index]
        nxt = questions[index + 1]
        previous_number = _parse_numeric_source_number(previous.source_number)
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if (
            previous_number is None
            or current_number is None
            or next_number is None
            or previous.option_lines
            or not current.option_lines
            or next_number != previous_number + 1
            or current_number < next_number + 5
        ):
            index += 1
            continue

        current_stem_texts = [
            _rich_line_text_value(line).strip()
            for line in current.stem_lines
            if _rich_line_text_value(line).strip()
        ]
        if not current_stem_texts or not all(
            _looks_like_digit_leading_chinese_stem(text) or len(text) <= 4
            for text in current_stem_texts
        ):
            index += 1
            continue

        previous.stem_lines = _clone_rich_lines([*previous.stem_lines, *current.stem_lines])
        previous.option_lines = _clone_rich_lines(current.option_lines)
        questions.pop(index)


def _build_question_from_segment_with_orphan_options(
    items: list[LineItem],
    start: int,
    end: int,
) -> tuple[ExamQuestion | None, list[RichLine]]:
    orphan_option_lines: list[RichLine] = []
    question = _build_question_from_segment(items, start, end)
    if question is None:
        return None, orphan_option_lines
    spans = _collect_option_spans(items, start, end)
    if spans:
        _opt_start, opt_end, _opts = spans[0]
        orphan_option_lines = _extract_option_like_rich_lines(items[opt_end:end])
    return question, orphan_option_lines


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

    placeholder_cluster = _collect_blank_option_placeholder_cluster(items, start, end)
    if placeholder_cluster is not None:
        opt_start, _opt_end, placeholder_opts = placeholder_cluster
        source_number = _extract_source_number_from_segment(segment)
        stem_seg = items[start:opt_start]
        stem_lines = _segment_to_stem_only(stem_seg)
        if stem_lines or placeholder_opts or source_number:
            return ExamQuestion(
                stem_lines=stem_lines,
                option_lines=_clone_rich_lines(placeholder_opts),
                source_number=source_number,
            )

    source_number = _extract_source_number_from_segment(segment)
    stem_lines = _segment_to_stem_only(segment)
    embedded_option_lines: list[RichLine] = []
    if stem_lines:
        stem_lines, embedded_option_lines = _split_embedded_option_lines_from_stem(stem_lines)
        if not embedded_option_lines:
            candidate_stem, single_embedded = _split_embedded_option_lines_from_stem(
                stem_lines,
                allow_single_option_lines=True,
            )
            merged_single = _merge_option_rich_lines(single_embedded)
            if merged_single:
                stem_lines = candidate_stem
                embedded_option_lines = merged_single
    if not stem_lines and source_number and not any(image_path for _text, image_path in segment):
        leading_images: list[str] = []
        index = start - 1
        blocked_by_prior_question = False
        while index >= 0:
            text, image_path = items[index]
            if image_path:
                leading_images.append(image_path)
                index -= 1
                continue
            raw = (text or "").strip()
            if not raw:
                index -= 1
                continue
            normalized_raw = _normalize_pdf_option_text(raw)
            blocked_by_prior_question = (
                _starts_new_question_line(normalized_raw)
                or _looks_like_question_stem_text(normalized_raw)
                or _looks_like_generic_visual_prompt_text(normalized_raw)
                or _has_terminal_question_signal(normalized_raw)
            )
            break
        if leading_images and not blocked_by_prior_question:
            stem_lines = [_rich_img(image_path) for image_path in reversed(leading_images)]
    if stem_lines or embedded_option_lines or source_number:
        if source_number and not stem_lines and len(segment) == 1 and _is_question_no_line((segment[0][0] or "").strip()):
            return ExamQuestion(stem_lines=[], option_lines=[], source_number=source_number)
        has_image_only_payload = any(image_path for _text, image_path in segment) and not any(
            (text or "").strip() and not _is_question_no_line((text or "").strip())
            for text, image_path in segment
            if not image_path
        )
        has_inline_option_markers = any(
            not image_path
            and bool(re.search(r"[:：?？]\s*A[\.．、:：)\uFF09]", _normalize_pdf_option_text((text or "").strip())))
            for text, image_path in segment
            if (text or "").strip()
        )
        has_visual_context = has_image_only_payload or any(_rich_line_has_image(line) for line in stem_lines)
        if (
            not _rich_lines_look_like_question_text(stem_lines)
            and not has_image_only_payload
            and not has_inline_option_markers
            and not embedded_option_lines
            and not (source_number and has_visual_context and stem_lines)
        ):
            return None
        return ExamQuestion(
            stem_lines=stem_lines,
            option_lines=embedded_option_lines,
            source_number=source_number,
        )
    return None


def _repair_local_question_number_anomalies(questions: list[ExamQuestion]) -> None:
    numeric_numbers: list[int | None] = []
    for question in questions:
        raw = str(question.source_number or "").strip()
        numeric_numbers.append(int(raw) if raw.isdigit() else None)

    index = 0
    while index < len(questions) - 1:
        prev_number = numeric_numbers[index]
        if prev_number is None:
            index += 1
            continue
        repaired = False
        for next_index in range(index + 1, len(questions)):
            next_number = numeric_numbers[next_index]
            if next_number is None or next_number <= prev_number:
                continue
            between_count = next_index - index - 1
            gap = next_number - prev_number - 1
            if (
                between_count > 0
                and between_count <= _LOCAL_NUMBER_REPAIR_MAX_SPAN
                and gap == between_count
            ):
                for offset, repair_index in enumerate(range(index + 1, next_index), start=1):
                    repaired_number = prev_number + offset
                    questions[repair_index].source_number = str(repaired_number)
                    numeric_numbers[repair_index] = repaired_number
                index = next_index
                repaired = True
                break
            if between_count == 0:
                index = next_index
                repaired = True
                break
        if not repaired:
            index += 1

    for index in range(1, len(questions) - 1):
        prev_number = numeric_numbers[index - 1]
        current_number = numeric_numbers[index]
        next_number = numeric_numbers[index + 1]
        if prev_number is None or current_number is None or next_number is None:
            continue
        expected = prev_number + 1
        if next_number == prev_number + 2 and current_number != expected:
            questions[index].source_number = str(expected)
            numeric_numbers[index] = expected

    for index in range(1, len(questions)):
        prev_number = numeric_numbers[index - 1]
        current_number = numeric_numbers[index]
        if prev_number is None or current_number is None:
            continue
        expected = prev_number + 1
        if current_number == expected or current_number >= expected:
            continue
        if prev_number < 10:
            continue
        current_text = str(current_number)
        expected_text = str(expected)
        if len(expected_text) <= len(current_text):
            continue
        if not expected_text.endswith(current_text):
            continue
        questions[index].source_number = expected_text
        numeric_numbers[index] = expected


def _parse_numeric_source_number(value: str) -> int | None:
    raw = (value or "").strip()
    return int(raw) if raw.isdigit() else None


def _merge_adjacent_split_questions(questions: list[ExamQuestion]) -> None:
    index = 0
    while index < len(questions) - 1:
        current = questions[index]
        nxt = questions[index + 1]
        if current.option_lines or not nxt.option_lines:
            index += 1
            continue
        if _question_is_image_only_payload(current):
            index += 1
            continue
        current_stem_text = _rich_lines_text(current.stem_lines)
        current_has_prompt_signal = _has_terminal_question_signal(current_stem_text)

        prev_number = (
            _parse_numeric_source_number(questions[index - 1].source_number)
            if index > 0
            else None
        )
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if prev_number is None:
            index += 1
            continue

        keep_number: int | None = None
        if current_number == prev_number + 1 and next_number != current_number + 1:
            if current_stem_text and current_has_prompt_signal:
                index += 1
                continue
            keep_number = current_number
        elif next_number == prev_number + 1 and current_number != next_number:
            if current_has_prompt_signal and not (
                current_number is not None and current_number + 20 < prev_number
            ):
                index += 1
                continue
            keep_number = next_number

        if keep_number is None:
            index += 1
            continue

        merged_stem_lines = list(current.stem_lines)
        merged_stem_lines.extend(nxt.stem_lines)
        questions[index] = ExamQuestion(
            stem_lines=merged_stem_lines,
            option_lines=list(nxt.option_lines),
            source_number=str(keep_number),
        )
        del questions[index + 1]


def _repair_shifted_trailing_image_payload_questions(questions: list[ExamQuestion]) -> None:
    index = 0
    while index < len(questions) - 2:
        current = questions[index]
        nxt = questions[index + 1]
        following = questions[index + 2]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        following_number = _parse_numeric_source_number(following.source_number)
        if (
            current_number is None
            or next_number != current_number + 1
            or following_number is None
            or following_number <= next_number + 1
            or not _question_is_image_only_payload(nxt)
        ):
            index += 1
            continue

        if not current.option_lines or not _option_cluster_has_substance(current.option_lines[:-1]):
            index += 1
            continue

        carried_image = _pop_trailing_image_only_option_line(current)
        if carried_image is None:
            index += 1
            continue

        shifted_payload = _clone_rich_lines(nxt.stem_lines)
        if not shifted_payload:
            current.option_lines = _clone_rich_lines([*current.option_lines, carried_image])
            index += 1
            continue

        nxt.stem_lines = [carried_image]
        insert_at = index + 2
        questions.insert(
            insert_at,
            ExamQuestion(
                stem_lines=shifted_payload,
                option_lines=[],
                source_number=str(next_number + 1),
            ),
        )
        insert_at += 1
        for missing_number in range(next_number + 2, following_number):
            questions.insert(
                insert_at,
                ExamQuestion(
                    stem_lines=[],
                    option_lines=[],
                    source_number=str(missing_number),
                ),
            )
            insert_at += 1
        index = insert_at


def _looks_like_objective_stem_header_line(line: RichLine) -> bool:
    text = _normalize_pdf_option_text(_rich_line_text_value(line))
    if not text:
        return False
    if _rich_option_letter(line) or _is_question_no_line(text):
        return False
    if _match_leading_question_with_stem(text):
        return True
    if any(
        marker in text
        for marker in (
            "根据上述",
            "根据以下",
            "下列",
            "以下",
            "从所给的",
            "把下面",
            "将下列",
            "如果",
            "定义",
            "图形",
            "图中",
        )
    ):
        return True
    if re.match(r"^[^\s]{2,24}[：:].+$", text):
        return True
    return False


def _repair_spilled_leading_option_lines(questions: list[ExamQuestion]) -> None:
    for index in range(len(questions) - 1):
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if current_number is None or next_number != current_number + 1:
            continue
        if current.option_lines or not nxt.option_lines or not nxt.stem_lines:
            continue

        current_letters = [letter for letter in (_rich_option_letter(line) for line in current.stem_lines) if letter]
        if current_letters != ["A", "B", "C"]:
            continue

        d_index: int | None = None
        for probe_index, line in enumerate(nxt.stem_lines[:4]):
            if _rich_option_letter(line) == "D":
                d_index = probe_index
                break
        if d_index is None:
            continue

        split_index: int | None = None
        for probe_index in range(d_index + 1, len(nxt.stem_lines)):
            if _looks_like_objective_stem_header_line(nxt.stem_lines[probe_index]):
                split_index = probe_index
                break
        if split_index is None:
            continue

        current.stem_lines = _clone_rich_lines([*current.stem_lines, *nxt.stem_lines[:split_index]])
        nxt.stem_lines = _clone_rich_lines(nxt.stem_lines[split_index:])


def _repair_malformed_question_tail_lines(questions: list[ExamQuestion]) -> None:
    for index in range(len(questions) - 1):
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if current_number is None or next_number not in {None, current_number + 1}:
            continue
        if not current.option_lines or nxt.stem_lines or not nxt.option_lines:
            continue

        tail_line = current.option_lines[-1]
        if _rich_line_has_image(tail_line):
            continue
        tail_text = _normalize_pdf_option_text(_rich_line_text_value(tail_line))
        matched = _MALFORMED_NEXT_QUESTION_TAIL.match(tail_text)
        if not matched:
            continue

        stem = (matched.group("stem") or "").strip()
        if not stem or not _looks_like_question_stem_text(stem):
            continue

        current.option_lines = _clone_rich_lines(current.option_lines[:-1])
        nxt.stem_lines = [_rich_text(stem), *_clone_rich_lines(nxt.stem_lines)]
        if next_number is None:
            nxt.source_number = str(current_number + 1)


def _rich_lines_to_stem_only(lines: list[RichLine]) -> list[RichLine]:
    segment: list[LineItem] = []
    prefix_stripped = False
    for line in lines:
        for text, image_path in line.parts:
            if image_path:
                segment.append(("", image_path))
            elif (text or "").strip():
                stripped = (text or "").strip()
                if not prefix_stripped:
                    prefix_match = re.match(
                        r"^\s*(?P<number>\d{1,5})\s*[\.．、)\uFF09]\s*(?P<stem>.*)$",
                        stripped,
                    )
                    if prefix_match:
                        stripped = (prefix_match.group("stem") or "").strip()
                        prefix_stripped = True
                    elif _is_question_no_line(stripped):
                        prefix_stripped = True
                        continue
                segment.append((stripped, None))
    return _segment_to_stem_only(segment)


def _split_embedded_next_stem_from_option_line(
    line: RichLine,
    *,
    expected_number: int,
) -> tuple[RichLine, list[RichLine]] | None:
    if _rich_line_has_image(line):
        return None
    text = _normalize_pdf_option_text(_rich_line_text_value(line))
    matched = _match_single_option_line(text)
    if not matched:
        return None
    letter = matched.group(1).upper()
    body = (matched.group(2) or "").strip()
    if letter != "D" or not body:
        return None

    stem_match = re.search(
        rf"(?<!\d){expected_number}\s*[\.．、)\uFF09]?\s*(?P<stem>.+)$",
        body,
    )
    if not stem_match:
        return None

    stem = (stem_match.group("stem") or "").strip()
    option_text = body[: stem_match.start()].strip()
    if not stem or not option_text or not (
        _looks_like_question_stem_text(stem) or len(stem) >= 12
    ):
        return None
    return _rich_text(f"{letter}．{option_text}"), [_rich_text(stem)]


def _repair_missing_stem_questions_from_previous_option_tail(questions: list[ExamQuestion]) -> None:
    for index in range(len(questions) - 1):
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        expected_number = current_number + 1 if current_number is not None else None
        if (
            current_number is None
            or nxt.stem_lines
            or not nxt.option_lines
            or len(current.option_lines) < 4
            or (
                next_number is not None
                and expected_number is not None
                and next_number != expected_number
            )
        ):
            continue

        option_letters = [_rich_option_letter(line) for line in current.option_lines[:4]]
        if option_letters != ["A", "B", "C", "D"]:
            continue

        clean_option_lines = _clone_rich_lines(current.option_lines[:4])
        tail_lines = _clone_rich_lines(current.option_lines[4:])
        embedded_split = _split_embedded_next_stem_from_option_line(
            clean_option_lines[-1],
            expected_number=expected_number,
        )
        if embedded_split is not None:
            clean_option_lines[-1] = embedded_split[0]
            tail_lines = embedded_split[1] + tail_lines

        if not tail_lines:
            continue

        stem_lines = _rich_lines_to_stem_only(tail_lines)
        if not stem_lines:
            continue

        current.option_lines = clean_option_lines
        nxt.stem_lines = stem_lines
        if not (nxt.source_number or "").strip():
            nxt.source_number = str(expected_number)


def _looks_like_generic_visual_prompt_text(text: str) -> bool:
    normalized = _normalize_pdf_option_text((text or "").strip())
    if not normalized:
        return False
    if "从所给的四个选项中" in normalized and "填入问号处" in normalized:
        return True
    if normalized.startswith("把下面的六个图形分为两类"):
        return True
    return False


def _question_is_image_only_payload(question: ExamQuestion) -> bool:
    if question.option_lines or not question.stem_lines:
        return False
    return all(
        _rich_line_has_image(line) and not _rich_line_text_value(line)
        for line in question.stem_lines
    )


def _pop_trailing_image_only_option_line(question: ExamQuestion) -> RichLine | None:
    if not question.option_lines:
        return None
    tail = question.option_lines[-1]
    if not _rich_line_has_image(tail) or _rich_line_text_value(tail):
        return None
    question.option_lines = _clone_rich_lines(question.option_lines[:-1])
    return RichLine(parts=list(tail.parts))


def _repair_shifted_visual_prompt_questions(questions: list[ExamQuestion]) -> None:
    index = 1
    while index < len(questions):
        previous = questions[index - 1]
        current = questions[index]
        previous_number = _parse_numeric_source_number(previous.source_number)
        current_number = _parse_numeric_source_number(current.source_number)
        if (
            previous_number is None
            or current_number is None
            or current_number != previous_number + 2
            or len(current.stem_lines) < 2
        ):
            index += 1
            continue

        first_text = _rich_line_text_value(current.stem_lines[0])
        if not _looks_like_generic_visual_prompt_text(first_text):
            index += 1
            continue
        if not any(
            _rich_line_text_value(line) or _rich_line_has_image(line)
            for line in current.stem_lines[1:]
        ) and not current.option_lines:
            index += 1
            continue

        prompt_line = RichLine(parts=list(current.stem_lines[0].parts))
        current.stem_lines = _clone_rich_lines(current.stem_lines[1:])
        questions.insert(
            index,
            ExamQuestion(
                stem_lines=[prompt_line],
                option_lines=[],
                source_number=str(previous_number + 1),
            ),
        )
        index += 2


def _repair_prompt_only_placeholder_questions(questions: list[ExamQuestion]) -> None:
    for index in range(1, len(questions) - 1):
        previous = questions[index - 1]
        current = questions[index]
        nxt = questions[index + 1]
        previous_number = _parse_numeric_source_number(previous.source_number)
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if (
            previous_number is None
            or current_number != previous_number + 1
            or next_number != current_number + 1
            or current.stem_lines
            or current.option_lines
            or not nxt.stem_lines
        ):
            continue

        first_text = _rich_line_text_value(nxt.stem_lines[0])
        if not _looks_like_generic_visual_prompt_text(first_text):
            continue

        current.stem_lines = [RichLine(parts=list(nxt.stem_lines[0].parts))]
        nxt.stem_lines = _clone_rich_lines(nxt.stem_lines[1:])


def _repair_shifted_image_only_question_sequences(questions: list[ExamQuestion]) -> None:
    index = 0
    while index < len(questions) - 1:
        current = questions[index]
        nxt = questions[index + 1]
        current_number = _parse_numeric_source_number(current.source_number)
        next_number = _parse_numeric_source_number(nxt.source_number)
        if (
            current_number is None
            or next_number is None
            or next_number <= current_number + 1
            or not _question_is_image_only_payload(current)
        ):
            index += 1
            continue
        if nxt.stem_lines and _looks_like_generic_visual_prompt_text(_rich_line_text_value(nxt.stem_lines[0])):
            index += 1
            continue

        donor_line: RichLine | None = None
        if index > 0:
            donor_line = _pop_trailing_image_only_option_line(questions[index - 1])

        shifted_payload = _clone_rich_lines(current.stem_lines)
        if donor_line is not None:
            current.stem_lines = [donor_line]
        else:
            current.stem_lines = []
        current.option_lines = []

        insert_at = index + 1
        questions.insert(
            insert_at,
            ExamQuestion(
                stem_lines=shifted_payload,
                option_lines=[],
                source_number=str(current_number + 1),
            ),
        )
        insert_at += 1
        for missing_number in range(current_number + 2, next_number):
            questions.insert(
                insert_at,
                ExamQuestion(
                    stem_lines=[],
                    option_lines=[],
                    source_number=str(missing_number),
                ),
            )
            insert_at += 1
        index = insert_at


def parse_material_body(
    items: list[LineItem],
    body_start: int,
    body_end: int,
    header: str,
) -> MaterialUnit | None:
    """解析 [body_start, body_end) 正文区间（无「材料X」标题行）。"""
    markers = _collect_question_markers(items, body_start, body_end)
    if markers:
        intro = _segment_to_rich_lines(items[body_start:markers[0]], strip_question_prefix=False)
        questions: list[ExamQuestion] = []
        orphan_option_lines: list[RichLine] = []
        for index, start in enumerate(markers):
            segment_end = markers[index + 1] if index + 1 < len(markers) else body_end
            question, segment_orphans = _build_question_from_segment_with_orphan_options(
                items,
                start,
                segment_end,
            )
            if question is not None:
                questions.append(question)
            orphan_option_lines.extend(segment_orphans)
        if questions:
            _redistribute_orphan_option_lines(questions, orphan_option_lines)
            _repair_spilled_leading_option_lines(questions)
            _repair_page_footer_split_questions(questions)
            _repair_malformed_question_tail_lines(questions)
            _repair_local_question_number_anomalies(questions)
            return MaterialUnit(
                header=_normalize_material_header_question_range(header, questions),
                intro_lines=intro,
                questions=questions,
            )

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

    _repair_local_question_number_anomalies(questions)
    return MaterialUnit(
        header=_normalize_material_header_question_range(header, questions),
        intro_lines=intro,
        questions=questions,
    )


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


def _clone_question(
    question: ExamQuestion,
    *,
    stem_lines: list[RichLine] | None = None,
    option_lines: list[RichLine] | None = None,
) -> ExamQuestion:
    return ExamQuestion(
        stem_lines=list(question.stem_lines if stem_lines is None else stem_lines),
        option_lines=list(question.option_lines if option_lines is None else option_lines),
        source_number=question.source_number,
    )


def _strict_five_question_material_header(start: int, end: int) -> str:
    return f"材料（回答{start}-{end}题）"


def _partition_material_intro_lines(intro_lines: list[RichLine], group_count: int) -> list[list[RichLine]]:
    if group_count <= 1:
        return [list(intro_lines)]
    meaningful = [line for line in intro_lines if _line_has_text(line) or _rich_line_has_image(line)]
    if not meaningful:
        return [list(intro_lines)] + [[] for _ in range(group_count - 1)]

    segments: list[list[RichLine]] = []
    start = 0
    for index in range(len(intro_lines) - 1):
        line = intro_lines[index]
        next_line = intro_lines[index + 1]
        current_is_boundary = _rich_line_has_image(line) or not _line_has_text(line)
        next_has_content = _line_has_text(next_line) or _rich_line_has_image(next_line)
        if not (current_is_boundary and next_has_content):
            continue
        segment = list(intro_lines[start : index + 1])
        if any(_line_has_text(item) or _rich_line_has_image(item) for item in segment):
            segments.append(segment)
            start = index + 1
    tail = list(intro_lines[start:])
    if any(_line_has_text(item) or _rich_line_has_image(item) for item in tail):
        segments.append(tail)

    if len(segments) == group_count:
        return segments
    return [list(intro_lines)] + [[] for _ in range(group_count - 1)]


def _repartition_strict_five_question_materials(materials: list[MaterialUnit]) -> list[MaterialUnit]:
    """
    单科资料分析题本通常严格遵循「一则材料 + 5 道小题」。
    当原始材料头抽取不稳定时，按全局连续题号每 5 题重组材料。
    """
    if not materials:
        return materials

    flat_questions: list[tuple[int, MaterialUnit, int, ExamQuestion]] = []
    for unit in materials:
        for question_index, question in enumerate(unit.questions):
            numeric = _parse_numeric_source_number(question.source_number)
            if numeric is None:
                return materials
            flat_questions.append((numeric, unit, question_index, question))

    if len(flat_questions) < 10 or len(flat_questions) % 5 != 0:
        return materials

    flat_questions.sort(key=lambda entry: entry[0])
    numeric_numbers = [entry[0] for entry in flat_questions]
    start_number = numeric_numbers[0]
    if (start_number - 1) % 5 != 0:
        return materials
    expected_numbers = list(range(start_number, start_number + len(flat_questions)))
    if numeric_numbers != expected_numbers:
        return materials
    if all(len(unit.questions) == 5 for unit in materials):
        return materials

    intro_partitions: dict[int, list[list[RichLine]]] = {}
    for unit in materials:
        chunk_count = max(1, len(unit.questions) // 5)
        intro_partitions[id(unit)] = _partition_material_intro_lines(unit.intro_lines, chunk_count)

    rebuilt: list[MaterialUnit] = []
    pending_intro_spill: list[RichLine] = []
    for chunk_start in range(0, len(flat_questions), 5):
        chunk = flat_questions[chunk_start : chunk_start + 5]
        first_number, first_unit, first_question_index, first_question = chunk[0]
        last_number = chunk[-1][0]
        intro_partition_index = first_question_index // 5
        unit_intro_partitions = intro_partitions.get(id(first_unit), [list(first_unit.intro_lines)])
        intro_seed = (
            list(unit_intro_partitions[intro_partition_index])
            if intro_partition_index < len(unit_intro_partitions)
            else []
        )

        if not rebuilt:
            intro_lines = list(pending_intro_spill)
            intro_lines.extend(intro_seed)
            first_stem_lines = _clone_rich_lines(first_question.stem_lines)
        else:
            prev_last_question = rebuilt[-1].questions[-1]
            clean_options, spill_intro = _split_material_intro_from_option_lines(prev_last_question.option_lines)
            prev_last_question.option_lines = clean_options
            intro_from_stem, first_stem_lines = _split_rich_intro_stem(first_question.stem_lines)
            intro_lines = list(pending_intro_spill)
            intro_lines.extend(spill_intro)
            intro_lines.extend(intro_seed)
            intro_lines.extend(intro_from_stem)

        donor_intro_patches: list[tuple[int, list[RichLine], list[list[RichLine]]]] = []
        seen_units = {id(first_unit)}
        for chunk_index, (_number, unit, question_index, _question) in enumerate(chunk):
            unit_id = id(unit)
            if unit_id in seen_units:
                continue
            seen_units.add(unit_id)
            if question_index != 0:
                continue
            donor_intro, donor_option_clusters = _extract_option_clusters_from_rich_lines(unit.intro_lines)
            donor_intro_patches.append((chunk_index, donor_intro, donor_option_clusters))

        question_intro_spill: list[RichLine] = []
        trailing_intro_spill: list[RichLine] = []
        rebuilt_questions: list[ExamQuestion] = []
        for index, (_number, _unit, _question_index, question) in enumerate(chunk):
            clean_options, spill_intro = _split_material_intro_from_option_lines(question.option_lines)
            if spill_intro:
                if index == len(chunk) - 1:
                    trailing_intro_spill.extend(spill_intro)
                else:
                    question_intro_spill.extend(spill_intro)
            rebuilt_questions.append(
                _clone_question(
                    question,
                    stem_lines=first_stem_lines if index == 0 else None,
                    option_lines=clean_options,
                )
            )

        for split_index, donor_intro, donor_option_clusters in donor_intro_patches:
            intro_lines.extend(donor_intro)
            missing_targets = [
                index
                for index in range(split_index)
                if not rebuilt_questions[index].option_lines
            ]
            assigned = 0
            for target_index, option_cluster in zip(missing_targets, donor_option_clusters):
                rebuilt_questions[target_index].option_lines = _clone_rich_lines(option_cluster)
                assigned += 1
            for option_cluster in donor_option_clusters[assigned:]:
                intro_lines.extend(_clone_rich_lines(option_cluster))

        if question_intro_spill:
            intro_lines.extend(question_intro_spill)
        pending_intro_spill = trailing_intro_spill

        rebuilt.append(
            MaterialUnit(
                header=_strict_five_question_material_header(first_number, last_number),
                intro_lines=intro_lines,
                questions=rebuilt_questions,
            )
        )
    return rebuilt


def _maybe_repartition_single_subject_data_sections(
    exam: ParsedExam,
    *,
    source_name: str | None,
) -> ParsedExam:
    filename_profile = infer_pdf_filename_profile(source_name or "")
    if filename_profile.form != "single_subject_book" or filename_profile.subject_hint != "data":
        return exam
    for section in exam.data_sections:
        section.materials = _repartition_strict_five_question_materials(section.materials)
        for _ in range(3):
            section.materials = _realign_adjacent_data_material_intros(section.materials)
    return exam


_MATERIAL_CONTINUATION_ENDS = re.compile(r"[，,、；;：:—–\-]$")
_MATERIAL_TABLE_HEADER = re.compile(
    r"^\s*表\s*[一二三四五六七八九十\d1-9]|^\s*[（(]\s*单位\s*[：:)]|"
    r"^\s*年\s*份\s*[|│]|^\s*项\s*目\s*[|│]|^\s*指\s*标\s*[|│]"
)
_TABLE_CONTINUATION_CHARS = frozenset("│|—─┼┤├┬┴┌┐└┘╔╗╚╝═║")
_MATERIAL_QUESTION_RANGE = re.compile(
    r"(回答\s*)(\d{1,3})(\s*[-—–~～至到]+\s*)(\d{1,3})(\s*题)"
)
_MATERIAL_HEADER_ORDINAL = re.compile(
    r"^\s*材料\s*([一二三四五六七八九十百千万\d〇零两]+)"
)
_MATERIAL_INTRO_MATCH_CHARS = re.compile(r"[^A-Za-z0-9\u4e00-\u9fff]+")
_GENERIC_MATERIAL_PROMPT = re.compile(
    r"^\s*(?:根据(?:以下|上述|所给)?(?:资料|材料)|结合(?:以下|上述|所给)?(?:资料|材料))\s*[，,:：]?\s*回答?(?:下列)?问题[。．.]?\s*$"
)


def _material_intro_match_ngrams(lines: list[RichLine]) -> set[str]:
    text = _nfkc("".join(_rich_line_text_value(line) for line in lines if _rich_line_text_value(line)))
    normalized = _MATERIAL_INTRO_MATCH_CHARS.sub("", text)
    if len(normalized) < 2:
        return {normalized} if normalized else set()
    return {normalized[index : index + 2] for index in range(len(normalized) - 1)}


def _question_stem_match_ngrams(question: ExamQuestion) -> set[str]:
    return _material_intro_match_ngrams(question.stem_lines)


def _material_question_match_ngrams(material: MaterialUnit) -> set[str]:
    stem_lines: list[RichLine] = []
    for question in material.questions:
        stem_lines.extend(question.stem_lines)
    return _material_intro_match_ngrams(stem_lines)


def _intro_is_generic_prompt_or_images(intro_lines: list[RichLine]) -> bool:
    if not intro_lines:
        return False
    has_image = False
    for line in intro_lines:
        text = _nfkc(_rich_line_text_value(line))
        if _rich_line_has_image(line):
            has_image = True
        if not text:
            continue
        if not _GENERIC_MATERIAL_PROMPT.match(text.strip()):
            return False
    return has_image


def _intro_is_generic_prompt_text_only(intro_lines: list[RichLine]) -> bool:
    if not intro_lines:
        return False
    saw_text = False
    for line in intro_lines:
        text = _nfkc(_rich_line_text_value(line)).strip()
        if _rich_line_has_image(line):
            return False
        if not text:
            continue
        if not _GENERIC_MATERIAL_PROMPT.match(text):
            return False
        saw_text = True
    return saw_text


def _intro_is_generic_prompt_with_images(intro_lines: list[RichLine]) -> bool:
    if not intro_lines:
        return False
    saw_text = False
    saw_image = False
    for line in intro_lines:
        text = _nfkc(_rich_line_text_value(line)).strip()
        if _rich_line_has_image(line):
            saw_image = True
        if not text:
            continue
        if not _GENERIC_MATERIAL_PROMPT.match(text):
            return False
        saw_text = True
    return saw_text and saw_image


def _realign_adjacent_data_material_intros(materials: list[MaterialUnit]) -> list[MaterialUnit]:
    if len(materials) < 2:
        return materials
    for index in range(len(materials) - 1):
        previous = materials[index]
        current = materials[index + 1]
        if len(previous.questions) != 5 or len(current.questions) != 5:
            continue
        if not previous.intro_lines and current.intro_lines and _intro_is_generic_prompt_text_only(current.intro_lines):
            previous.intro_lines = _clone_rich_lines(current.intro_lines)
            current.intro_lines = []
            continue
        if not previous.intro_lines or current.intro_lines:
            continue
        previous_stem_ngrams = _material_question_match_ngrams(previous)
        current_stem_ngrams = _material_question_match_ngrams(current)
        if not current_stem_ngrams:
            continue

        best_split: int | None = None
        best_margin = 0
        for split_index in range(1, len(previous.intro_lines)):
            suffix = previous.intro_lines[split_index:]
            if not _looks_like_material_intro_lines(suffix):
                continue
            suffix_ngrams = _material_intro_match_ngrams(suffix)
            if len(suffix_ngrams) < 4:
                continue
            previous_score = len(suffix_ngrams & previous_stem_ngrams)
            current_score = len(suffix_ngrams & current_stem_ngrams)
            if current_score < max(previous_score + 3, 4):
                continue
            margin = current_score - previous_score
            if margin > best_margin:
                best_margin = margin
                best_split = split_index

        if best_split is not None:
            moved = _clone_rich_lines(previous.intro_lines[best_split:])
            previous.intro_lines = _clone_rich_lines(previous.intro_lines[:best_split])
            current.intro_lines = moved + _clone_rich_lines(current.intro_lines)
            continue

        full_intro_ngrams = _material_intro_match_ngrams(previous.intro_lines)
        if len(full_intro_ngrams) < 4:
            continue
        previous_score = len(full_intro_ngrams & previous_stem_ngrams)
        current_score = len(full_intro_ngrams & current_stem_ngrams)
        if current_score >= max(previous_score + 4, 5):
            current.intro_lines = _clone_rich_lines(previous.intro_lines) + _clone_rich_lines(current.intro_lines)
            previous.intro_lines = []
    return materials


def _material_intro_looks_truncated(intro_lines: list[RichLine]) -> bool:
    """判断材料正文是否在跨页时被截断：句末是逗号/分号等非终结标点。"""
    if not intro_lines:
        return False
    text_lines = [_rich_line_text_value(line) for line in intro_lines if _line_has_text(line)]
    if not text_lines:
        return False
    last_text = text_lines[-1].rstrip()
    if not last_text:
        return False
    if _MATERIAL_CONTINUATION_ENDS.search(last_text):
        return True
    if last_text[-1] not in "。！？!?.）)」》】":
        if len(last_text) >= 8 and not _starts_new_question_line(last_text):
            return True
    return False


def _looks_like_material_continuation_text(text: str) -> bool:
    """判断文本行是否像是材料正文的续行（非题干、非选项、非标题）。"""
    s = _nfkc((text or "").strip())
    if not s:
        return False
    if _starts_new_question_line(s):
        return False
    if _has_option_markers_pdf(s):
        return False
    if _detect_subject_section_kind(s):
        return False
    if _is_other_section_title(s):
        return False
    if material_header_line(s) or generic_material_header_line(s):
        return False
    if len(s) >= 8:
        return True
    return bool(_TABLE_CONTINUATION_CHARS & set(s)) or bool(_MATERIAL_TABLE_HEADER.match(s))


def _looks_like_table_continuation(intro_lines: list[RichLine], next_lines: list[RichLine]) -> bool:
    """判断前后两组 intro 是否像被跨页拆断的同一张表格。"""
    if not intro_lines or not next_lines:
        return False
    last_text = _rich_line_text_value(intro_lines[-1]) if _line_has_text(intro_lines[-1]) else ""
    first_text = _rich_line_text_value(next_lines[0]) if _line_has_text(next_lines[0]) else ""
    if _TABLE_CONTINUATION_CHARS & set(last_text) and _TABLE_CONTINUATION_CHARS & set(first_text):
        return True
    return False


def _material_header_question_range(header: str) -> tuple[int, int] | None:
    match = _MATERIAL_QUESTION_RANGE.search(_nfkc(header or ""))
    if not match:
        return None
    start = int(match.group(2))
    end = int(match.group(4))
    if end < start:
        start, end = end, start
    return start, end


def _actual_question_range(questions: list[ExamQuestion]) -> tuple[int, int] | None:
    numeric_numbers: list[int] = []
    for question in questions:
        text = str(getattr(question, "source_number", "") or "").strip()
        if not text:
            continue
        try:
            numeric_numbers.append(int(text))
        except ValueError:
            continue
    if not numeric_numbers:
        return None
    return min(numeric_numbers), max(numeric_numbers)


def _normalize_material_header_question_range(header: str, questions: list[ExamQuestion]) -> str:
    actual_range = _actual_question_range(questions)
    if actual_range is None:
        return header
    parsed_range = _material_header_question_range(header)
    if parsed_range is None or parsed_range == actual_range:
        return header
    match = _MATERIAL_QUESTION_RANGE.search(header or "")
    if not match:
        return header
    corrected = (
        f"{match.group(1)}{actual_range[0]}{match.group(3)}{actual_range[1]}{match.group(5)}"
    )
    return f"{header[:match.start()]}{corrected}{header[match.end():]}"


def _material_header_ordinal(header: str) -> str | None:
    normalized = _nfkc((header or "").strip())
    if not normalized:
        return None
    normalized = normalized.strip("【】[]［］")
    match = _MATERIAL_HEADER_ORDINAL.match(normalized)
    if not match:
        return None
    return match.group(1)


def _headers_indicate_distinct_material(prev: MaterialUnit, current: MaterialUnit) -> bool:
    prev_ordinal = _material_header_ordinal(prev.header)
    current_ordinal = _material_header_ordinal(current.header)
    if (
        prev_ordinal
        and current_ordinal
        and prev_ordinal != current_ordinal
        and len(prev.questions) >= 5
        and len(current.questions) >= 5
    ):
        return True
    prev_range = _material_header_question_range(prev.header)
    current_range = _material_header_question_range(current.header)
    if prev_range and current_range and prev_range != current_range:
        return True
    return False


def _merge_cross_page_material_units(materials: list[MaterialUnit]) -> list[MaterialUnit]:
    """合并因跨页截断而被拆成相邻两个 MaterialUnit 的材料组。"""
    if len(materials) <= 1:
        return materials
    merged: list[MaterialUnit] = [materials[0]]
    for current in materials[1:]:
        prev = merged[-1]
        if _headers_indicate_distinct_material(prev, current):
            merged.append(current)
            continue
        should_merge = False
        if _material_intro_looks_truncated(prev.intro_lines) and current.intro_lines:
            first_text = _rich_line_text_value(current.intro_lines[0]) if current.intro_lines else ""
            if _looks_like_material_continuation_text(first_text):
                should_merge = True
        if not should_merge and _looks_like_table_continuation(prev.intro_lines, current.intro_lines):
            should_merge = True
        if not should_merge and not current.intro_lines and prev.questions and current.questions:
            prev_last_q = prev.questions[-1]
            prev_last_opts = [_rich_line_text_value(line) for line in prev_last_q.option_lines if _line_has_text(line)]
            if len(prev_last_opts) < 2:
                should_merge = True

        if should_merge:
            prev.intro_lines = list(prev.intro_lines) + list(current.intro_lines)
            prev.questions.extend(current.questions)
        else:
            merged.append(current)
    return merged


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
            if (
                segment_end == cursor + 1
                and items[cursor][1] is None
                and _is_question_no_line(_normalize_pdf_option_text((items[cursor][0] or "").strip()))
            ):
                if not _question_number_has_leading_visual_payload(items, cursor, a):
                    cursor = segment_end
                    continue
            segment_spans = _collect_option_spans(items, cursor, segment_end)
            if len(segment_spans) > 1:
                questions.extend(_build_questions_from_option_spans(items, cursor, segment_spans))
            else:
                question = _build_question_from_segment(items, cursor, segment_end)
                if question is not None:
                    questions.append(question)
            cursor = segment_end
        if questions:
            _repair_local_question_number_anomalies(questions)
            _merge_adjacent_split_questions(questions)
            _repair_shifted_trailing_image_payload_questions(questions)
            _repair_spilled_leading_option_lines(questions)
            _repair_page_footer_split_questions(questions)
            _repair_malformed_question_tail_lines(questions)
            _repair_missing_stem_questions_from_previous_option_tail(questions)
            _repair_shifted_visual_prompt_questions(questions)
            _repair_prompt_only_placeholder_questions(questions)
            _repair_shifted_image_only_question_sequences(questions)
            _repair_local_question_number_anomalies(questions)
            return questions

    spans = _collect_option_spans(items, a, b)
    if not spans:
        return []
    questions = _build_questions_from_option_spans(items, a, spans)
    _repair_local_question_number_anomalies(questions)
    _merge_adjacent_split_questions(questions)
    _repair_shifted_trailing_image_payload_questions(questions)
    _repair_spilled_leading_option_lines(questions)
    _repair_page_footer_split_questions(questions)
    _repair_malformed_question_tail_lines(questions)
    _repair_missing_stem_questions_from_previous_option_tail(questions)
    _repair_shifted_visual_prompt_questions(questions)
    _repair_prompt_only_placeholder_questions(questions)
    _repair_shifted_image_only_question_sequences(questions)
    _repair_local_question_number_anomalies(questions)
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
    if numeric_numbers != sorted(numeric_numbers):
        return False
    min_number = min(numeric_numbers)
    if min_number > 20:
        return False
    resolved = _resolve_set_paper_objective_kinds(questions)
    if not resolved:
        return False
    distinct_kinds = {kind for kind in resolved if kind != "unknown"}
    return len(distinct_kinds) >= 2


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
                split_units = _split_into_material_units(unit)
                for split_unit in _merge_cross_page_material_units(split_units):
                    sec.materials.append(split_unit)
        else:
            raw_units: list[MaterialUnit] = []
            for m_start, m_next, header in material_blocks:
                unit = parse_material_block(items, m_start, m_next, header=header)
                if unit:
                    raw_units.append(unit)
            for unit in _merge_cross_page_material_units(raw_units):
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
            split_units = _split_into_material_units(unit)
            for split_unit in _merge_cross_page_material_units(split_units):
                sec.materials.append(split_unit)
    else:
        raw_units: list[MaterialUnit] = []
        for m_start, m_next, header in material_blocks:
            unit = parse_material_block(items, m_start, m_next, header=header)
            if unit:
                raw_units.append(unit)
        for unit in _merge_cross_page_material_units(raw_units):
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
    filename_profile = infer_pdf_filename_profile(source_name or "")

    if document_subject_hint and document_subject_hint != "unknown":
        _parse_whole_document_as_subject(items, exam, document_subject_hint)
        return _maybe_repartition_single_subject_data_sections(exam, source_name=source_name)
    if (
        filename_profile.form == "single_subject_book"
        and filename_profile.subject_hint in selected_subjects
        and filename_profile.subject_hint is not None
        and filename_profile.subject_hint in {"politics", "common_sense"}
    ):
        _parse_whole_document_as_subject(
            items,
            exam,
            filename_profile.subject_hint,
            title=default_subject_title(filename_profile.subject_hint),
        )
        return _maybe_repartition_single_subject_data_sections(exam, source_name=source_name)
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

    has_subject_titles = any(kind in ALL_SUBJECT_KINDS for kind, *_rest in title_entries)
    if not title_entries or not has_subject_titles:
        _parse_without_titles(items, exam, selected_subjects, source_name=source_name)
        return _maybe_repartition_single_subject_data_sections(exam, source_name=source_name)

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

    return _maybe_repartition_single_subject_data_sections(exam, source_name=source_name)
