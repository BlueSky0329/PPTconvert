from __future__ import annotations

import logging
from typing import Iterable

from docx.oxml.ns import qn

LOGGER = logging.getLogger(__name__)


def _attrib_val(elem, name: str = "val") -> str:
    for key, value in elem.attrib.items():
        if key == name or key.endswith("}" + name):
            return value or ""
    return ""


def local_tag(elem) -> str:
    if elem.tag.startswith("{"):
        return elem.tag.split("}", 1)[-1]
    return elem.tag


def _collect_children_text(elem) -> str:
    return "".join(_omml_to_unicode(child) for child in elem)


def _decode_symbol(token: str) -> str:
    if not token:
        return "∑"
    if len(token) == 1:
        return token

    normalized = token.strip()
    if normalized.lower().startswith("0x"):
        normalized = normalized[2:]
    if normalized.startswith("\\u") or normalized.startswith("\\U"):
        normalized = normalized[2:]
    try:
        return chr(int(normalized, 16))
    except ValueError:
        return token


def _nary_symbol(ch: str, sub: str, sup: str, inner: str) -> str:
    symbol = _decode_symbol(ch)
    parts = [symbol]
    if sub:
        parts.append(f"_({sub})")
    if sup:
        parts.append(f"^({sup})")
    parts.append(inner)
    return "".join(parts)


def _omml_to_unicode_impl(elem) -> str:
    tag = local_tag(elem)

    if tag == "t":
        return elem.text or ""

    if tag == "r":
        return _collect_children_text(elem)

    if tag == "rad":
        degree_text = ""
        inner = ""
        for child in elem:
            child_tag = local_tag(child)
            if child_tag == "deg":
                degree_text = _collect_children_text(child)
            elif child_tag == "e":
                inner = _omml_to_unicode(child)
        if degree_text.strip():
            return f"root[{degree_text.strip()}]({inner})"
        return f"sqrt({inner})"

    if tag == "f":
        num = den = ""
        for child in elem:
            child_tag = local_tag(child)
            if child_tag == "num":
                num = _collect_children_text(child)
            elif child_tag == "den":
                den = _collect_children_text(child)
        return f"{num}/{den}"

    if tag == "sSup":
        base = sup = ""
        for child in elem:
            child_tag = local_tag(child)
            if child_tag == "e":
                base = _omml_to_unicode(child)
            elif child_tag == "sup":
                sup = _omml_to_unicode(child)
        return f"{base}^{sup}" if sup else (base or _collect_children_text(elem))

    if tag == "sSub":
        base = sub = ""
        for child in elem:
            child_tag = local_tag(child)
            if child_tag == "e":
                base = _omml_to_unicode(child)
            elif child_tag == "sub":
                sub = _omml_to_unicode(child)
        return f"{base}_{sub}" if sub else (base or _collect_children_text(elem))

    if tag == "sPre":
        return _collect_children_text(elem)

    if tag == "nary":
        ch = ""
        sub = sup = ""
        inner = ""
        for child in elem:
            child_tag = local_tag(child)
            if child_tag == "chr":
                ch = _attrib_val(child, "val") or (child.text or "")
            elif child_tag == "sub":
                sub = _collect_children_text(child)
            elif child_tag == "sup":
                sup = _collect_children_text(child)
            elif child_tag == "e":
                inner = _omml_to_unicode(child)
        if inner or ch or sub or sup:
            return _nary_symbol(ch, sub, sup, inner)
        return _collect_children_text(elem)

    if tag == "func":
        func_name = ""
        inner = ""
        for child in elem:
            child_tag = local_tag(child)
            if child_tag == "fName":
                func_name = _collect_children_text(child)
            elif child_tag == "e":
                inner = _omml_to_unicode(child)
        return f"{func_name}({inner})" if func_name else (inner or _collect_children_text(elem))

    if tag in ("d", "e", "oMath", "oMathPara"):
        return _collect_children_text(elem)

    return _collect_children_text(elem)


def _omml_to_unicode(elem) -> str:
    try:
        return _omml_to_unicode_impl(elem)
    except Exception:
        LOGGER.warning(
            "Failed to convert OMML element %s",
            local_tag(elem),
            exc_info=True,
        )
        return ""


def _text_from_run(run_elem) -> str:
    parts: list[str] = []
    for child in run_elem:
        child_tag = local_tag(child)
        if child_tag == "t":
            if child.text:
                parts.append(child.text)
            if child.tail:
                parts.append(child.tail)
        elif child_tag in ("oMath", "oMathPara"):
            parts.append(_omml_to_unicode(child))
        elif child_tag == "tab":
            parts.append("\t")
        elif child_tag == "br":
            parts.append("\n")
        elif child_tag in {"drawing", "object", "fldChar", "pict"}:
            continue
        elif child_tag == "instrText" and child.text:
            parts.append(child.text)
    return "".join(parts)


def _walk_block_for_text(elem) -> Iterable[str]:
    for child in elem:
        yield from _walk_p_child(child)


def _walk_p_child(elem) -> Iterable[str]:
    tag = local_tag(elem)

    if tag == "r":
        yield _text_from_run(elem)
        return

    if tag in ("oMath", "oMathPara"):
        yield _omml_to_unicode(elem)
        return

    if tag in ("hyperlink", "fldSimple", "sdt", "smartTag", "ins", "del", "moveFrom", "moveTo"):
        for child in elem:
            yield from _walk_p_child(child)
        return

    if tag in {"bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd", "pict"}:
        return

    for child in elem:
        yield from _walk_p_child(child)


def paragraph_full_text(paragraph) -> str:
    parts: list[str] = []
    for child in paragraph._element:
        tag = local_tag(child)
        if tag == "r":
            parts.append(_text_from_run(child))
        elif tag in ("oMath", "oMathPara"):
            parts.append(_omml_to_unicode(child))
        elif tag in ("hyperlink", "fldSimple", "sdt", "smartTag", "ins", "del"):
            parts.extend(_walk_block_for_text(child))
        elif tag in {
            "bookmarkStart",
            "bookmarkEnd",
            "proofErr",
            "pPr",
            "commentRangeStart",
            "commentRangeEnd",
        }:
            continue
        else:
            parts.extend(_walk_block_for_text(child))
    return "".join(parts).strip()


def paragraph_has_drawing(paragraph) -> bool:
    for run in paragraph.runs:
        if run._element.findall(qn("w:drawing")):
            return True
    return False
