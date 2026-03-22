"""
从 Word 段落中提取完整文本，包含公式（Office Math / OMML）。

python-docx 的 paragraph.text 会忽略 m:oMath 区域，导致根号、分式、π 等丢失。
本模块按文档顺序遍历 w:p 子节点，将 OMML 转为近似 Unicode 字符串。
"""

from __future__ import annotations

from typing import Iterable

from docx.oxml.ns import qn


def _attrib_val(elem, name: str = "val") -> str:
    for k, v in elem.attrib.items():
        if k == name or k.endswith("}" + name):
            return v or ""
    return ""


def local_tag(elem) -> str:
    if elem.tag.startswith("{"):
        return elem.tag.split("}", 1)[-1]
    return elem.tag


def _omml_to_unicode(elem) -> str:
    """将单个 OMML 节点转为可读字符串（近似）。"""
    tag = local_tag(elem)

    if tag == "t":
        return elem.text or ""

    if tag == "r":
        parts: list[str] = []
        for child in elem:
            parts.append(_omml_to_unicode(child))
        return "".join(parts)

    if tag == "rad":
        deg_txt = ""
        inner = ""
        for child in elem:
            ct = local_tag(child)
            if ct == "deg":
                deg_txt = _collect_e_content(child)
            elif ct == "e":
                inner = _omml_to_unicode(child)
        if deg_txt.strip():
            return f"({deg_txt.strip()})√({inner})"
        return f"√({inner})"

    if tag == "f":
        num = den = ""
        for child in elem:
            ct = local_tag(child)
            if ct == "num":
                num = _collect_e_content(child)
            elif ct == "den":
                den = _collect_e_content(child)
        return f"{num}/{den}"

    if tag == "sSup":
        base = sup = ""
        for child in elem:
            ct = local_tag(child)
            if ct == "e":
                base = _omml_to_unicode(child)
            elif ct == "sup":
                sup = _omml_to_unicode(child)
        if sup:
            return f"{base}^{sup}"
        return base or _concat_children(elem)

    if tag == "sSub":
        base = sub = ""
        for child in elem:
            ct = local_tag(child)
            if ct == "e":
                base = _omml_to_unicode(child)
            elif ct == "sub":
                sub = _omml_to_unicode(child)
        if sub:
            return f"{base}_{sub}"
        return base or _concat_children(elem)

    if tag == "sPre":
        return _concat_children(elem)

    if tag == "nary":
        ch = ""
        sub = sup = ""
        inner = ""
        for child in elem:
            ct = local_tag(child)
            if ct == "chr":
                ch = _attrib_val(child, "val") or (child.text or "")
            elif ct == "sub":
                sub = _collect_e_content(child)
            elif ct == "sup":
                sup = _collect_e_content(child)
            elif ct == "e":
                inner = _omml_to_unicode(child)
        if inner or ch or sub or sup:
            return _nary_symbol(ch, sub, sup, inner)
        return _concat_children(elem)

    if tag == "func":
        fname = ""
        inner = ""
        for child in elem:
            ct = local_tag(child)
            if ct == "fName":
                fname = _collect_e_content(child)
            elif ct == "e":
                inner = _omml_to_unicode(child)
        if fname:
            return f"{fname}({inner})"
        return inner or _concat_children(elem)

    if tag == "d":
        return _concat_children(elem)

    if tag == "e":
        return _concat_children(elem)

    if tag in ("oMath", "oMathPara"):
        return _concat_children(elem)

    # 其他：累加子节点
    return _concat_children(elem)


def _nary_symbol(ch: str, sub: str, sup: str, inner: str) -> str:
    sym_map = {
        "∫": "∫",
        "∑": "∑",
        "∏": "∏",
        "∐": "∐",
        "⋃": "⋃",
        "⋂": "⋂",
    }
    # chr 属性可能是十六进制如 "∫" 或 Unicode 字符
    s = ch
    if ch and ch.startswith("\\"):
        s = ch
    elif len(ch) == 1:
        s = ch
    else:
        s = sym_map.get(ch, ch or "∫")
    parts = [s]
    if sub:
        parts.append(f"_({sub})")
    if sup:
        parts.append(f"^({sup})")
    parts.append(inner)
    return "".join(parts)


def _collect_e_content(elem) -> str:
    """收集容器内所有子节点内容。"""
    return "".join(_omml_to_unicode(child) for child in elem)


def _concat_children(elem) -> str:
    return "".join(_omml_to_unicode(child) for child in elem)


def _text_from_run(run_elem) -> str:
    """w:r 内文本（含 w:t、行内 m:oMath，忽略绘图等）。"""
    parts: list[str] = []
    for child in run_elem:
        ct = local_tag(child)
        if ct == "t":
            if child.text:
                parts.append(child.text)
            if child.tail:
                parts.append(child.tail)
        elif ct in ("oMath", "oMathPara"):
            parts.append(_omml_to_unicode(child))
        elif ct == "tab":
            parts.append("\t")
        elif ct == "br":
            parts.append("\n")
        elif ct == "drawing":
            continue
        elif ct == "object":
            continue
        elif ct == "fldChar":
            continue
        elif ct == "instrText":
            if child.text:
                parts.append(child.text)
    return "".join(parts)


def _walk_block_for_text(elem) -> Iterable[str]:
    """遍历块级元素（如 hyperlink、sdt 内容）。"""
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

    if tag == "hyperlink":
        for ch in elem:
            yield from _walk_p_child(ch)
        return

    if tag in ("fldSimple", "sdt", "smartTag", "ins", "del", "moveFrom", "moveTo"):
        for ch in elem:
            yield from _walk_p_child(ch)
        return

    if tag in ("bookmarkStart", "bookmarkEnd", "proofErr", "commentRangeStart", "commentRangeEnd"):
        return

    if tag == "pict":
        return

    # 未知：尝试深入
    for ch in elem:
        yield from _walk_p_child(ch)


def paragraph_full_text(paragraph) -> str:
    """
    段落完整纯文本：含公式展开、制表符与软换行。
    与 paragraph.text 不同，会包含 OMML 公式内容。
    """
    parts: list[str] = []
    p_el = paragraph._element
    for child in p_el:
        tag = local_tag(child)
        if tag == "r":
            parts.append(_text_from_run(child))
        elif tag in ("oMath", "oMathPara"):
            parts.append(_omml_to_unicode(child))
        elif tag == "hyperlink":
            parts.extend(_walk_block_for_text(child))
        elif tag in ("fldSimple", "sdt", "smartTag", "ins", "del"):
            parts.extend(_walk_block_for_text(child))
        elif tag in ("bookmarkStart", "bookmarkEnd", "proofErr", "pPr", "commentRangeStart", "commentRangeEnd"):
            continue
        else:
            parts.extend(_walk_block_for_text(child))

    return "".join(parts).strip()


def paragraph_has_drawing(paragraph) -> bool:
    """段落是否含嵌入式图片（绘图）。"""
    for run in paragraph.runs:
        if run._element.findall(qn("w:drawing")):
            return True
    return False
