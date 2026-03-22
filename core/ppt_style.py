"""PPT 样式：颜色解析与对齐映射"""

from typing import Optional

from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN


def parse_hex_color(s: str) -> Optional[RGBColor]:
    """将 #RRGGBB 或 RRGGBB 转为 RGBColor，失败返回 None。"""
    if not s:
        return None
    s = s.strip().lstrip("#")
    if len(s) != 6:
        return None
    try:
        r = int(s[0:2], 16)
        g = int(s[2:4], 16)
        b = int(s[4:6], 16)
        return RGBColor(r, g, b)
    except ValueError:
        return None


def align_from_string(name: str) -> int:
    """left / center / right -> PP_ALIGN"""
    m = {
        "left": PP_ALIGN.LEFT,
        "center": PP_ALIGN.CENTER,
        "right": PP_ALIGN.RIGHT,
        "justify": PP_ALIGN.JUSTIFY,
    }
    return m.get((name or "left").lower(), PP_ALIGN.LEFT)
