"""常用字体列表（优先显示），其余从系统补充。"""

CORE_FONTS = [
    "微软雅黑",
    "宋体",
    "黑体",
    "楷体",
    "仿宋",
    "华文细黑",
    "Arial",
    "Arial Black",
    "Calibri",
    "Cambria",
    "Consolas",
    "Segoe UI",
    "Tahoma",
    "Times New Roman",
    "Verdana",
]


def build_font_values() -> list[str]:
    import tkinter.font as tkfont

    seen = set()
    out: list[str] = []
    for name in CORE_FONTS:
        if name not in seen:
            seen.add(name)
            out.append(name)
    try:
        for name in sorted(tkfont.families()):
            if name.startswith("@"):
                continue
            if name not in seen:
                seen.add(name)
                out.append(name)
    except Exception:
        pass
    return out
