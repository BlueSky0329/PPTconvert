"""PDF 试卷 → 结构化真题 Word 的入口。"""

from __future__ import annotations

import shutil

from core.exam_docx_writer import parsed_exam_is_empty, write_parsed_exam_to_docx
from core.pdf_exam_extract import extract_pdf_line_items
from core.pdf_exam_parse import parse_line_items

def pdf_to_exam_docx(
    pdf_path: str,
    out_docx: str,
    mode: str = "all",
    font_name: str = "微软雅黑",
) -> tuple[int, int]:
    """
    解析 PDF 并写出 Word。
    返回 (资料分析篇数, 数量关系篇数)；若解析结果为空则抛出 ValueError。
    """
    items, tmp_dir = extract_pdf_line_items(pdf_path)
    try:
        exam = parse_line_items(items, mode=mode)
        if parsed_exam_is_empty(exam):
            raise ValueError(
                "未能识别篇界。请确认 PDF 文字可选中，且含「一. 政治理论」到「六. 资料分析」"
                "等篇题标记之一。"
            )
        write_parsed_exam_to_docx(exam, out_docx, font_name=font_name)
        data_n = len(exam.data_sections)
        quant_n = len(exam.quant_sections)
        return data_n, quant_n
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)
