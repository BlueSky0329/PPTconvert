import re
from typing import Iterator, Optional

from docx import Document
from docx.text.paragraph import Paragraph

from core.models import Question, Option
from core.image_extractor import ImageExtractor
from core.word_math import paragraph_full_text, paragraph_has_drawing


# 匹配题目开头：（2020·上海）或（2020·上海 ）或（2025国考·地市、副省级卷）
RE_QUESTION_START = re.compile(r'^（(\d{4})[^）]*）\s*(.*)')

# 匹配章节标题：1·标题 或 1、标题 或 1．标题（后面紧跟中文，不是题目）
RE_SECTION_HEADER = re.compile(r'^\d+\s*[·、．.]\s*[\u4e00-\u9fff]')

# 匹配单个选项：A．xxx 或 A. xxx 或 A、xxx
RE_SINGLE_OPTION = re.compile(r'([A-Da-d])\s*[．.、)）]\s*(.*?)(?=\s*$)')

# 匹配选项行（可能一行有两个选项，用 tab 分隔）
RE_OPTION_LINE = re.compile(r'^[A-Da-d]\s*[．.、)）]')

# 匹配答案行
RE_ANSWER = re.compile(r'(?:答案|正确答案)\s*[：:]\s*([A-Da-d])', re.IGNORECASE)


def _paragraph_has_image(paragraph) -> bool:
    return paragraph_has_drawing(paragraph)


def _get_paragraph_text(paragraph) -> str:
    """含 Word 公式（OMML）展开后的完整文本"""
    return paragraph_full_text(paragraph)


def iter_all_paragraphs(doc) -> Iterator[Paragraph]:
    """正文 + 表格（含嵌套表格）内全部段落，避免题目在表格中时被漏读"""
    for p in doc.paragraphs:
        yield p

    def walk_table(table):
        for row in table.rows:
            for cell in row.cells:
                for p in cell.paragraphs:
                    yield p
                for nested in cell.tables:
                    yield from walk_table(nested)

    for table in doc.tables:
        yield from walk_table(table)


def _parse_options_from_line(text: str) -> list[Option]:
    """
    从一行文本中解析出选项（可能有 1~2 个，用 tab 分隔）。
    支持格式：A．xxx\tB．xxx 或 A．xxx
    """
    parts = re.split(r'\t+', text)
    options = []
    for part in parts:
        part = part.strip()
        if not part:
            continue
        m = RE_SINGLE_OPTION.match(part)
        if m:
            letter = m.group(1).upper()
            opt_text = m.group(2).strip()
            options.append(Option(letter=letter, text=opt_text))
    return options


class WordParser:
    """Word 文档选择题解析器 — 适配（年份·地区）格式"""

    def __init__(self, temp_dir: Optional[str] = None):
        self.image_extractor = ImageExtractor(temp_dir)

    def parse(self, docx_path: str) -> list[Question]:
        doc = Document(docx_path)
        questions: list[Question] = []
        current_question: Optional[Question] = None
        q_counter = 0

        for paragraph in iter_all_paragraphs(doc):
            text = _get_paragraph_text(paragraph)
            has_img = _paragraph_has_image(paragraph)

            if not text and not has_img:
                continue

            # --- 跳过章节标题行 ---
            if text and RE_SECTION_HEADER.match(text) and not RE_QUESTION_START.match(text):
                continue

            # --- 答案行 ---
            if text:
                answer_match = RE_ANSWER.search(text)
                if answer_match and current_question:
                    current_question.answer = answer_match.group(1).upper()
                    continue

            # --- 题目开头 ---
            q_match = RE_QUESTION_START.match(text) if text else None
            if q_match:
                if current_question and current_question.is_complete:
                    questions.append(current_question)

                q_counter += 1
                stem_text = q_match.group(0).strip()
                current_question = Question(
                    number=q_counter,
                    stem=stem_text,
                )

                if has_img:
                    imgs = self.image_extractor.extract_from_paragraph(
                        paragraph, q_counter
                    )
                    current_question.image_paths.extend(imgs)
                continue

            # --- 选项行 ---
            if text and RE_OPTION_LINE.match(text) and current_question:
                opts = _parse_options_from_line(text)
                current_question.options.extend(opts)

                if has_img:
                    imgs = self.image_extractor.extract_from_paragraph(
                        paragraph, q_counter
                    )
                    current_question.image_paths.extend(imgs)
                continue

            # --- 纯图片段落 ---
            if has_img and current_question:
                imgs = self.image_extractor.extract_from_paragraph(
                    paragraph, q_counter
                )
                current_question.image_paths.extend(imgs)
                continue

            # --- 其余文本追加到当前题目 ---
            if current_question and text:
                if current_question.options:
                    last_opt = current_question.options[-1]
                    last_opt.text += " " + text
                else:
                    current_question.stem += "\n" + text

        if current_question and current_question.is_complete:
            questions.append(current_question)

        return questions

    def cleanup(self):
        self.image_extractor.cleanup()
