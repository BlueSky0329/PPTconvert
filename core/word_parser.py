import logging
import re
from typing import Iterator, Optional

from docx import Document
from docx.text.paragraph import Paragraph

from core.image_extractor import ImageExtractor
from core.models import Option, Question
from core.word_math import paragraph_full_text, paragraph_has_drawing

LOGGER = logging.getLogger(__name__)

SECTION_HEADER = re.compile(r"^\d+\s*[.\uFF0E\u3001]\s*[\u4e00-\u9fff]+$")
OPTION_MARKER = re.compile(
    r"(?<![A-Za-z])([A-Z])\s*[.\uFF0E\u3001)\uFF09]\s*",
    re.IGNORECASE,
)
OPTION_LINE = re.compile(r"^\s*[A-Z]\s*[.\uFF0E\u3001)\uFF09]\s*", re.IGNORECASE)
ANSWER_LINE = re.compile(
    r"(?:\u7b54\u6848|\u6b63\u786e\u7b54\u6848)\s*[\uff1a:]\s*([A-Za-z](?:\s*[,/\u3001\uff0c\s]\s*[A-Za-z])*)",
    re.IGNORECASE,
)

QUESTION_START_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    (
        "year_area",
        re.compile(
            r"^(?P<label>[\uFF08(]\s*\d{4}(?:\s*[\u00b7.\uFF0E-]\s*[^)\uFF09]+)?\s*[)\uFF09])\s*(?P<stem>.*)$"
        ),
    ),
    (
        "chinese_number",
        re.compile(
            r"^\u7b2c\s*(?P<number>\d{1,3})\s*\u9898(?:\s*[\uff1a:.\uFF0E\u3001)\uFF09]\s*)?(?P<stem>.*)$"
        ),
    ),
    (
        "numeric_paren",
        re.compile(r"^[\uFF08(]\s*(?P<number>\d{1,3})\s*[)\uFF09]\s*(?P<stem>.*)$"),
    ),
    (
        "numeric_plain",
        re.compile(r"^(?P<number>\d{1,3})\s*[.\uFF0E\u3001)\uFF09]\s*(?P<stem>.*)$"),
    ),
)


def _paragraph_has_image(paragraph) -> bool:
    return paragraph_has_drawing(paragraph)


def _get_paragraph_text(paragraph) -> str:
    return paragraph_full_text(paragraph)


def iter_all_paragraphs(doc) -> Iterator[Paragraph]:
    for paragraph in doc.paragraphs:
        yield paragraph

    def walk_table(table):
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    yield paragraph
                for nested in cell.tables:
                    yield from walk_table(nested)

    for table in doc.tables:
        yield from walk_table(table)


def _match_question_start(text: str) -> Optional[tuple[Optional[str], str]]:
    stripped = text.strip()
    for pattern_name, pattern in QUESTION_START_PATTERNS:
        match = pattern.match(stripped)
        if not match:
            continue
        if pattern_name == "year_area":
            return match.group("label").strip(), (match.group("stem") or "").strip()
        return None, (match.groupdict().get("stem") or "").strip()
    return None


def _normalize_answer(raw: str) -> Optional[str]:
    letters: list[str] = []
    for letter in re.findall(r"[A-Za-z]", raw.upper()):
        if letter not in letters:
            letters.append(letter)
    return "".join(letters) or None


def _parse_options_from_line(text: str) -> list[Option]:
    matches = list(OPTION_MARKER.finditer(text))
    if not matches:
        return []

    options: list[Option] = []
    for idx, match in enumerate(matches):
        letter = match.group(1).upper()
        body_start = match.end()
        body_end = matches[idx + 1].start() if idx + 1 < len(matches) else len(text)
        option_text = text[body_start:body_end].strip()
        options.append(Option(letter=letter, text=option_text))
    return options


class WordParser:
    def __init__(self, temp_dir: Optional[str] = None):
        self.image_extractor = ImageExtractor(temp_dir)

    def _append_images(self, question: Question, paragraph, question_number: int) -> None:
        question.image_paths.extend(
            self.image_extractor.extract_from_paragraph(paragraph, question_number)
        )

    def _finalize_question(
        self,
        questions: list[Question],
        current_question: Optional[Question],
    ) -> Optional[Question]:
        if current_question is None:
            return None

        if current_question.is_complete:
            questions.append(current_question)
            return None

        if current_question.display_stem or current_question.options or current_question.image_paths:
            LOGGER.warning(
                "Skipping incomplete question %s with %s option(s)",
                current_question.number,
                len(current_question.options),
            )
        return None

    def parse(self, docx_path: str) -> list[Question]:
        doc = Document(docx_path)
        questions: list[Question] = []
        current_question: Optional[Question] = None
        q_counter = 0

        for paragraph in iter_all_paragraphs(doc):
            text = _get_paragraph_text(paragraph).strip()
            has_img = _paragraph_has_image(paragraph)

            if not text and not has_img:
                continue

            matched_question = _match_question_start(text) if text else None
            if text and SECTION_HEADER.match(text) and matched_question is None:
                continue

            if text and current_question:
                answer_match = ANSWER_LINE.search(text)
                if answer_match:
                    normalized_answer = _normalize_answer(answer_match.group(1))
                    if normalized_answer:
                        current_question.answer = normalized_answer
                    continue

            if matched_question is not None:
                source_label, stem = matched_question
                current_question = self._finalize_question(questions, current_question)
                q_counter += 1
                current_question = Question(
                    number=q_counter,
                    stem=stem,
                    source_label=source_label,
                )
                if has_img:
                    self._append_images(current_question, paragraph, q_counter)
                continue

            if text and current_question and OPTION_LINE.match(text):
                options = _parse_options_from_line(text)
                if not options:
                    LOGGER.warning(
                        "Question %s looks like an option line but could not be parsed: %r",
                        current_question.number,
                        text,
                    )
                else:
                    current_question.options.extend(options)
                if has_img:
                    self._append_images(current_question, paragraph, q_counter)
                continue

            if has_img and current_question:
                self._append_images(current_question, paragraph, q_counter)
                continue

            if current_question and text:
                if current_question.options:
                    last_option = current_question.options[-1]
                    last_option.text = " ".join(
                        part for part in (last_option.text, text) if part
                    )
                else:
                    current_question.stem = "\n".join(
                        part for part in (current_question.stem, text) if part
                    )

        self._finalize_question(questions, current_question)
        return questions

    def cleanup(self):
        self.image_extractor.cleanup()

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc, tb):
        self.cleanup()
