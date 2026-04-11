from __future__ import annotations

from typing import Optional

from domain.models import ExamProject, MaterialSet, OptionNode, QuestionNode, Section, SUBJECT_DISPLAY_NAMES

_QUESTION_OPTION_LAYOUTS = {"grid", "list", "one_row"}
_OPTION_LETTERS = "ABCDEFGHIJKLMNOPQRSTUVWXYZ"


def _material_body_lines(material: MaterialSet) -> list[str]:
    if material.body_lines:
        return list(material.body_lines)
    body = (material.body or "").strip()
    return body.splitlines() if body else []


def _refresh_material_body(material: MaterialSet) -> None:
    lines = _material_body_lines(material)
    material.body_lines = lines
    material.body = "\n".join(line for line in lines if line).strip()


def _merge_material_content(dest: MaterialSet, source: MaterialSet) -> None:
    dest.body_lines = _material_body_lines(dest) + _material_body_lines(source)
    _refresh_material_body(dest)
    dest.body_assets.extend(source.body_assets)
    dest.body_regions.extend(source.body_regions)
    dest.body_regions.sort(
        key=lambda region: (
            region.page_number,
            region.y0,
            region.x0,
            region.y1,
            region.x1,
        )
    )


def _next_material_id(section: Section, base_material_id: str) -> str:
    stem = (base_material_id or "material").strip() or "material"
    existing_ids = {material.material_id for material in section.material_sets}
    suffix = 1
    candidate = f"{stem}_extra{suffix}"
    while candidate in existing_ids:
        suffix += 1
        candidate = f"{stem}_extra{suffix}"
    return candidate


def renumber_question(question: QuestionNode, new_number: str) -> None:
    question.source_number = (new_number or "").strip()


def update_question_stem(question: QuestionNode, new_stem: str) -> None:
    question.stem = (new_stem or "").strip()


def _find_option(question: QuestionNode, letter: str) -> Optional[OptionNode]:
    normalized = (letter or "").strip().upper()
    if not normalized:
        return None
    for option in question.options:
        if (option.letter or "").strip().upper() == normalized:
            return option
    return None


def update_option_text(question: QuestionNode, letter: str, new_text: str) -> bool:
    option = _find_option(question, letter)
    if option is None:
        return False
    option.text = (new_text or "").strip()
    return True


def replace_option_image(question: QuestionNode, letter: str, image_path: str | None) -> bool:
    option = _find_option(question, letter)
    if option is None:
        return False
    normalized = (image_path or "").strip()
    option.image_path = normalized or None
    return True


def clear_option_image(question: QuestionNode, letter: str) -> bool:
    return replace_option_image(question, letter, None)


def _answer_letters(value: str | None) -> list[str]:
    letters: list[str] = []
    for char in (value or "").upper():
        if "A" <= char <= "Z" and char not in letters:
            letters.append(char)
    return letters


def _capture_answer_targets(question: QuestionNode) -> tuple[list[OptionNode], list[str]]:
    matched: list[OptionNode] = []
    unmatched: list[str] = []
    for letter in _answer_letters(question.answer):
        option = _find_option(question, letter)
        if option is None:
            unmatched.append(letter)
        elif option not in matched:
            matched.append(option)
    return matched, unmatched


def _restore_answer_targets(question: QuestionNode, matched: list[OptionNode], unmatched: list[str]) -> None:
    letters: list[str] = []
    for option in question.options:
        if option in matched and option.letter not in letters:
            letters.append(option.letter)
    for letter in unmatched:
        if letter not in letters:
            letters.append(letter)
    question.answer = "".join(letters) or None


def _reletter_options(question: QuestionNode) -> None:
    for index, option in enumerate(question.options):
        if index < len(_OPTION_LETTERS):
            option.letter = _OPTION_LETTERS[index]


def move_option(question: QuestionNode, letter: str, direction: int) -> bool:
    if direction not in (-1, 1):
        return False
    option = _find_option(question, letter)
    if option is None:
        return False
    index = question.options.index(option)
    target_index = index + direction
    if target_index < 0 or target_index >= len(question.options):
        return False
    matched, unmatched = _capture_answer_targets(question)
    question.options[index], question.options[target_index] = question.options[target_index], question.options[index]
    _reletter_options(question)
    _restore_answer_targets(question, matched, unmatched)
    return True


def insert_option_after(question: QuestionNode, letter: str | None = None) -> bool:
    if len(question.options) >= len(_OPTION_LETTERS):
        return False
    if not question.options:
        question.options.append(OptionNode(letter="A", text=""))
        return True

    matched, unmatched = _capture_answer_targets(question)
    insert_at = len(question.options)
    if letter:
        option = _find_option(question, letter)
        if option is None:
            return False
        insert_at = question.options.index(option) + 1

    question.options.insert(insert_at, OptionNode(letter="", text=""))
    _reletter_options(question)
    _restore_answer_targets(question, matched, unmatched)
    return True


def remove_option(question: QuestionNode, letter: str) -> bool:
    option = _find_option(question, letter)
    if option is None:
        return False
    matched, unmatched = _capture_answer_targets(question)
    matched = [item for item in matched if item is not option]
    question.options.remove(option)
    _reletter_options(question)
    _restore_answer_targets(question, matched, unmatched)
    return True


def set_question_option_layout(question: QuestionNode, layout: str | None) -> None:
    normalized = (layout or "").strip().lower()
    question.option_layout = normalized if normalized in _QUESTION_OPTION_LAYOUTS else None


def rename_material(material: MaterialSet, new_header: str) -> None:
    material.header = (new_header or "").strip()


def remove_question(project: ExamProject, target: QuestionNode) -> bool:
    for section in project.sections:
        if section.kind == "data":
            for material in list(section.material_sets):
                if target in material.questions:
                    material.questions.remove(target)
                    _cleanup_project(project)
                    return True
        else:
            if target in section.questions:
                section.questions.remove(target)
                _cleanup_project(project)
                return True
    return False


def move_data_question(project: ExamProject, target: QuestionNode, direction: int) -> bool:
    if direction not in (-1, 1):
        return False
    for section in project.sections:
        if section.kind != "data":
            continue
        for material_index, material in enumerate(section.material_sets):
            if target not in material.questions:
                continue
            neighbor_index = material_index + direction
            if neighbor_index < 0 or neighbor_index >= len(section.material_sets):
                return False
            material.questions.remove(target)
            neighbor = section.material_sets[neighbor_index]
            if direction < 0:
                neighbor.questions.append(target)
            else:
                neighbor.questions.insert(0, target)
            _cleanup_project(project)
            return True
    return False


def locate_question(project: ExamProject, target: QuestionNode) -> tuple[Optional[Section], Optional[MaterialSet]]:
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                if target in material.questions:
                    return section, material
        elif target in section.questions:
            return section, None
    return None, None


def insert_material_after(project: ExamProject, target: MaterialSet, header: str = "新材料") -> bool:
    for section in project.sections:
        if section.kind != "data":
            continue
        for idx, material in enumerate(section.material_sets):
            if material is not target:
                continue
            new_id = _next_material_id(section, material.material_id)
            section.material_sets.insert(
                idx + 1,
                MaterialSet(
                    material_id=new_id,
                    header=(header or "新材料").strip(),
                    body="",
                ),
            )
            return True
    return False


def merge_adjacent_materials(project: ExamProject, target: MaterialSet, direction: int) -> bool:
    if direction not in (-1, 1):
        return False
    for section in project.sections:
        if section.kind != "data":
            continue
        for idx, material in enumerate(section.material_sets):
            if material is not target:
                continue
            if direction < 0:
                source_index = idx
                dest_index = idx - 1
            else:
                source_index = idx + 1
                dest_index = idx
            if source_index < 0 or source_index >= len(section.material_sets):
                return False
            if dest_index < 0 or dest_index >= len(section.material_sets):
                return False

            source = section.material_sets[source_index]
            dest = section.material_sets[dest_index]
            if source is dest:
                return False

            _merge_material_content(dest, source)
            dest.questions.extend(source.questions)
            section.material_sets.remove(source)
            _cleanup_project(project)
            return True
    return False


def reclassify_objective_section(section: Section, new_kind: str) -> bool:
    normalized = (new_kind or "").strip().lower()
    if not normalized or normalized == "data":
        return False
    if section.kind == "data":
        return False
    if section.kind == normalized:
        return False
    section.kind = normalized
    section.title = SUBJECT_DISPLAY_NAMES.get(normalized, "未知科目")
    return True


def _cleanup_project(project: ExamProject) -> None:
    filtered_sections: list[Section] = []
    for section in project.sections:
        if section.kind == "data":
            section.material_sets = [material for material in section.material_sets if material.questions]
            if section.material_sets:
                filtered_sections.append(section)
        else:
            if section.questions:
                filtered_sections.append(section)
    project.sections = filtered_sections
