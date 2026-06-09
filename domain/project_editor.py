from __future__ import annotations

import math
from typing import Optional

from domain.models import AssetRef, ExamProject, MaterialSet, OptionNode, QuestionNode, Section, SUBJECT_DISPLAY_NAMES
from core.ppt_layout import sanitize_ppt_layout

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


def reassign_stem_assets_to_options(question: QuestionNode, *, count: int | None = None) -> int:
    if not question.stem_assets or not question.options:
        return 0
    target_count = min(len(question.options), len(question.stem_assets))
    if count is not None:
        target_count = min(target_count, max(0, int(count)))
    if target_count <= 0:
        return 0

    changes = 0
    remaining_assets = list(question.stem_assets)
    for index in range(target_count):
        option = question.options[index]
        asset: AssetRef = remaining_assets[index]
        if option.image_path == asset.path:
            continue
        option.image_path = asset.path
        option.source_page = asset.source_page
        option.page_region = asset.page_region
        changes += 1
    if changes > 0:
        question.stem_assets = remaining_assets[target_count:]
    return changes


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


def set_question_ppt_layout_block(
    question: QuestionNode,
    block: str,
    rect: dict[str, float] | None,
) -> None:
    layout = dict(getattr(question, "ppt_layout", {}) or {})
    normalized_block = (block or "").strip().lower()
    if not rect:
        layout.pop(normalized_block, None)
    else:
        sanitized = sanitize_ppt_layout({normalized_block: rect})
        if normalized_block in sanitized:
            layout[normalized_block] = sanitized[normalized_block]
    question.ppt_layout = layout


def clear_question_ppt_layout(question: QuestionNode) -> None:
    question.ppt_layout = {}


def rename_material(material: MaterialSet, new_header: str) -> None:
    material.header = (new_header or "").strip()


def prepend_material_body_text(material: MaterialSet, text: str) -> bool:
    normalized = (text or "").strip()
    if not normalized:
        return False
    existing = _material_body_lines(material)
    material.body_lines = [normalized, *existing]
    _refresh_material_body(material)
    return True


def move_stem_assets_to_material(material: MaterialSet, question: QuestionNode, *, count: int | None = None) -> int:
    if material is None or not question.stem_assets:
        return 0
    move_count = len(question.stem_assets)
    if count is not None:
        move_count = min(move_count, max(0, int(count)))
    if move_count <= 0:
        return 0

    moving_assets = list(question.stem_assets[:move_count])
    remaining_assets = list(question.stem_assets[move_count:])
    existing_asset_keys = {
        (
            asset.kind,
            asset.path,
            asset.source_page,
            getattr(asset.page_region, "page_number", None),
            getattr(asset.page_region, "x0", None),
            getattr(asset.page_region, "y0", None),
            getattr(asset.page_region, "x1", None),
            getattr(asset.page_region, "y1", None),
        )
        for asset in material.body_assets
    }
    existing_region_keys = {
        (
            region.page_number,
            region.x0,
            region.y0,
            region.x1,
            region.y1,
        )
        for region in material.body_regions
    }

    changes = 0
    for asset in moving_assets:
        asset_key = (
            asset.kind,
            asset.path,
            asset.source_page,
            getattr(asset.page_region, "page_number", None),
            getattr(asset.page_region, "x0", None),
            getattr(asset.page_region, "y0", None),
            getattr(asset.page_region, "x1", None),
            getattr(asset.page_region, "y1", None),
        )
        if asset_key not in existing_asset_keys:
            material.body_assets.append(asset)
            existing_asset_keys.add(asset_key)
            changes += 1
        if asset.page_region is not None:
            region_key = (
                asset.page_region.page_number,
                asset.page_region.x0,
                asset.page_region.y0,
                asset.page_region.x1,
                asset.page_region.y1,
            )
            if region_key not in existing_region_keys:
                material.body_regions.append(asset.page_region)
                existing_region_keys.add(region_key)
    if changes > 0:
        material.body_regions.sort(
            key=lambda region: (
                region.page_number,
                region.y0,
                region.x0,
                region.y1,
                region.x1,
            )
        )
    question.stem_assets = remaining_assets
    return changes


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


def insert_question_before(project: ExamProject, target: QuestionNode, new_question: QuestionNode) -> bool:
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                if target in material.questions:
                    index = material.questions.index(target)
                    material.questions.insert(index, new_question)
                    return True
        else:
            if target in section.questions:
                index = section.questions.index(target)
                section.questions.insert(index, new_question)
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


def _question_container(
    project: ExamProject, target: QuestionNode
) -> tuple[Optional[list[QuestionNode]], int, Optional[Section], Optional[MaterialSet]]:
    """返回承载 target 的题目列表、下标、所属 section / material。找不到返回 (None, -1, None, None)。"""
    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                if target in material.questions:
                    return material.questions, material.questions.index(target), section, material
        else:
            if target in section.questions:
                return section.questions, section.questions.index(target), section, None
    return None, -1, None, None


def split_question_options(
    project: ExamProject,
    question: QuestionNode,
    split_index: int,
    *,
    new_number: str = "",
    new_stem: str = "",
) -> Optional[QuestionNode]:
    """把 ``question`` 第 ``split_index``（0-based）及之后的选项拆成新的一道题，插在其后。

    用于切题把两道题的选项粘连成一题的情形。新题题干由调用方给（可空，留给用户填）。
    答案保序：仍留在原题的答案选项归原题、被移走的归新题。返回新题或 None。
    """
    container, index, _section, _material = _question_container(project, question)
    if container is None:
        return None
    options = question.options or []
    if split_index <= 0 or split_index >= len(options):
        return None
    matched, unmatched = _capture_answer_targets(question)
    moved = options[split_index:]
    question.options = options[:split_index]
    new_question = QuestionNode(
        source_number=(new_number or "").strip(),
        stem=(new_stem or "").strip(),
        options=moved,
    )
    _reletter_options(question)
    _reletter_options(new_question)
    _restore_answer_targets(question, [opt for opt in matched if opt in question.options], unmatched)
    _restore_answer_targets(new_question, [opt for opt in matched if opt in new_question.options], [])
    container.insert(index + 1, new_question)
    return new_question


def merge_with_next_question(project: ExamProject, question: QuestionNode) -> bool:
    """把同一容器内的下一题合并进 ``question``（题干拼接、选项/题干图/页码追加），删除下一题。

    用于切题把一道题（常因跨页）拆成两题的情形。答案保序。返回是否合并成功。
    """
    container, index, _section, _material = _question_container(project, question)
    if container is None or index + 1 >= len(container):
        return False
    nxt = container[index + 1]
    current_matched, current_unmatched = _capture_answer_targets(question)
    next_matched, next_unmatched = _capture_answer_targets(nxt)
    next_stem = (nxt.stem or "").strip()
    if next_stem:
        base_stem = (question.stem or "").strip()
        question.stem = (base_stem + ("　" if base_stem else "") + next_stem).strip()
    question.options.extend(nxt.options)
    question.stem_assets.extend(nxt.stem_assets)
    seen_pages = set(question.page_numbers)
    for page in nxt.page_numbers:
        if page not in seen_pages:
            question.page_numbers.append(page)
            seen_pages.add(page)
    _reletter_options(question)
    _restore_answer_targets(
        question,
        current_matched + next_matched,
        current_unmatched + next_unmatched,
    )
    container.remove(nxt)
    _cleanup_project(project)
    return True


def insert_blank_question_after(project: ExamProject, target: QuestionNode) -> Optional[QuestionNode]:
    """在 ``target`` 之后插入一道空白题（4 个空选项），供手动补录被漏识别的整道题。"""
    container, index, _section, _material = _question_container(project, target)
    if container is None:
        return None
    new_question = QuestionNode(
        source_number="",
        stem="",
        options=[OptionNode(letter=letter, text="") for letter in "ABCD"],
    )
    container.insert(index + 1, new_question)
    return new_question


def move_question_to_kind(project: ExamProject, question: QuestionNode, target_kind: str) -> bool:
    """把单道客观题移到目标科目的篇题（没有则新建），用于单题分错科目。

    不支持 data（资料分析题靠材料组织，应走跨材料移动）。返回是否成功。
    """
    normalized = (target_kind or "").strip().lower()
    if normalized in ("", "data"):
        return False
    container, _index, section, material = _question_container(project, question)
    if container is None or material is not None or section is None:
        return False
    if section.kind == normalized:
        return False
    container.remove(question)
    target = next(
        (sec for sec in project.sections if sec.kind == normalized and not sec.material_sets),
        None,
    )
    if target is None:
        target = Section(kind=normalized, title=SUBJECT_DISPLAY_NAMES.get(normalized, "未知科目"))
        project.sections.append(target)
    target.questions.append(question)
    _cleanup_project(project)
    return True


def move_question_in_container(project: ExamProject, question: QuestionNode, direction: int) -> bool:
    """在同一篇题 / 材料内上移(-1)或下移(1)一道题，修正切题后的题序。返回是否成功。"""
    if direction not in (-1, 1):
        return False
    container, index, _section, _material = _question_container(project, question)
    if container is None:
        return False
    target_index = index + direction
    if target_index < 0 or target_index >= len(container):
        return False
    container[index], container[target_index] = container[target_index], container[index]
    return True


def set_question_answer(question: QuestionNode, answer: str | None) -> None:
    """设置 / 清空题目答案，只保留 A-Z 字母并去重（如 "a, b" -> "AB"，空 -> None）。"""
    letters: list[str] = []
    for char in (answer or "").upper():
        if "A" <= char <= "Z" and char not in letters:
            letters.append(char)
    question.answer = "".join(letters) or None


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


def reclassify_objective_section(section: Section, new_kind: str, *, project: ExamProject | None = None) -> bool:
    normalized = (new_kind or "").strip().lower()
    if not normalized or normalized == "data":
        return False
    if section.kind == "data":
        return False
    if section.kind == normalized:
        return False
    section.kind = normalized
    section.title = SUBJECT_DISPLAY_NAMES.get(normalized, "未知科目")
    if project is not None:
        _cleanup_project(project)
    return True


def section_subject_suggestion(section: Section) -> tuple[Optional[str], str]:
    if section.kind == "data":
        return None, "资料分析不走这条建议链路。"

    suggestions = [
        getattr(question, "suggested_subject", None)
        for question in section.questions
        if getattr(question, "suggested_subject", None) not in {None, "data"}
    ]
    if not suggestions:
        return None, "当前篇题没有稳定的 AI 科目建议。"

    target_kinds = {str(kind) for kind in suggestions if kind}
    if len(target_kinds) != 1:
        return None, "当前篇题内部存在冲突建议，暂不自动应用。"

    target_kind = next(iter(target_kinds))
    if target_kind == section.kind:
        return None, "AI 建议与当前篇题科目一致。"

    suggestion_count = len(suggestions)
    total_questions = len(section.questions)
    flagged_questions = [
        question
        for question in section.questions
        if getattr(question, "suggested_subject", None) == target_kind
    ]

    if section.kind == "unknown":
        return (
            target_kind,
            f"当前未知篇题中有 {suggestion_count}/{total_questions} 道题一致建议改为"
            f" {SUBJECT_DISPLAY_NAMES.get(target_kind, target_kind)}。",
        )

    required_count = max(1, math.ceil(total_questions * 0.75))
    if suggestion_count < required_count or suggestion_count != len(flagged_questions):
        return None, "当前篇题只有部分题目建议换科目，暂不作为整段安全建议。"

    return (
        target_kind,
        f"当前篇题 {suggestion_count}/{total_questions} 道题都一致建议改为"
        f" {SUBJECT_DISPLAY_NAMES.get(target_kind, target_kind)}。",
    )


def apply_section_subject_suggestion(section: Section, *, project: ExamProject | None = None) -> bool:
    target_kind, _reason = section_subject_suggestion(section)
    if not target_kind:
        return False
    return reclassify_objective_section(section, target_kind, project=project)


def apply_all_safe_subject_suggestions(project: ExamProject) -> int:
    applied = 0
    for section in list(project.sections):
        if apply_section_subject_suggestion(section, project=project):
            applied += 1
    return applied


def _refresh_selected_subjects(project: ExamProject) -> None:
    seen: list[str] = []
    for section in project.sections:
        if section.kind in SUBJECT_DISPLAY_NAMES and section.kind not in {"data", "unknown"} and section.kind not in seen:
            seen.append(section.kind)
        elif section.kind == "data" and section.kind not in seen:
            seen.append(section.kind)
    project.selected_subjects = seen


def _merge_adjacent_objective_sections(project: ExamProject) -> None:
    from core.subject_inference import preferred_subject_title, should_merge_subject_sections

    merged: list[Section] = []
    for section in project.sections:
        if (
            merged
            and section.kind != "data"
            and merged[-1].kind == section.kind
            and should_merge_subject_sections(section.kind, merged[-1].title, section.title)
        ):
            merged[-1].title = preferred_subject_title(section.kind, merged[-1].title, section.title)
            merged[-1].questions.extend(section.questions)
            continue
        merged.append(section)
    project.sections = merged


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
    _merge_adjacent_objective_sections(project)
    _refresh_selected_subjects(project)
