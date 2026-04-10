from __future__ import annotations

from typing import Optional

from domain.models import ExamProject, MaterialSet, QuestionNode, Section


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
