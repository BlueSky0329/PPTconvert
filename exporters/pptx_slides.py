from __future__ import annotations

import tempfile

from core.models import Option, Question
from core.ppt_generator import PPTConfig, PPTGenerator
from domain.models import ExamProject
from exporters.material_crops import crop_material_regions


def project_to_ppt_questions(
    project: ExamProject,
    *,
    material_image_map: dict[str, list[str]] | None = None,
) -> list[Question]:
    questions: list[Question] = []
    material_image_map = material_image_map or {}
    display_number = 1

    for section in project.sections:
        if section.kind == "data":
            for material in section.material_sets:
                material_images = material_image_map.get(material.material_id, [])
                fallback_images = [asset.path for asset in material.body_assets if asset.path]
                for question in material.questions:
                    image_paths = list(material_images or fallback_images)
                    image_paths.extend(asset.path for asset in question.stem_assets if asset.path)
                    questions.append(
                        Question(
                            number=display_number,
                            stem=question.stem,
                            options=[
                                Option(
                                    letter=option.letter,
                                    text=option.text,
                                    image_path=option.image_path,
                                )
                                for option in question.options
                            ],
                            image_paths=image_paths,
                            source_question_number=question.source_number or None,
                            material_header=None if material_images else material.header or None,
                            material_text=None if material_images else (material.body or None),
                        )
                    )
                    display_number += 1
        else:
            for question in section.questions:
                questions.append(
                    Question(
                        number=display_number,
                        stem=question.stem,
                        options=[
                            Option(
                                letter=option.letter,
                                text=option.text,
                                image_path=option.image_path,
                            )
                            for option in question.options
                        ],
                        image_paths=[asset.path for asset in question.stem_assets if asset.path],
                        source_question_number=question.source_number or None,
                    )
                )
                display_number += 1
    return questions


def export_project_to_pptx(
    project: ExamProject,
    out_path: str,
    *,
    template_path: str | None = None,
    config: PPTConfig | None = None,
) -> str:
    crop_dir = tempfile.mkdtemp(prefix="pptconvert_material_crops_")
    material_image_map: dict[str, list[str]] = {}
    try:
        pdf_path = project.source.pdf_path or ""
        for section in project.sections:
            if section.kind != "data":
                continue
            for material in section.material_sets:
                material_image_map[material.material_id] = crop_material_regions(
                    pdf_path,
                    material,
                    crop_dir,
                )
        generator = PPTGenerator(config=config or PPTConfig())
        generator.generate(
            project_to_ppt_questions(project, material_image_map=material_image_map),
            out_path,
            template_path=template_path,
        )
        return out_path
    finally:
        try:
            import shutil

            shutil.rmtree(crop_dir, ignore_errors=True)
        except Exception:
            pass
