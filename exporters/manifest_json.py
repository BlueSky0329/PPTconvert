from __future__ import annotations

from dataclasses import asdict
import json
import os

from domain.models import ExamProject


def export_project_manifest(project: ExamProject, out_path: str) -> str:
    output_dir = os.path.dirname(os.path.abspath(out_path))
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)
    with open(out_path, "w", encoding="utf-8") as file_obj:
        json.dump(asdict(project), file_obj, ensure_ascii=False, indent=2)
    return out_path


def load_project_manifest(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as file_obj:
        return json.load(file_obj)
