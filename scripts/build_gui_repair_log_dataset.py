from __future__ import annotations

import argparse
import json
from collections import Counter
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from exporters.manifest_json import load_project_manifest_project


DEFAULT_OUTPUT = ROOT / "data" / "datasets" / "gui_repair_logs.jsonl"


def _iter_manifest_paths(inputs: list[str]) -> list[Path]:
    paths: list[Path] = []
    for item in inputs:
        path = Path(item)
        if path.is_dir():
            paths.extend(sorted(path.rglob("*.json")))
        elif any(token in item for token in ("*", "?")):
            paths.extend(sorted(Path().glob(item)))
        else:
            paths.append(path)
    deduped: list[Path] = []
    seen: set[str] = set()
    for path in paths:
        resolved = str(path.resolve())
        if resolved in seen:
            continue
        seen.add(resolved)
        deduped.append(path)
    return deduped


def build_gui_repair_log_dataset(manifest_inputs: list[str], output_path: Path) -> dict:
    manifest_paths = _iter_manifest_paths(manifest_inputs)
    rows: list[dict] = []
    action_counter: Counter[str] = Counter()
    source_counter: Counter[str] = Counter()

    for manifest_path in manifest_paths:
        if not manifest_path.exists():
            continue
        try:
            project = load_project_manifest_project(str(manifest_path))
        except Exception:
            continue
        source_pdf = str(getattr(getattr(project, "source", None), "pdf_path", "") or "").strip()
        for entry in getattr(project, "repair_log", []) or []:
            if not (entry.question_id and entry.before_state and entry.after_state):
                continue
            row = {
                "session_id": getattr(project, "repair_session_id", "") or "",
                "manifest_path": str(manifest_path),
                "source_pdf": source_pdf,
                "timestamp": entry.timestamp,
                "source": entry.source,
                "action": entry.action,
                "question_id": entry.question_id,
                "question_no": entry.question_no,
                "section_kind": entry.section_kind,
                "material_id": entry.material_id,
                "before_state": dict(entry.before_state or {}),
                "after_state": dict(entry.after_state or {}),
                "metadata": dict(entry.metadata or {}),
            }
            rows.append(row)
            action_counter[entry.action] += 1
            source_counter[entry.source] += 1

    output_path.parent.mkdir(parents=True, exist_ok=True)
    with output_path.open("w", encoding="utf-8") as fh:
        for row in rows:
            fh.write(json.dumps(row, ensure_ascii=False) + "\n")

    summary = {
        "output": str(output_path),
        "manifest_count": len(manifest_paths),
        "row_count": len(rows),
        "action_distribution": dict(action_counter),
        "source_distribution": dict(source_counter),
    }
    output_path.with_suffix(".summary.json").write_text(
        json.dumps(summary, ensure_ascii=False, indent=2),
        encoding="utf-8",
    )
    return summary


def main() -> None:
    parser = argparse.ArgumentParser(description="从工程 JSON 中汇总 GUI 修题日志数据集")
    parser.add_argument("inputs", nargs="+", help="工程 JSON 文件、目录或 glob")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    args = parser.parse_args()

    summary = build_gui_repair_log_dataset(args.inputs, args.output)
    print(json.dumps(summary, ensure_ascii=False, indent=2))


if __name__ == "__main__":
    main()
