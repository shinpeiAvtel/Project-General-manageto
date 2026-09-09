from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pydantic import BaseModel, Field

from scripts.utils import MISSING, bool_from_value, normalize_text, split_dependencies, split_people


class NormalizedTask(BaseModel):
    project_id: str = MISSING
    project_name: str = MISSING
    site: str = MISSING
    project_manager: str = MISSING
    task_id: str = MISSING
    task_name: str = MISSING
    category: str = MISSING
    start_date: str = MISSING
    end_date: str = MISSING
    progress: float = Field(default=0)
    status: str = MISSING
    resource_count: int = 0
    assigned_person: list[str] = Field(default_factory=list)
    milestone: bool = False
    dependency: list[str] = Field(default_factory=list)
    priority: str = MISSING
    last_updated: str = MISSING
    remarks: str = MISSING
    source_file: str = MISSING
    source_sheet: str = MISSING


def normalize_rows(imported: dict[str, Any]) -> list[dict[str, Any]]:
    normalized: list[dict[str, Any]] = []
    for item in imported.get("rows", []):
        row = item["data"]
        task = NormalizedTask(
            project_id=normalize_text(row.get("Project ID")),
            project_name=normalize_text(row.get("Project Name")),
            site=normalize_text(row.get("Site")),
            project_manager=normalize_text(row.get("Project Manager")),
            task_id=normalize_text(row.get("Task ID")),
            task_name=normalize_text(row.get("Task Name")),
            category=normalize_text(row.get("Category")),
            start_date=normalize_text(row.get("Start Date")),
            end_date=normalize_text(row.get("End Date")),
            progress=float(row.get("Progress %", 0) if isinstance(row.get("Progress %"), (int, float)) else 0),
            status=normalize_text(row.get("Status")),
            resource_count=int(row.get("Resource Count", 0) or 0),
            assigned_person=split_people(row.get("Assigned Person")),
            milestone=bool_from_value(row.get("Milestone")),
            dependency=split_dependencies(row.get("Dependency")),
            priority=normalize_text(row.get("Priority")),
            last_updated=normalize_text(row.get("Last Updated")),
            remarks=normalize_text(row.get("Remarks")),
            source_file=item["metadata"].get("source_file", MISSING),
            source_sheet=item["metadata"].get("source_sheet", MISSING),
        )
        normalized.append(task.model_dump())
    return normalized


def save_normalized(normalized: list[dict[str, Any]], normalized_dir: Path) -> None:
    normalized_dir.mkdir(parents=True, exist_ok=True)
    all_path = normalized_dir / "normalized_schedule.json"
    with all_path.open("w", encoding="utf-8") as f:
        json.dump(normalized, f, ensure_ascii=False, indent=2)

    grouped: dict[str, list[dict[str, Any]]] = {}
    for task in normalized:
        grouped.setdefault(task["project_name"], []).append(task)

    for project_name, tasks in grouped.items():
        safe_name = project_name.replace("/", "_").replace(" ", "_")
        with (normalized_dir / f"{safe_name}.json").open("w", encoding="utf-8") as f:
            json.dump(tasks, f, ensure_ascii=False, indent=2)


if __name__ == "__main__":
    from scripts.import_excel import import_all

    root = Path(__file__).resolve().parents[1]
    imported = import_all(root / "Input")
    normalized = normalize_rows(imported)
    save_normalized(normalized, root / "data" / "normalized")
    print(f"Normalized rows: {len(normalized)}")
