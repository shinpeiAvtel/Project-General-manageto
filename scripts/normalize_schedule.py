from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from pydantic import BaseModel

try:
    from scripts.utils import MISSING, as_list, bool_from_value, to_date_str, to_float, to_int, to_text
except ModuleNotFoundError:  # pragma: no cover
    from utils import MISSING, as_list, bool_from_value, to_date_str, to_float, to_int, to_text


class NormalizedTask(BaseModel):
    project_id: str
    project_name: str
    site: str
    project_manager: str
    task_id: str
    task_name: str
    category: str
    start_date: str
    end_date: str
    progress: float | str
    status: str
    resource_count: int
    assigned_person: list[str]
    milestone: bool
    dependency: list[str]
    priority: str
    last_updated: str
    remarks: str
    source_file: str
    source_sheet: str
    source_row: int
    folder_owner: str
    folder_project: str


def normalize_record(record: dict[str, Any]) -> dict[str, Any]:
    progress = to_float(record.get("Progress %"))
    task = NormalizedTask(
        project_id=to_text(record.get("Project ID")),
        project_name=to_text(record.get("Project Name")),
        site=to_text(record.get("Site")),
        project_manager=to_text(record.get("Project Manager")),
        task_id=to_text(record.get("Task ID")),
        task_name=to_text(record.get("Task Name")),
        category=to_text(record.get("Category")),
        start_date=to_date_str(record.get("Start Date")),
        end_date=to_date_str(record.get("End Date")),
        progress=progress if progress is not None else MISSING,
        status=to_text(record.get("Status")),
        resource_count=to_int(record.get("Resource Count")) or 0,
        assigned_person=as_list(record.get("Assigned Person")),
        milestone=bool_from_value(record.get("Milestone")),
        dependency=as_list(record.get("Dependency")),
        priority=to_text(record.get("Priority")),
        last_updated=to_date_str(record.get("Last Updated")),
        remarks=to_text(record.get("Remarks")),
        source_file=to_text(record.get("_source_file")),
        source_sheet=to_text(record.get("_source_sheet")),
        source_row=record.get("_source_row", -1),
        folder_owner=to_text(record.get("_folder_owner")),
        folder_project=to_text(record.get("_folder_project")),
    )
    return task.model_dump()


def normalize_records(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    return [normalize_record(r) for r in records]


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    records = json.loads((repo_root / "data" / "raw" / "imported_records.json").read_text(encoding="utf-8"))
    normalized = normalize_records(records)
    out = repo_root / "data" / "normalized" / "normalized_schedule.json"
    out.write_text(json.dumps(normalized, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Normalized records: {len(normalized)}")


if __name__ == "__main__":
    main()
