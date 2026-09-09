from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timezone
from pathlib import Path
from typing import Any

from scripts.utils import MISSING, load_yaml

REQUIRED_FIELDS = ["Project Name", "Task Name", "Start Date", "End Date"]


def _is_date_invalid(value: Any) -> bool:
    return value in {None, "", "INVALID_DATE"}


def validate_imported(imported: dict[str, Any], config_dir: Path) -> dict[str, Any]:
    schema = load_yaml(config_dir / "schedule_schema.yaml")
    allowed_status = set(schema.get("allowed_status", []))
    required_columns = set(schema.get("required_columns", []))

    issues: list[dict[str, Any]] = []
    task_ids_by_project: dict[str, set[str]] = defaultdict(set)
    all_task_ids_by_project: dict[str, set[str]] = defaultdict(set)
    checked_sheets: set[tuple[str, str]] = set()

    for item in imported.get("rows", []):
        row = item["data"]
        source_file = item["metadata"].get("source_file", "")
        source_sheet = item["metadata"].get("source_sheet", "")
        sheet_key = (source_file, source_sheet)
        if sheet_key not in checked_sheets:
            checked_sheets.add(sheet_key)
            headers = set(item["metadata"].get("headers", []))
            missing_headers = sorted(required_columns - headers)
            for h in missing_headers:
                issues.append(
                    {
                        "severity": "ERROR",
                        "code": "MISSING_HEADER",
                        "field": h,
                        "message": f"Required header missing: {h}",
                        "project_name": "UNKNOWN",
                        "task_id": "UNKNOWN",
                        "task_name": "UNKNOWN",
                        "source_file": source_file,
                        "source_sheet": source_sheet,
                    }
                )

        project_name = row.get("Project Name", MISSING)
        task_id = str(row.get("Task ID") or "").strip()
        task_name = row.get("Task Name", MISSING)
        start_date = row.get("Start Date")
        end_date = row.get("End Date")
        progress = row.get("Progress %")
        status = str(row.get("Status") or "").strip()
        last_updated = row.get("Last Updated")
        folder_project = item["metadata"].get("folder_project", "")

        for field in REQUIRED_FIELDS:
            if row.get(field) in {None, "", MISSING}:
                issues.append(_issue("ERROR", "MISSING_REQUIRED", field, item, f"{field} is missing"))

        if _is_date_invalid(start_date):
            issues.append(_issue("ERROR", "INVALID_DATE", "Start Date", item, "Invalid Start Date format"))
        if _is_date_invalid(end_date):
            issues.append(_issue("ERROR", "INVALID_DATE", "End Date", item, "Invalid End Date format"))

        if start_date not in {None, "", "INVALID_DATE"} and end_date not in {None, "", "INVALID_DATE"}:
            s = date.fromisoformat(start_date)
            e = date.fromisoformat(end_date)
            if s > e:
                issues.append(_issue("ERROR", "DATE_ORDER", "Start Date/End Date", item, "Start Date > End Date"))

        if isinstance(progress, (int, float)):
            if progress < 0:
                issues.append(_issue("ERROR", "PROGRESS_RANGE", "Progress %", item, "Progress < 0"))
            if progress > 100:
                issues.append(_issue("ERROR", "PROGRESS_RANGE", "Progress %", item, "Progress > 100"))
        else:
            issues.append(_issue("WARNING", "PROGRESS_FORMAT", "Progress %", item, "Progress is not numeric"))

        if task_id:
            if task_id in task_ids_by_project[project_name]:
                issues.append(_issue("ERROR", "DUPLICATE_TASK_ID", "Task ID", item, f"Duplicated Task ID: {task_id}"))
            task_ids_by_project[project_name].add(task_id)
            all_task_ids_by_project[project_name].add(task_id)

        if folder_project and project_name not in {None, "", MISSING} and str(project_name).strip() != folder_project:
            issues.append(_issue("ERROR", "FOLDER_PROJECT_MISMATCH", "Project Name", item, "Folder project and Excel Project Name mismatch"))

        if status and status not in allowed_status:
            issues.append(_issue("WARNING", "INVALID_STATUS", "Status", item, f"Invalid status: {status}"))

        if last_updated in {None, "", "INVALID_DATE"}:
            issues.append(_issue("WARNING", "LAST_UPDATED_MISSING", "Last Updated", item, "Last Updated missing or invalid"))

        if project_name in {None, "", MISSING}:
            issues.append(_issue("ERROR", "PROJECT_NAME_MISSING", "Project Name", item, "Project Name is missing"))
        if task_name in {None, "", MISSING}:
            issues.append(_issue("ERROR", "TASK_NAME_MISSING", "Task Name", item, "Task Name is missing"))

    for item in imported.get("rows", []):
        row = item["data"]
        project_name = row.get("Project Name", MISSING)
        dependencies = str(row.get("Dependency") or "").strip()
        if not dependencies:
            continue
        for dep in [d.strip() for d in dependencies.replace(";", ",").split(",") if d.strip()]:
            if dep not in all_task_ids_by_project[project_name]:
                issues.append(_issue("WARNING", "DEPENDENCY_NOT_FOUND", "Dependency", item, f"Dependency not found: {dep}"))

    return {
        "generated_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "total_rows": len(imported.get("rows", [])),
        "issues": issues,
    }


def _issue(severity: str, code: str, field: str, item: dict[str, Any], message: str) -> dict[str, Any]:
    row = item["data"]
    return {
        "severity": severity,
        "code": code,
        "field": field,
        "message": message,
        "project_name": row.get("Project Name", "MISSING"),
        "task_id": row.get("Task ID", "MISSING"),
        "task_name": row.get("Task Name", "MISSING"),
        "source_file": item["metadata"].get("source_file"),
        "source_sheet": item["metadata"].get("source_sheet"),
    }


if __name__ == "__main__":
    from scripts.import_excel import import_all

    root = Path(__file__).resolve().parents[1]
    imported = import_all(root / "Input")
    report = validate_imported(imported, root / "config")
    print(f"Validated rows: {report['total_rows']}, issues: {len(report['issues'])}")
