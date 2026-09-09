from __future__ import annotations

import json
from collections import defaultdict
from pathlib import Path
from typing import Any

try:
    from scripts.utils import load_yaml, parse_date, to_float, to_text
except ModuleNotFoundError:  # pragma: no cover
    from utils import load_yaml, parse_date, to_float, to_text


def issue(severity: str, code: str, message: str, record: dict[str, Any]) -> dict[str, Any]:
    return {
        "severity": severity,
        "code": code,
        "message": message,
        "project_name": to_text(record.get("Project Name"), default="MISSING"),
        "task_id": to_text(record.get("Task ID"), default="MISSING"),
        "task_name": to_text(record.get("Task Name"), default="MISSING"),
        "source_file": record.get("_source_file", "MISSING"),
        "source_sheet": record.get("_source_sheet", "MISSING"),
        "source_row": record.get("_source_row", -1),
    }


def validate_records(records: list[dict[str, Any]], schema: dict[str, Any]) -> list[dict[str, Any]]:
    issues: list[dict[str, Any]] = []
    allowed_status = set(schema.get("allowed_status", []))
    required_fields = schema.get("required_fields_for_validation", [])

    task_ids_by_project: dict[str, set[str]] = defaultdict(set)

    for r in records:
        for field in required_fields:
            val = r.get(field)
            if val is None or str(val).strip() == "":
                issues.append(issue("ERROR", "MISSING_REQUIRED", f"{field} is required", r))

        start = parse_date(r.get("Start Date"))
        end = parse_date(r.get("End Date"))
        if r.get("Start Date") not in (None, "") and start is None:
            issues.append(issue("ERROR", "INVALID_DATE", "Start Date format is invalid", r))
        if r.get("End Date") not in (None, "") and end is None:
            issues.append(issue("ERROR", "INVALID_DATE", "End Date format is invalid", r))
        if start and end and start > end:
            issues.append(issue("ERROR", "START_AFTER_END", "Start Date is after End Date", r))

        progress = to_float(r.get("Progress %"))
        if progress is not None and progress < 0:
            issues.append(issue("ERROR", "PROGRESS_LT_0", "Progress is below 0", r))
        if progress is not None and progress > 100:
            issues.append(issue("ERROR", "PROGRESS_GT_100", "Progress is above 100", r))

        status = str(r.get("Status", "")).strip()
        if status and status not in allowed_status:
            issues.append(issue("WARNING", "INVALID_STATUS", f"Status '{status}' is not allowed", r))

        if r.get("Last Updated") in (None, ""):
            issues.append(issue("WARNING", "MISSING_LAST_UPDATED", "Last Updated is missing", r))

        project_name = str(r.get("Project Name", "")).strip()
        folder_project = str(r.get("_folder_project", "")).strip()
        if project_name and folder_project and folder_project != "MISSING" and project_name != folder_project:
            issues.append(issue("ERROR", "PROJECT_MISMATCH", "Project folder and Excel Project Name mismatch", r))

        task_id = str(r.get("Task ID", "")).strip()
        if project_name and task_id:
            if task_id in task_ids_by_project[project_name]:
                issues.append(issue("ERROR", "DUPLICATE_TASK_ID", f"Duplicate Task ID: {task_id}", r))
            task_ids_by_project[project_name].add(task_id)

    for project_name in {str(r.get("Project Name", "")).strip() for r in records if str(r.get("Project Name", "")).strip()}:
        project_records = [r for r in records if str(r.get("Project Name", "")).strip() == project_name]
        existing_ids = {str(r.get("Task ID", "")).strip() for r in project_records if str(r.get("Task ID", "")).strip()}
        for r in project_records:
            raw_deps = r.get("Dependency")
            if raw_deps is None:
                continue
            deps = str(raw_deps).strip()
            if not deps or deps.lower() == "none":
                continue
            dep_ids = [d.strip() for d in deps.replace(";", ",").split(",") if d.strip()]
            missing = [d for d in dep_ids if d not in existing_ids]
            if missing:
                issues.append(issue("WARNING", "MISSING_DEPENDENCY", f"Dependency not found: {', '.join(missing)}", r))

    return issues


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    schema = load_yaml(repo_root / "config" / "schedule_schema.yaml")
    records = json.loads((repo_root / "data" / "raw" / "imported_records.json").read_text(encoding="utf-8"))
    import_issues_path = repo_root / "data" / "raw" / "import_issues.json"
    import_issues = json.loads(import_issues_path.read_text(encoding="utf-8")) if import_issues_path.exists() else []
    validation_issues = validate_records(records, schema)
    all_issues = import_issues + validation_issues
    out = repo_root / "data" / "master" / "validation_report.json"
    out.write_text(json.dumps(all_issues, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"Validation issues: {len(all_issues)}")


if __name__ == "__main__":
    main()
