from __future__ import annotations

from pathlib import Path
from typing import Any

from .read_personal_schedules import read_personal_schedules
from .utils import (
    VALID_STATUSES,
    cell_to_bool,
    format_date,
    format_time,
    load_schedule_config,
    normalise_text,
    parse_excel_date,
    parse_excel_time,
    repo_root,
    save_json,
    serialise_record,
)


def build_issue(record: dict[str, Any], severity: str, code: str, message: str) -> dict[str, Any]:
    return {
        "severity": severity,
        "code": code,
        "message": message,
        "person": record.get("person") or "",
        "source_sheet": record.get("source_sheet") or "",
        "source_row": record.get("source_row") or "",
    }


def validate_records(raw_records: list[dict[str, Any]]) -> tuple[list[dict[str, Any]], list[dict[str, Any]]]:
    issues: list[dict[str, Any]] = []
    valid_records: list[dict[str, Any]] = []
    seen_duplicates: set[tuple[Any, ...]] = set()

    for raw_record in raw_records:
        record_issues: list[dict[str, Any]] = []
        person = normalise_text(raw_record.get("person"))
        schedule_task = normalise_text(raw_record.get("schedule_task"))
        project = normalise_text(raw_record.get("project")) or "UNKNOWN"
        site = normalise_text(raw_record.get("site")) or "UNKNOWN"
        status = normalise_text(raw_record.get("status"))
        remarks = normalise_text(raw_record.get("remarks"))
        category = normalise_text(raw_record.get("category"))
        priority = normalise_text(raw_record.get("priority"))
        last_updated_raw = raw_record.get("last_updated")
        last_updated_date = parse_excel_date(last_updated_raw)
        last_updated = format_date(last_updated_date) if last_updated_date else normalise_text(last_updated_raw)
        parsed_date = parse_excel_date(raw_record.get("date"))
        parsed_start_time = parse_excel_time(raw_record.get("start_time"))
        parsed_end_time = parse_excel_time(raw_record.get("end_time"))
        all_day = cell_to_bool(raw_record.get("all_day"))

        if not person:
            record_issues.append(build_issue(raw_record, "ERROR", "MISSING_PERSON", "Person name is missing."))
        if raw_record.get("date") in (None, ""):
            record_issues.append(build_issue(raw_record, "ERROR", "DATE_MISSING", "Date is missing."))
        elif not parsed_date:
            record_issues.append(build_issue(raw_record, "ERROR", "INVALID_DATE", "Date must use YYYY-MM-DD."))
        if not schedule_task:
            record_issues.append(build_issue(raw_record, "ERROR", "MISSING_TASK", "Schedule / Task is missing."))
        if raw_record.get("start_time") not in (None, "") and not parsed_start_time:
            record_issues.append(build_issue(raw_record, "ERROR", "INVALID_START_TIME", "Start Time must use HH:MM."))
        if raw_record.get("end_time") not in (None, "") and not parsed_end_time:
            record_issues.append(build_issue(raw_record, "ERROR", "INVALID_END_TIME", "End Time must use HH:MM."))
        if ((parsed_start_time and not parsed_end_time) or (parsed_end_time and not parsed_start_time)) and not all_day:
            record_issues.append(build_issue(raw_record, "ERROR", "INVALID_TIME", "Start Time and End Time must both be present for timed schedules."))
        if parsed_start_time and parsed_end_time and parsed_start_time > parsed_end_time:
            record_issues.append(build_issue(raw_record, "ERROR", "TIME_ORDER", "Start Time cannot be after End Time."))
        if status and status not in VALID_STATUSES:
            record_issues.append(build_issue(raw_record, "WARNING", "UNKNOWN_STATUS", f"Unknown Status '{status}'."))

        normalized_record = {
            "date": parsed_date,
            "day": parsed_date.strftime("%a") if parsed_date else "",
            "person": person,
            "project": project,
            "site": site,
            "schedule_task": schedule_task,
            "start_time": parsed_start_time,
            "end_time": parsed_end_time,
            "status": status,
            "category": category,
            "priority": priority,
            "remarks": remarks,
            "all_day": all_day,
            "source_sheet": raw_record.get("source_sheet"),
            "source_row": raw_record.get("source_row"),
            "source_workbook": raw_record.get("source_workbook"),
            "last_updated": last_updated,
        }

        duplicate_key = (
            format_date(parsed_date),
            person,
            project,
            site,
            schedule_task,
            format_time(parsed_start_time),
            format_time(parsed_end_time),
            status,
            all_day,
        )
        if duplicate_key in seen_duplicates:
            record_issues.append(build_issue(raw_record, "WARNING", "DUPLICATE_SCHEDULE", "Duplicate schedule detected."))
        elif parsed_date and person and schedule_task:
            seen_duplicates.add(duplicate_key)

        issues.extend(record_issues)
        if not any(issue["severity"] == "ERROR" for issue in record_issues):
            valid_records.append(normalized_record)

    return valid_records, issues


def run_validation(config_path: str | Path | None = None) -> dict[str, Any]:
    config = load_schedule_config(config_path)
    input_path = repo_root() / config["workbook"]["input_path"]
    normalized_path = repo_root() / config["workbook"]["normalized_json_path"]
    validation_path = repo_root() / config["workbook"]["validation_report_path"]

    payload = read_personal_schedules(input_path)
    records, issues = validate_records(payload["records"])
    save_json(
        normalized_path,
        {
            "input_workbook": str(input_path),
            "detected_personal_sheets": payload["detected_personal_sheets"],
            "record_count": len(records),
            "records": [serialise_record(record) for record in records],
        },
    )
    save_json(
        validation_path,
        {
            "input_workbook": str(input_path),
            "detected_personal_sheets": payload["detected_personal_sheets"],
            "issue_count": len(issues),
            "issues": issues,
        },
    )
    return {
        "records": records,
        "issues": issues,
        "detected_personal_sheets": payload["detected_personal_sheets"],
        "normalized_path": str(normalized_path),
        "validation_path": str(validation_path),
    }


def main() -> None:
    result = run_validation()
    severities: dict[str, int] = {"ERROR": 0, "WARNING": 0, "INFO": 0}
    for issue in result["issues"]:
        severities[issue["severity"]] = severities.get(issue["severity"], 0) + 1
    print(f"Validated records: {len(result['records'])}")
    print(f"Issues: ERROR={severities['ERROR']} WARNING={severities['WARNING']} INFO={severities['INFO']}")


if __name__ == "__main__":
    main()
