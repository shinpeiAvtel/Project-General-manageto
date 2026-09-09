from __future__ import annotations

from collections import defaultdict
from datetime import date, datetime, timedelta, timezone
from typing import Any

from scripts.utils import load_yaml


def analyze(normalized_rows: list[dict[str, Any]], validation_report: dict[str, Any], config_dir) -> dict[str, Any]:
    rules = load_yaml(config_dir / "project_rules.yaml")
    today = date.today()
    upcoming_days = int(rules.get("upcoming_days", 14))
    warning_days = int(rules.get("stale_days_warning", 7))
    high_warning_days = int(rules.get("stale_days_high_warning", 14))
    scoring = rules.get("risk_scoring", {})
    important_categories = set(rules.get("important_categories", []))

    delayed_tasks: list[dict[str, Any]] = []
    not_started_late_tasks: list[dict[str, Any]] = []
    milestone_candidates: list[dict[str, Any]] = []
    milestones: list[dict[str, Any]] = []
    upcoming: list[dict[str, Any]] = []

    issues_by_project: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for issue in validation_report.get("issues", []):
        issues_by_project[issue.get("project_name", "MISSING")].append(issue)

    for row in normalized_rows:
        start = _safe_date(row.get("start_date"))
        end = _safe_date(row.get("end_date"))
        progress = float(row.get("progress", 0))

        days_remaining = None
        if end:
            days_remaining = (end - today).days
        row["days_remaining"] = days_remaining

        row["delay_status"] = "ON_TRACK"
        if end and today > end and progress < 100:
            row["delay_status"] = "DELAYED"
            delayed_tasks.append(row)
        elif start and today > start and progress == 0:
            row["delay_status"] = "NOT_STARTED_LATE"
            not_started_late_tasks.append(row)

        if row.get("milestone") is True:
            milestones.append(row)
        elif row.get("category") in important_categories:
            candidate = dict(row)
            candidate["flag"] = "MILESTONE_CANDIDATE"
            milestone_candidates.append(candidate)

        if _is_upcoming(start, end, today, upcoming_days):
            upcoming.append(row)

    conflicts = _find_resource_conflicts(normalized_rows)
    project_summary = _build_project_summary(
        normalized_rows,
        delayed_tasks,
        not_started_late_tasks,
        conflicts,
        issues_by_project,
        scoring,
        warning_days,
        high_warning_days,
    )

    return {
        "generated_at": datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"),
        "project_summary": project_summary,
        "delayed_tasks": delayed_tasks,
        "not_started_late_tasks": not_started_late_tasks,
        "milestones": milestones,
        "milestone_candidates": milestone_candidates,
        "resource_conflicts": conflicts,
        "upcoming_14_days": sorted(
            upcoming,
            key=lambda r: (
                0 if r.get("priority", "").upper() == "HIGH" else 1,
                0 if r.get("milestone") else 1,
                r.get("end_date", "9999-12-31"),
            ),
        ),
    }


def _safe_date(value: Any) -> date | None:
    if not value or value in {"MISSING", "INVALID_DATE", "NOT PROVIDED"}:
        return None
    try:
        return date.fromisoformat(str(value))
    except ValueError:
        return None


def _is_upcoming(start: date | None, end: date | None, today: date, days: int) -> bool:
    horizon = today + timedelta(days=days)
    if start and today <= start <= horizon:
        return True
    if end and today <= end <= horizon:
        return True
    return False


def _daterange(start: date, end: date):
    cursor = start
    while cursor <= end:
        yield cursor
        cursor += timedelta(days=1)


def _find_resource_conflicts(rows: list[dict[str, Any]]) -> list[dict[str, Any]]:
    assignments: dict[tuple[str, str], list[str]] = defaultdict(list)
    for row in rows:
        start = _safe_date(row.get("start_date"))
        end = _safe_date(row.get("end_date"))
        if not start or not end or end < start:
            continue
        people = set(row.get("assigned_person", []))
        if row.get("project_manager") and row.get("project_manager") != "MISSING":
            people.add(row.get("project_manager"))
        for d in _daterange(start, end):
            date_key = d.isoformat()
            for person in people:
                if person and person != "MISSING":
                    assignments[(person, date_key)].append(row.get("project_name", "MISSING"))

    conflicts: list[dict[str, Any]] = []
    for (person, day), projects in assignments.items():
        uniq = sorted(set(projects))
        if len(uniq) > 1:
            conflicts.append(
                {
                    "person": person,
                    "date": day,
                    "projects": uniq,
                    "status": "POTENTIAL_CONFLICT",
                }
            )
    return conflicts


def _build_project_summary(
    rows: list[dict[str, Any]],
    delayed_tasks: list[dict[str, Any]],
    not_started_late_tasks: list[dict[str, Any]],
    conflicts: list[dict[str, Any]],
    issues_by_project: dict[str, list[dict[str, Any]]],
    scoring: dict[str, int],
    warning_days: int,
    high_warning_days: int,
) -> list[dict[str, Any]]:
    grouped: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in rows:
        grouped[row.get("project_name", "MISSING")].append(row)

    conflict_projects = defaultdict(int)
    for c in conflicts:
        for p in c.get("projects", []):
            conflict_projects[p] += 1

    today = date.today()
    results: list[dict[str, Any]] = []
    for project_name, tasks in grouped.items():
        start_dates = [_safe_date(t.get("start_date")) for t in tasks if _safe_date(t.get("start_date"))]
        end_dates = [_safe_date(t.get("end_date")) for t in tasks if _safe_date(t.get("end_date"))]
        total = len(tasks)
        completed = len([t for t in tasks if float(t.get("progress", 0)) >= 100])
        delayed = len([t for t in tasks if t.get("delay_status") == "DELAYED"])
        milestones = len([t for t in tasks if t.get("milestone") is True])
        manager = tasks[0].get("project_manager", "MISSING")
        latest_updated_days = _calc_stale_days(tasks, today)

        score = 0
        score += delayed * int(scoring.get("delayed_task", 3))
        score += len([t for t in tasks if t.get("delay_status") == "DELAYED" and str(t.get("priority", "")).upper() == "HIGH"]) * int(
            scoring.get("high_priority_delayed_task", 5)
        )
        score += len([t for t in tasks if t.get("delay_status") == "DELAYED" and t.get("milestone") is True]) * int(scoring.get("milestone_delay", 5))
        score += (1 if conflict_projects.get(project_name, 0) > 0 else 0) * int(scoring.get("resource_conflict", 2))

        if latest_updated_days is not None and latest_updated_days > high_warning_days:
            score += int(scoring.get("stale_over_high_warning", 4))
            update_status = "HIGH WARNING"
        elif latest_updated_days is not None and latest_updated_days > warning_days:
            score += int(scoring.get("stale_over_warning", 2))
            update_status = "WARNING"
        else:
            update_status = "OK"

        score += len([t for t in tasks if t.get("delay_status") == "NOT_STARTED_LATE"]) * int(scoring.get("not_started_late", 2))
        score += len([i for i in issues_by_project.get(project_name, []) if i.get("severity") == "ERROR"]) * int(scoring.get("validation_error", 2))

        risk_level = _risk_level(score)
        results.append(
            {
                "project_name": project_name,
                "project_manager": manager,
                "start_date": min(start_dates).isoformat() if start_dates else "MISSING",
                "end_date": max(end_dates).isoformat() if end_dates else "MISSING",
                "total_tasks": total,
                "completed_tasks": completed,
                "progress": round((completed / total) * 100, 1) if total else 0,
                "delayed_tasks": delayed,
                "milestones": milestones,
                "last_updated_days_ago": latest_updated_days if latest_updated_days is not None else "MISSING",
                "update_status": update_status,
                "risk_score": score,
                "risk_level": risk_level,
                "validation_errors": len([i for i in issues_by_project.get(project_name, []) if i.get("severity") == "ERROR"]),
                "not_started_late": len([t for t in tasks if t.get("delay_status") == "NOT_STARTED_LATE"]),
                "resource_conflict_count": conflict_projects.get(project_name, 0),
            }
        )
    return sorted(results, key=lambda x: (-x["risk_score"], x["project_name"]))


def _calc_stale_days(tasks: list[dict[str, Any]], today: date) -> int | None:
    dates: list[date] = []
    for t in tasks:
        d = _safe_date(t.get("last_updated"))
        if d:
            dates.append(d)
    if not dates:
        return None
    latest = max(dates)
    return (today - latest).days


def _risk_level(score: int) -> str:
    if score <= 2:
        return "LOW"
    if score <= 6:
        return "MEDIUM"
    if score <= 10:
        return "HIGH"
    return "CRITICAL"
