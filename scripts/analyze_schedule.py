from __future__ import annotations

import json
from collections import defaultdict
from datetime import date, timedelta
from pathlib import Path
from typing import Any

try:
    from scripts.utils import MISSING, load_yaml, parse_date
except ModuleNotFoundError:  # pragma: no cover
    from utils import MISSING, load_yaml, parse_date


def _progress_value(progress: Any) -> float | None:
    if isinstance(progress, (int, float)):
        return float(progress)
    return None


def _risk_level(score: int, thresholds: dict[str, int]) -> str:
    if score <= thresholds.get("low_max", 2):
        return "LOW"
    if score <= thresholds.get("medium_max", 6):
        return "MEDIUM"
    if score <= thresholds.get("high_max", 10):
        return "HIGH"
    return "CRITICAL"


def analyze_records(
    records: list[dict[str, Any]],
    validation_issues: list[dict[str, Any]],
    rules: dict[str, Any],
    today: date | None = None,
) -> dict[str, Any]:
    today = today or date.today()
    milestone_categories = {c.lower() for c in rules.get("milestone_categories", [])}
    weights = rules.get("risk_weights", {})
    stale_cfg = rules.get("stale_update_threshold_days", {"warning": 7, "high_warning": 14})

    master_rows: list[dict[str, Any]] = []
    delayed_tasks: list[dict[str, Any]] = []
    milestones: list[dict[str, Any]] = []
    milestone_candidates: list[dict[str, Any]] = []
    upcoming: list[dict[str, Any]] = []

    for r in records:
        start = parse_date(r.get("start_date"))
        end = parse_date(r.get("end_date"))
        progress = _progress_value(r.get("progress"))
        delayed_status = "ON_TRACK"
        days_remaining = None

        if end:
            days_remaining = (end - today).days
        if end and progress is not None and today > end and progress < 100:
            delayed_status = "DELAYED"
        elif start and progress == 0 and today > start:
            delayed_status = "NOT_STARTED_LATE"

        row = dict(r)
        row["delayed_status"] = delayed_status
        row["days_remaining"] = days_remaining if days_remaining is not None else MISSING
        master_rows.append(row)

        if delayed_status in {"DELAYED", "NOT_STARTED_LATE"}:
            delayed_tasks.append(row)
        if r.get("milestone") is True:
            milestones.append(row)
        if r.get("milestone") is False and str(r.get("category", "")).lower() in milestone_categories:
            cand = dict(row)
            cand["tag"] = "MILESTONE_CANDIDATE"
            milestone_candidates.append(cand)

        window_end = today + timedelta(days=14)
        for key in ("start_date", "end_date"):
            d = parse_date(r.get(key))
            if d and today <= d <= window_end:
                upcoming.append(row)
                break

    # project summary
    project_groups: dict[str, list[dict[str, Any]]] = defaultdict(list)
    for row in master_rows:
        project_groups[str(row.get("project_name", MISSING))].append(row)

    stale_projects: dict[str, str] = {}
    project_summary: list[dict[str, Any]] = []
    for project_name, rows in project_groups.items():
        starts = [parse_date(r.get("start_date")) for r in rows if parse_date(r.get("start_date"))]
        ends = [parse_date(r.get("end_date")) for r in rows if parse_date(r.get("end_date"))]
        completed = sum(1 for r in rows if _progress_value(r.get("progress")) == 100)
        delayed = sum(1 for r in rows if r.get("delayed_status") == "DELAYED")
        ms_count = sum(1 for r in rows if r.get("milestone") is True)
        progress_vals = [_progress_value(r.get("progress")) for r in rows if _progress_value(r.get("progress")) is not None]
        avg_progress = round(sum(progress_vals) / len(progress_vals), 1) if progress_vals else 0.0

        latest = None
        for r in rows:
            d = parse_date(r.get("last_updated"))
            if d and (latest is None or d > latest):
                latest = d
        stale_state = ""
        if latest is not None:
            age = (today - latest).days
            if age >= stale_cfg.get("high_warning", 14):
                stale_state = "HIGH WARNING"
            elif age >= stale_cfg.get("warning", 7):
                stale_state = "WARNING"
        if stale_state:
            stale_projects[project_name] = stale_state

        project_summary.append(
            {
                "project_name": project_name,
                "project_manager": rows[0].get("project_manager", MISSING),
                "start_date": min(starts).isoformat() if starts else MISSING,
                "end_date": max(ends).isoformat() if ends else MISSING,
                "total_tasks": len(rows),
                "completed_tasks": completed,
                "progress": avg_progress,
                "delayed_tasks": delayed,
                "milestones": ms_count,
                "last_updated": latest.isoformat() if latest else MISSING,
            }
        )

    # resource conflicts
    assignments: list[dict[str, Any]] = []
    for row in master_rows:
        start = parse_date(row.get("start_date"))
        end = parse_date(row.get("end_date"))
        if not start or not end:
            continue
        persons = set(row.get("assigned_person", []))
        pm = row.get("project_manager")
        if isinstance(pm, str) and pm and pm != MISSING:
            persons.add(pm)
        for person in persons:
            assignments.append({"person": person, "project_name": row.get("project_name"), "task_id": row.get("task_id"), "start": start, "end": end})

    conflicts: list[dict[str, Any]] = []
    for i, a in enumerate(assignments):
        for b in assignments[i + 1 :]:
            if a["person"] != b["person"]:
                continue
            if a["project_name"] == b["project_name"]:
                continue
            overlap_start = max(a["start"], b["start"])
            overlap_end = min(a["end"], b["end"])
            if overlap_start <= overlap_end:
                conflicts.append(
                    {
                        "person": a["person"],
                        "project_a": a["project_name"],
                        "project_b": b["project_name"],
                        "overlap_start": overlap_start.isoformat(),
                        "overlap_end": overlap_end.isoformat(),
                        "type": "POTENTIAL_CONFLICT",
                    }
                )

    # risk score
    err_count: dict[str, int] = defaultdict(int)
    for e in validation_issues:
        if e.get("severity") == "ERROR":
            err_count[str(e.get("project_name", MISSING))] += 1

    conflict_projects = defaultdict(int)
    for c in conflicts:
        conflict_projects[c["project_a"]] += 1
        conflict_projects[c["project_b"]] += 1

    risk_by_project: dict[str, dict[str, Any]] = {}
    thresholds = rules.get("risk_level_thresholds", {})
    for summary in project_summary:
        pname = summary["project_name"]
        rows = project_groups[pname]
        score = 0
        for row in rows:
            if row.get("delayed_status") == "DELAYED":
                score += weights.get("delayed_task", 3)
                if str(row.get("priority", "")).upper() == "HIGH":
                    score += weights.get("high_priority_delayed_task", 5)
                if row.get("milestone") is True:
                    score += weights.get("milestone_delay", 5)
            if row.get("delayed_status") == "NOT_STARTED_LATE":
                score += weights.get("not_started_late", 2)
        if pname in stale_projects:
            if stale_projects[pname] == "HIGH WARNING":
                score += weights.get("stale_high_warning", 4)
            else:
                score += weights.get("stale_warning", 2)
        score += conflict_projects[pname] * weights.get("resource_conflict", 2)
        score += err_count[pname] * weights.get("validation_error", 2)
        risk_by_project[pname] = {"score": score, "level": _risk_level(score, thresholds)}

    for summary in project_summary:
        risk = risk_by_project.get(summary["project_name"], {"score": 0, "level": "LOW"})
        summary["risk_score"] = risk["score"]
        summary["risk_level"] = risk["level"]

    return {
        "generated_at": today.isoformat(),
        "master_schedule": master_rows,
        "project_summary": project_summary,
        "milestones": milestones,
        "milestone_candidates": milestone_candidates,
        "delayed_tasks": delayed_tasks,
        "resource_conflicts": conflicts,
        "upcoming_14_days": upcoming,
        "stale_projects": stale_projects,
        "risk_by_project": risk_by_project,
    }


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    records = json.loads((repo_root / "data" / "normalized" / "normalized_schedule.json").read_text(encoding="utf-8"))
    validation = json.loads((repo_root / "data" / "master" / "validation_report.json").read_text(encoding="utf-8"))
    rules = load_yaml(repo_root / "config" / "project_rules.yaml")
    analysis = analyze_records(records, validation, rules)
    (repo_root / "data" / "master" / "analysis.json").write_text(json.dumps(analysis, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Analysis completed")


if __name__ == "__main__":
    main()
