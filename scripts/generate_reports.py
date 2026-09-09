from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path
from openpyxl import Workbook

from scripts.utils import ensure_dir


def generate_reports(analysis: dict, validation: dict, output_reports_dir: Path, data_source: str) -> None:
    ensure_dir(output_reports_dir)

    wb = Workbook()
    ws = wb.active
    ws.title = "Upcoming 14 Days"
    ws.append(["Project Name", "Task ID", "Task Name", "Start Date", "End Date", "Priority", "Delay Status", "Category", "Milestone"])
    for row in analysis.get("upcoming_14_days", []):
        ws.append([
            row.get("project_name"),
            row.get("task_id"),
            row.get("task_name"),
            row.get("start_date"),
            row.get("end_date"),
            row.get("priority"),
            row.get("delay_status"),
            row.get("category"),
            row.get("milestone"),
        ])
    wb.save(output_reports_dir / "Upcoming_14_Days.xlsx")

    with (output_reports_dir / "Validation_Report.json").open("w", encoding="utf-8") as f:
        json.dump(validation, f, ensure_ascii=False, indent=2)

    _generate_manager_summary(analysis, output_reports_dir / "Manager_Summary.md", data_source)


def _generate_manager_summary(analysis: dict, out_path: Path, data_source: str) -> None:
    project_summary = analysis.get("project_summary", [])
    delayed = analysis.get("delayed_tasks", [])
    conflicts = analysis.get("resource_conflicts", [])
    milestones = analysis.get("milestones", [])
    upcoming = analysis.get("upcoming_14_days", [])

    total_projects = len(project_summary)
    active_projects = len([p for p in project_summary if p.get("progress", 0) < 100])
    delayed_projects = len([p for p in project_summary if p.get("delayed_tasks", 0) > 0])
    critical_projects = len([p for p in project_summary if p.get("risk_level") == "CRITICAL"])

    stale_projects = [p for p in project_summary if p.get("update_status") in {"WARNING", "HIGH WARNING"}]

    lines = [
        "# Project General Management Summary",
        "",
        f"Generated: {datetime.now(timezone.utc).replace(microsecond=0).isoformat().replace('+00:00', 'Z')}",
        f"Data Source: {data_source}",
        "",
        "## Overall Status",
        f"Total Projects: {total_projects}",
        f"Active Projects: {active_projects}",
        f"Delayed Projects: {delayed_projects}",
        f"Critical Projects: {critical_projects}",
        "",
        "## Critical Issues",
    ]
    if critical_projects == 0:
        lines.append("- None")
    else:
        for p in [x for x in project_summary if x.get("risk_level") == "CRITICAL"]:
            lines.append(f"- {p['project_name']}: Risk {p['risk_level']} (Score {p['risk_score']})")

    lines += ["", "## Upcoming 14 Days"]
    if not upcoming:
        lines.append("- None")
    else:
        for r in upcoming[:20]:
            lines.append(f"- {r['project_name']} / {r['task_name']} ({r['start_date']} - {r['end_date']})")

    lines += ["", "## Delayed Tasks"]
    if not delayed:
        lines.append("- None")
    else:
        for r in delayed[:20]:
            lines.append(f"- {r['project_name']} / {r['task_name']} End:{r['end_date']} Progress:{r['progress']}")

    lines += ["", "## Resource Conflicts"]
    if not conflicts:
        lines.append("- None")
    else:
        for c in conflicts[:20]:
            lines.append(f"- {c['person']} on {c['date']}: {', '.join(c['projects'])} ({c['status']})")

    lines += ["", "## Major Milestones"]
    if not milestones:
        lines.append("- None")
    else:
        for m in milestones[:20]:
            lines.append(f"- {m['project_name']} / {m['task_name']} End:{m['end_date']}")

    lines += ["", "## Projects Requiring Update"]
    if not stale_projects:
        lines.append("- None")
    else:
        for p in stale_projects:
            lines.append(f"- {p['project_name']}: {p['update_status']} ({p['last_updated_days_ago']} days)")

    lines += ["", "## Recommended Attention"]
    if critical_projects == 0 and not delayed and not conflicts:
        lines.append("- Current data indicates no immediate critical action is required.")
    else:
        for p in project_summary[:10]:
            if p.get("risk_score", 0) >= 7:
                lines.append(
                    f"- {p['project_name']}: Focus on delayed tasks ({p['delayed_tasks']}), conflicts ({p['resource_conflict_count']}), and update freshness ({p['update_status']})."
                )

    out_path.write_text("\n".join(lines) + "\n", encoding="utf-8")
