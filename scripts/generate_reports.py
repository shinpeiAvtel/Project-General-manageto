from __future__ import annotations

import json
from datetime import datetime, timezone
from pathlib import Path

from openpyxl import Workbook

try:
    from scripts.utils import ensure_dir
except ModuleNotFoundError:  # pragma: no cover
    from utils import ensure_dir


def _write_upcoming(upcoming_rows: list[dict], out_file: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Upcoming 14 Days"
    if not upcoming_rows:
        ws.append(["NO DATA"])
    else:
        headers = list(upcoming_rows[0].keys())
        ws.append(headers)
        for row in upcoming_rows:
            vals = []
            for h in headers:
                v = row.get(h, "")
                if isinstance(v, (list, dict)):
                    v = str(v)
                vals.append(v)
            ws.append(vals)
    wb.save(out_file)


def _manager_summary_md(analysis: dict, validation_issues: list[dict]) -> str:
    project_summary = analysis.get("project_summary", [])
    delayed = analysis.get("delayed_tasks", [])
    conflicts = analysis.get("resource_conflicts", [])
    upcoming = analysis.get("upcoming_14_days", [])
    critical_projects = [p for p in project_summary if p.get("risk_level") == "CRITICAL"]
    stale_projects = analysis.get("stale_projects", {})

    lines = [
        "# Project General Management Summary",
        "",
        f"Generated: {datetime.now(timezone.utc).isoformat()}",
        "Data Source: Input/**/*.xlsx, data/normalized/normalized_schedule.json",
        "",
        "## Overall Status",
        f"Total Projects: {len(project_summary)}",
        f"Active Projects: {len([p for p in project_summary if p.get('progress', 0) < 100])}",
        f"Delayed Projects: {len({d.get('project_name') for d in delayed})}",
        f"Critical Projects: {len(critical_projects)}",
        "",
        "## Critical Issues",
    ]
    if critical_projects:
        for p in critical_projects:
            lines.append(f"- {p['project_name']}: Risk {p['risk_level']} (Score {p['risk_score']})")
    else:
        lines.append("- なし")

    lines += ["", "## Upcoming 14 Days"]
    lines.append(f"- 件数: {len(upcoming)}")

    lines += ["", "## Delayed Tasks", f"- 件数: {len(delayed)}"]
    lines += ["", "## Resource Conflicts", f"- 件数: {len(conflicts)}"]
    lines += ["", "## Major Milestones", f"- 件数: {len(analysis.get('milestones', []))}"]
    lines += ["", "## Projects Requiring Update"]
    if stale_projects:
        for p, level in stale_projects.items():
            lines.append(f"- {p}: {level}")
    else:
        lines.append("- なし")

    lines += ["", "## Recommended Attention"]
    lines.append("- Validation Errorのあるプロジェクトを優先確認してください。")
    lines.append("- DELAYEDおよびNOT_STARTED_LATEタスクを優先的に是正してください。")
    lines.append("- POTENTIAL_CONFLICTは担当者調整の要確認事項です。")
    lines.append(f"- Validation Error件数: {len([e for e in validation_issues if e.get('severity') == 'ERROR'])}")

    return "\n".join(lines) + "\n"


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    analysis = json.loads((repo_root / "data" / "master" / "analysis.json").read_text(encoding="utf-8"))
    validation = json.loads((repo_root / "data" / "master" / "validation_report.json").read_text(encoding="utf-8"))

    out_dir = repo_root / "Output" / "Reports"
    ensure_dir(out_dir)

    _write_upcoming(analysis.get("upcoming_14_days", []), out_dir / "Upcoming_14_Days.xlsx")
    (out_dir / "Manager_Summary.md").write_text(_manager_summary_md(analysis, validation), encoding="utf-8")
    (out_dir / "Validation_Report.json").write_text(json.dumps(validation, ensure_ascii=False, indent=2), encoding="utf-8")
    print("Reports generated")


if __name__ == "__main__":
    main()
