from __future__ import annotations

import json
from pathlib import Path

from openpyxl import Workbook

try:
    from scripts.utils import ensure_dir
except ModuleNotFoundError:  # pragma: no cover
    from utils import ensure_dir


def _write_sheet(ws, rows):
    if not rows:
        ws.append(["NO DATA"])
        return
    headers = list(rows[0].keys())
    ws.append(headers)
    for row in rows:
        values = []
        for h in headers:
            v = row.get(h, "")
            if isinstance(v, (list, dict)):
                v = str(v)
            values.append(v)
        ws.append(values)


def generate_master_schedule(analysis: dict, validation_issues: list[dict], output_file: str | Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.title = "Master Schedule"
    _write_sheet(ws, analysis.get("master_schedule", []))

    sheet_map = [
        ("Project Summary", analysis.get("project_summary", [])),
        ("Milestones", analysis.get("milestones", [])),
        ("Delayed Tasks", analysis.get("delayed_tasks", [])),
        ("Resource Conflicts", analysis.get("resource_conflicts", [])),
        ("Upcoming 14 Days", analysis.get("upcoming_14_days", [])),
        ("Validation Errors", validation_issues),
    ]
    for name, rows in sheet_map:
        _write_sheet(wb.create_sheet(name), rows)

    output_file = Path(output_file)
    ensure_dir(output_file.parent)
    wb.save(output_file)


def main() -> None:
    repo_root = Path(__file__).resolve().parents[1]
    analysis = json.loads((repo_root / "data" / "master" / "analysis.json").read_text(encoding="utf-8"))
    validation = json.loads((repo_root / "data" / "master" / "validation_report.json").read_text(encoding="utf-8"))
    output_file = repo_root / "Output" / "Master_Schedule" / "Master_Schedule.xlsx"
    generate_master_schedule(analysis, validation, output_file)
    print(f"Generated {output_file}")


if __name__ == "__main__":
    main()
