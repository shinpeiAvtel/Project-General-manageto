import json
from pathlib import Path
from openpyxl import Workbook

HEADERS = [
    "Project ID","Project Name","Site","Project Manager","Task ID","Task Name","Category","Start Date","End Date",
    "Progress %","Status","Resource Count","Assigned Person","Milestone","Dependency","Priority","Last Updated","Remarks"
]


def create_schedule_xlsx(path: Path, rows: list[dict]):
    wb = Workbook()
    ws = wb.active
    ws.title = "Schedule"
    ws.append(HEADERS)
    for r in rows:
        ws.append([r.get(h, "") for h in HEADERS])
    wb.save(path)


def load_fixture_rows() -> list[dict]:
    fixture = Path(__file__).parent / "fixtures" / "sample_schedule_rows.json"
    return json.loads(fixture.read_text(encoding="utf-8"))
