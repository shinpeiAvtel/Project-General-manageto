from datetime import date, timedelta
from pathlib import Path

from scripts.import_excel import import_all
from scripts.validate_schedule import validate_imported
from scripts.normalize_schedule import normalize_rows
from scripts.analyze_schedule import analyze
from scripts.generate_master_schedule import generate_master_workbook
from tests.conftest import create_schedule_xlsx


def test_analysis_and_master_generation(tmp_path: Path):
    today = date.today()
    rows_a = [
        {
            "Project ID": "P1",
            "Project Name": "NRT82",
            "Site": "NRT",
            "Project Manager": "Shiraishi",
            "Task ID": "T1",
            "Task Name": "Late Task",
            "Category": "Testing",
            "Start Date": (today - timedelta(days=10)).isoformat(),
            "End Date": (today - timedelta(days=1)).isoformat(),
            "Progress %": 50,
            "Status": "IN_PROGRESS",
            "Resource Count": 1,
            "Assigned Person": "Shiraishi",
            "Milestone": True,
            "Dependency": "",
            "Priority": "HIGH",
            "Last Updated": (today - timedelta(days=20)).isoformat(),
            "Remarks": "",
        }
    ]
    rows_b = [
        {
            "Project ID": "P2",
            "Project Name": "HND10",
            "Site": "HND",
            "Project Manager": "Tanaka",
            "Task ID": "X1",
            "Task Name": "Upcoming",
            "Category": "T&C",
            "Start Date": (today - timedelta(days=1)).isoformat(),
            "End Date": (today + timedelta(days=2)).isoformat(),
            "Progress %": 0,
            "Status": "NOT_STARTED",
            "Resource Count": 1,
            "Assigned Person": "Shiraishi",
            "Milestone": False,
            "Dependency": "",
            "Priority": "HIGH",
            "Last Updated": today.isoformat(),
            "Remarks": "",
        }
    ]

    input_a = tmp_path / "Input" / "Shiraishi" / "NRT82"
    input_b = tmp_path / "Input" / "Tanaka" / "HND10"
    input_a.mkdir(parents=True)
    input_b.mkdir(parents=True)
    create_schedule_xlsx(input_a / "a.xlsx", rows_a)
    create_schedule_xlsx(input_b / "b.xlsx", rows_b)

    imported = import_all(tmp_path / "Input")
    config_dir = Path("/home/runner/work/Project-General-manageto/Project-General-manageto/config")
    validation = validate_imported(imported, config_dir)
    normalized = normalize_rows(imported)
    analysis = analyze(normalized, validation, config_dir)

    assert len(analysis["delayed_tasks"]) >= 1
    assert len(analysis["resource_conflicts"]) >= 1
    assert any(p["risk_score"] > 0 for p in analysis["project_summary"])

    out_path = generate_master_workbook(normalized, analysis, validation, tmp_path / "Output" / "Master_Schedule")
    assert out_path.exists()
