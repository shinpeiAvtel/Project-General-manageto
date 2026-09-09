from pathlib import Path

from scripts.import_excel import import_all
from scripts.validate_schedule import validate_imported
from tests.conftest import create_schedule_xlsx, load_fixture_rows


def test_validation_required_and_ranges(tmp_path: Path):
    rows = load_fixture_rows()
    rows[0]["Project Name"] = ""
    rows[0]["Progress %"] = 120
    rows[1]["Progress %"] = -1
    rows[1]["Start Date"] = "2026-02-01"
    rows[1]["End Date"] = "2026-01-01"

    input_dir = tmp_path / "Input" / "Shiraishi" / "NRT82"
    input_dir.mkdir(parents=True)
    create_schedule_xlsx(input_dir / "bad.xlsx", rows)

    imported = import_all(tmp_path / "Input")
    report = validate_imported(imported, Path("/home/runner/work/Project-General-manageto/Project-General-manageto/config"))

    codes = [i["code"] for i in report["issues"]]
    assert "MISSING_REQUIRED" in codes
    assert "PROGRESS_RANGE" in codes
    assert "DATE_ORDER" in codes
