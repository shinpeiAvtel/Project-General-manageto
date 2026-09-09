from pathlib import Path

from scripts.import_excel import import_all
from tests.conftest import create_schedule_xlsx, load_fixture_rows


def test_import_excel_success(tmp_path: Path):
    input_dir = tmp_path / "Input" / "Shiraishi" / "NRT82"
    input_dir.mkdir(parents=True)
    create_schedule_xlsx(input_dir / "NRT82_schedule.xlsx", load_fixture_rows())

    result = import_all(tmp_path / "Input")
    assert len(result["files"]) == 1
    assert len(result["rows"]) == 2
    assert result["rows"][0]["data"]["Project Name"] == "NRT82"
