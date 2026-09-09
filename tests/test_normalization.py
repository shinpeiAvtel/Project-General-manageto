from pathlib import Path

from scripts.import_excel import import_all
from scripts.normalize_schedule import normalize_rows
from tests.conftest import create_schedule_xlsx, load_fixture_rows


def test_normalization(tmp_path: Path):
    input_dir = tmp_path / "Input" / "Shiraishi" / "NRT82"
    input_dir.mkdir(parents=True)
    create_schedule_xlsx(input_dir / "sample.xlsx", load_fixture_rows())

    imported = import_all(tmp_path / "Input")
    normalized = normalize_rows(imported)

    assert normalized[0]["project_name"] == "NRT82"
    assert isinstance(normalized[1]["assigned_person"], list)
