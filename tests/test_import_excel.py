from pathlib import Path

from scripts.import_excel import ExcelImporter


def test_import_excel_reads_fixture():
    repo = Path(__file__).resolve().parents[1]
    importer = ExcelImporter(repo / "config" / "schedule_schema.yaml")
    fixture = repo / "tests" / "fixtures" / "sample_schedule.xlsx"
    records, issues = importer.import_file(fixture, repo / "tests" / "fixtures")

    assert len(records) >= 1
    assert issues == []
    assert records[0]["Project Name"] == "NRT82"
