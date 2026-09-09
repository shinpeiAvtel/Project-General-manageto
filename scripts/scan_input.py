from __future__ import annotations

import json
from pathlib import Path

from scripts.import_excel import import_all
from scripts.validate_schedule import validate_imported
from scripts.normalize_schedule import normalize_rows, save_normalized
from scripts.analyze_schedule import analyze
from scripts.generate_master_schedule import generate_master_workbook
from scripts.generate_reports import generate_reports
from scripts.utils import ensure_dir


def run_pipeline(root: Path) -> dict:
    imported = import_all(root / "Input")
    validation = validate_imported(imported, root / "config")
    normalized = normalize_rows(imported)
    save_normalized(normalized, root / "data" / "normalized")

    analysis = analyze(normalized, validation, root / "config")

    ensure_dir(root / "data" / "master")
    with (root / "data" / "master" / "analysis.json").open("w", encoding="utf-8") as f:
        json.dump(analysis, f, ensure_ascii=False, indent=2)

    master_path = generate_master_workbook(normalized, analysis, validation, root / "Output" / "Master_Schedule")
    generate_reports(analysis, validation, root / "Output" / "Reports", data_source=str(root / "Input"))

    return {
        "imported_files": imported["files"],
        "imported_rows": len(imported["rows"]),
        "validation_issues": len(validation["issues"]),
        "master_schedule": str(master_path),
    }


if __name__ == "__main__":
    repo_root = Path(__file__).resolve().parents[1]
    result = run_pipeline(repo_root)
    print(json.dumps(result, ensure_ascii=False, indent=2))
