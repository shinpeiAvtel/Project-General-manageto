from __future__ import annotations

from pathlib import Path
from typing import Any
from openpyxl import load_workbook

from scripts.utils import now_iso, parse_date, parse_int, parse_progress


def find_excel_files(input_dir: Path) -> list[Path]:
    return sorted([p for p in input_dir.rglob("*.xlsx") if not p.name.startswith("~$")])


def _row_to_dict(headers: list[str], row_values: list[Any]) -> dict[str, Any]:
    return {headers[idx]: row_values[idx] if idx < len(row_values) else None for idx in range(len(headers))}


def import_workbook(path: Path) -> list[dict[str, Any]]:
    wb = load_workbook(path, data_only=True, read_only=True)
    imported: list[dict[str, Any]] = []
    for ws in wb.worksheets:
        rows = ws.iter_rows(values_only=True)
        try:
            headers = [str(h).strip() if h is not None else "" for h in next(rows)]
        except StopIteration:
            continue
        for raw in rows:
            if raw is None or all(v is None or str(v).strip() == "" for v in raw):
                continue
            data = _row_to_dict(headers, list(raw))
            data["Start Date"] = parse_date(data.get("Start Date"))
            data["End Date"] = parse_date(data.get("End Date"))
            data["Last Updated"] = parse_date(data.get("Last Updated"))
            data["Progress %"] = parse_progress(data.get("Progress %"))
            data["Resource Count"] = parse_int(data.get("Resource Count"), default=0)
            imported.append(
                {
                    "data": data,
                    "metadata": {
                        "source_file": str(path),
                        "source_sheet": ws.title,
                        "headers": headers,
                        "imported_at": now_iso(),
                        "folder_owner": path.parent.parent.name if len(path.parts) >= 3 else "",
                        "folder_project": path.parent.name,
                    },
                }
            )
    return imported


def import_all(input_dir: Path) -> dict[str, Any]:
    files = find_excel_files(input_dir)
    rows: list[dict[str, Any]] = []
    for file_path in files:
        rows.extend(import_workbook(file_path))
    return {"files": [str(f) for f in files], "rows": rows, "imported_at": now_iso()}


if __name__ == "__main__":
    root = Path(__file__).resolve().parents[1]
    result = import_all(root / "Input")
    print(f"Imported files: {len(result['files'])}, rows: {len(result['rows'])}")
