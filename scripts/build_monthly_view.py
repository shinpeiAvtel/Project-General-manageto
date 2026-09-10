from __future__ import annotations

from collections import defaultdict

from .utils import apply_table_style, style_header_row, summarise_record, write_banner


def build_monthly_view_sheet(workbook, records: list[dict], title: str = "01_Monthly_View"):
    worksheet = workbook.create_sheet(title)
    persons = sorted({record["person"] for record in records if record.get("person")})
    dates = sorted({record["date"] for record in records if record.get("date")})
    write_banner(worksheet, max(2, len(dates) + 1))
    worksheet.cell(row=2, column=1, value="Person")
    for offset, target_date in enumerate(dates, start=2):
        header_cell = worksheet.cell(row=2, column=offset, value=target_date)
        header_cell.number_format = "mmm dd"

    grouped = defaultdict(list)
    for record in records:
        grouped[(record.get("person"), record.get("date"))].append(record)

    for row_index, person in enumerate(persons, start=3):
        worksheet.cell(row=row_index, column=1, value=person)
        for col_index, target_date in enumerate(dates, start=2):
            entries = sorted(grouped.get((person, target_date), []), key=lambda record: (record.get("start_time") is None, record.get("start_time")))
            worksheet.cell(
                row=row_index,
                column=col_index,
                value="\n\n".join(summarise_record(record) for record in entries),
            )

    style_header_row(worksheet, 2)
    apply_table_style(worksheet, header_row=2, freeze_panes="B3")
    return worksheet


def main() -> None:
    from .build_master_schedule import build_output_artifacts

    build_output_artifacts()
    print("Monthly view updated.")


if __name__ == "__main__":
    main()
