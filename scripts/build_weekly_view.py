from __future__ import annotations

from collections import defaultdict
from datetime import date, timedelta

from .utils import apply_table_style, replace_sheet, style_header_row, summarise_record, write_banner


WEEKDAY_NAMES = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]


def build_weekly_view_sheet(workbook, records: list[dict], base_date: date, title: str = "02_Weekly_View"):
    worksheet = replace_sheet(workbook, title, 2)
    persons = sorted({record["person"] for record in records if record.get("person")})
    grouped = defaultdict(list)
    for record in records:
        grouped[(record.get("person"), record.get("date"))].append(record)

    headers = ["Week", "Person", *WEEKDAY_NAMES]
    write_banner(worksheet, len(headers))
    for column_index, header in enumerate(headers, start=1):
        worksheet.cell(row=2, column=column_index, value=header)

    week_monday = base_date - timedelta(days=base_date.weekday())
    week_starts = [week_monday, week_monday + timedelta(days=7)]

    row_index = 3
    for week_start in week_starts:
        week_label = f"W{week_start.isocalendar().week} ({week_start.isoformat()})"
        for person in persons:
            worksheet.cell(row=row_index, column=1, value=week_label)
            worksheet.cell(row=row_index, column=2, value=person)
            for offset in range(7):
                target_date = week_start + timedelta(days=offset)
                entries = sorted(grouped.get((person, target_date), []), key=lambda record: (record.get("start_time") is None, record.get("start_time")))
                worksheet.cell(
                    row=row_index,
                    column=3 + offset,
                    value="\n\n".join(summarise_record(record) for record in entries),
                )
            row_index += 1

    style_header_row(worksheet, 2)
    apply_table_style(worksheet, header_row=2, freeze_panes="C3")
    return worksheet


def main() -> None:
    from .build_master_schedule import build_output_artifacts

    build_output_artifacts()
    print("Weekly view updated.")


if __name__ == "__main__":
    main()
