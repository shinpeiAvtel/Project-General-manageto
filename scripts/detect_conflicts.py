from __future__ import annotations

from itertools import combinations
from typing import Any

from .utils import build_display_time, overlap_minutes, record_interval_minutes


def detect_conflicts(records: list[dict[str, Any]]) -> list[dict[str, Any]]:
    grouped: dict[tuple[Any, Any], list[dict[str, Any]]] = {}
    for record in records:
        grouped.setdefault((record.get("person"), record.get("date")), []).append(record)

    conflicts: list[dict[str, Any]] = []
    for (person, schedule_date), daily_records in grouped.items():
        for first, second in combinations(daily_records, 2):
            first_interval = record_interval_minutes(first)
            second_interval = record_interval_minutes(second)
            if not first_interval or not second_interval:
                continue
            if overlap_minutes(first_interval, second_interval) <= 0:
                continue
            conflicts.append(
                {
                    "person": person,
                    "date": schedule_date,
                    "schedule_a": first.get("schedule_task"),
                    "schedule_b": second.get("schedule_task"),
                    "project_a": first.get("project"),
                    "project_b": second.get("project"),
                    "time_a": build_display_time(first),
                    "time_b": build_display_time(second),
                    "conflict_type": "CONFLICT",
                    "source_sheet_a": first.get("source_sheet"),
                    "source_sheet_b": second.get("source_sheet"),
                    "source_row_a": first.get("source_row"),
                    "source_row_b": second.get("source_row"),
                }
            )
    return conflicts


def main() -> None:
    from .build_master_schedule import build_output_artifacts

    build_output_artifacts()
    print("Conflict report updated.")


if __name__ == "__main__":
    main()
