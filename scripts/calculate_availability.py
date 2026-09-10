from __future__ import annotations

from collections import defaultdict
from datetime import date

from .utils import daterange, overlap_minutes, parse_excel_date, record_interval_minutes, workday_minutes


def parse_config_date(value: str | None, field_name: str) -> date | None:
    if value in (None, ""):
        return None
    parsed = parse_excel_date(value)
    if not parsed:
        raise ValueError(f"Configured {field_name} must be a valid date value.")
    return parsed


def calculate_availability(records: list[dict], config: dict) -> list[dict]:
    if not records:
        return []

    persons = sorted({record["person"] for record in records if record.get("person")})
    dated_records = [record for record in records if record.get("date")]
    if not dated_records:
        return []

    configured_start = parse_config_date(config["availability"].get("start_date"), "availability.start_date")
    configured_end = parse_config_date(config["availability"].get("end_date"), "availability.end_date")
    start_date = configured_start or min(record["date"] for record in dated_records)
    end_date = configured_end or max(record["date"] for record in dated_records)
    work_start, work_end = workday_minutes(config)
    work_interval = (work_start, work_end)
    total_work_minutes = work_end - work_start

    grouped = defaultdict(list)
    for record in records:
        grouped[(record.get("person"), record.get("date"))].append(record)

    availability_rows = []
    for person in persons:
        for target_date in daterange(start_date, end_date):
            day_records = [
                record
                for record in grouped.get((person, target_date), [])
                if record.get("status") != "Cancelled"
            ]
            leave_records = [record for record in day_records if record.get("status") == "Leave"]
            if leave_records:
                status = "LEAVE"
                busy_minutes = total_work_minutes
                note = "Leave schedule present."
            else:
                intervals = []
                unbounded_schedule = False
                for record in day_records:
                    interval = record_interval_minutes(record)
                    if interval is None:
                        if record.get("schedule_task"):
                            unbounded_schedule = True
                        continue
                    intervals.append(interval)
                covered = []
                for start, end in sorted(intervals):
                    if not covered or start > covered[-1][1]:
                        covered.append([start, end])
                    else:
                        covered[-1][1] = max(covered[-1][1], end)
                busy_minutes = sum(overlap_minutes(work_interval, (start, end)) for start, end in covered)
                if any(record.get("all_day") for record in day_records):
                    status = "BOOKED"
                    busy_minutes = total_work_minutes
                    note = "All Day schedule present."
                elif busy_minutes == 0 and not day_records:
                    status = "AVAILABLE"
                    note = "No schedules within configured workday."
                elif busy_minutes >= total_work_minutes:
                    status = "BOOKED"
                    note = "Workday fully covered by schedules."
                else:
                    status = "PARTIAL" if (busy_minutes > 0 or unbounded_schedule or day_records) else "AVAILABLE"
                    note = "Schedules cover part of the workday." if busy_minutes > 0 else "Schedule exists without usable time range."
            availability_rows.append(
                {
                    "date": target_date,
                    "person": person,
                    "availability": status,
                    "scheduled_minutes": busy_minutes,
                    "working_minutes": total_work_minutes,
                    "notes": note,
                }
            )
    return availability_rows


def main() -> None:
    from .build_master_schedule import build_output_artifacts

    build_output_artifacts()
    print("Availability report updated.")


if __name__ == "__main__":
    main()
