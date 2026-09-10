from datetime import date, time

from scripts.build_schedule import detect_conflicts, normalize_date, normalize_time


def test_normalize_date():
    assert normalize_date("2026-09-10") == date(2026, 9, 10)
    assert normalize_date("bad") is None


def test_normalize_time():
    assert normalize_time("09:30") == time(9, 30)
    assert normalize_time("25:00") is None


def test_detect_conflict():
    records = [
        {
            "date": date(2026, 9, 16), "person": "A", "project": "P1", "task": "Task1",
            "start": time(10, 0), "end": time(16, 0), "status": "Planned", "all_day": False,
        },
        {
            "date": date(2026, 9, 16), "person": "A", "project": "P2", "task": "Task2",
            "start": time(15, 0), "end": time(17, 0), "status": "Planned", "all_day": False,
        },
    ]
    conflicts = detect_conflicts(records)
    assert len(conflicts) == 1
    assert conflicts[0][-1] == "TIME_OVERLAP"


def test_no_conflict_for_different_person():
    records = [
        {
            "date": date(2026, 9, 16), "person": "A", "project": "P1", "task": "Task1",
            "start": time(10, 0), "end": time(16, 0), "status": "Planned", "all_day": False,
        },
        {
            "date": date(2026, 9, 16), "person": "B", "project": "P2", "task": "Task2",
            "start": time(15, 0), "end": time(17, 0), "status": "Planned", "all_day": False,
        },
    ]
    assert detect_conflicts(records) == []
