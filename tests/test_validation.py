from scripts.validate_schedule import validate_records


def test_validation_detects_required_and_range_errors():
    schema = {
        "required_fields_for_validation": ["Project Name", "Task Name", "Start Date", "End Date"],
        "allowed_status": ["IN_PROGRESS", "COMPLETED"],
    }
    records = [
        {
            "Project Name": "NRT82",
            "Task Name": "",
            "Task ID": "T1",
            "Start Date": "2026-09-10",
            "End Date": "2026-09-09",
            "Progress %": 110,
            "Status": "BAD",
            "Last Updated": "",
            "_folder_project": "NRT82",
        },
        {
            "Project Name": "NRT82",
            "Task Name": "Task2",
            "Task ID": "T2",
            "Start Date": "2026-09-10",
            "End Date": "2026-09-11",
            "Progress %": -1,
            "Status": "IN_PROGRESS",
            "Last Updated": "2026-09-01",
            "Dependency": "X999",
            "_folder_project": "NRT82",
        },
    ]
    issues = validate_records(records, schema)
    codes = {i["code"] for i in issues}
    assert "MISSING_REQUIRED" in codes
    assert "START_AFTER_END" in codes
    assert "PROGRESS_GT_100" in codes
    assert "PROGRESS_LT_0" in codes
    assert "INVALID_STATUS" in codes
    assert "MISSING_LAST_UPDATED" in codes
    assert "MISSING_DEPENDENCY" in codes
