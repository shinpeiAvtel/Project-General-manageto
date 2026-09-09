from scripts.normalize_schedule import normalize_record


def test_normalization_sets_expected_fields():
    row = {
        "Project ID": "P1",
        "Project Name": "NRT82",
        "Site": "NRT",
        "Project Manager": "Shiraishi",
        "Task ID": "T1",
        "Task Name": "Kickoff",
        "Category": "NEA",
        "Start Date": "2026-09-01",
        "End Date": "2026-09-03",
        "Progress %": "50",
        "Status": "IN_PROGRESS",
        "Resource Count": "2",
        "Assigned Person": "A, B",
        "Milestone": "TRUE",
        "Dependency": "",
        "Priority": "HIGH",
        "Last Updated": "2026-09-01",
        "Remarks": "-",
    }
    out = normalize_record(row)
    assert out["project_name"] == "NRT82"
    assert out["assigned_person"] == ["A", "B"]
    assert out["milestone"] is True
    assert out["progress"] == 50.0
