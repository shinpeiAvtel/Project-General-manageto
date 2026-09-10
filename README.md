# Project General manage

Personal Schedule → Master Schedule aggregation system.

**Baseline date:** 2026-09-10

## Purpose

Each staff member manages their own schedule on a dedicated worksheet inside one Excel workbook. The system reads all personal worksheets and automatically creates a manager-facing master workbook.

## Source of truth

`Input/Personal_Schedule.xlsx`

Personal worksheet names are treated as person names. System sheets are excluded automatically.

Required columns:

1. Date
2. Project
3. Site
4. Schedule / Task
5. Start Time
6. End Time
7. Status
8. Remarks
9. Category
10. Priority
11. All Day
12. Remote / Site / Office
13. Last Updated

## Generated output

`Output/Master_Schedule.xlsx`

Generated sheets:

- `00_Master_Schedule` — all personal schedules merged into one list
- `01_Monthly_View` — person × date manager view
- `02_Weekly_View` — current and next week schedule list
- `03_Conflict_Report` — overlapping schedules for the same person
- `04_Availability` — AVAILABLE / PARTIAL / BOOKED / LEAVE
- `05_Validation_Report` — input errors and warnings

## Initial setup

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
python scripts/build_schedule.py
```

On Linux/macOS use `source .venv/bin/activate`.

If `Input/Personal_Schedule.xlsx` does not exist, the script creates a sample workbook containing three personal sheets: Shiraishi, Tanaka, and Suzuki.

## Operating rule

Edit only personal worksheets in the input workbook. Generated master/output sheets must not be used as source data.

Do not infer missing dates, projects, people, times, or availability. Missing source values are treated as unknown.

## GitHub Actions

`.github/workflows/update-master-schedule.yml` runs when schedule input, scripts, configuration, tests, or requirements change. It:

1. checks out the repository,
2. installs Python dependencies,
3. validates and aggregates schedules,
4. generates the master workbook,
5. runs pytest,
6. uploads `Master_Schedule.xlsx` as a GitHub Actions artifact.

## Workday configuration

Default workday is `09:00-18:00` and can be changed in `config/schedule_config.yaml`.

Availability is calculated only against configured work hours; the system does not guess working hours.

## Manager questions for Copilot

Examples:

- 今日の全員の予定を教えてください
- 9月15日に空いている人を教えてください
- Project Aに入る担当者を教えてください
- 今週のShiraishiの予定を教えてください
- 予定が重複している担当者を教えてください
- 今月の全体予定をまとめてください

Repository instructions are defined in `.github/copilot-instructions.md`.

## Current limitation

The GitHub repository is still named `Project-General-manageto`. The intended display/project name is `Project General manage`. Rename the repository in GitHub Settings when convenient; the system files themselves already use the new project name.
