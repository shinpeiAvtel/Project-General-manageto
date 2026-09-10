# Project-General-manageto

Project-General-manageto is an Excel schedule automation system that reads personal schedule sheets from `Input/Personal_Schedule.xlsx`, validates and normalizes the records, and generates manager-facing views in `Output/Master_Schedule.xlsx`.

## Features

- Dynamically detects personal sheets without hard-coding names
- Excludes system sheets (`00_Master_Schedule`, `01_Monthly_View`, `02_Weekly_View`, `03_Conflict_Report`, `04_Availability`, `Template`, `Config`)
- Validates dates, times, statuses, missing fields, and duplicates without stopping other valid records
- Preserves `Source Sheet` and `Source Row` traceability in the master output
- Generates:
  - `00_Master_Schedule`
  - `01_Monthly_View`
  - `02_Weekly_View`
  - `03_Conflict_Report`
  - `04_Availability`
- Applies `openpyxl` formatting for manager readability
- Produces normalized data and validation reports for repository-based schedule queries

## Repository Structure

```text
Project-General-manageto/
├── Input/
│   └── Personal_Schedule.xlsx
├── Output/
│   └── Master_Schedule.xlsx
├── scripts/
│   ├── __init__.py
│   ├── read_personal_schedules.py
│   ├── validate_schedule.py
│   ├── build_master_schedule.py
│   ├── build_monthly_view.py
│   ├── build_weekly_view.py
│   ├── detect_conflicts.py
│   ├── calculate_availability.py
│   └── utils.py
├── config/
│   └── schedule_config.yaml
├── data/
│   ├── normalized_schedule.json
│   ├── validation_report.json
│   └── build_summary.json
├── docs/
│   ├── system-overview.md
│   ├── user-manual.md
│   └── manager-manual.md
├── tests/
├── .github/
│   ├── workflows/
│   │   └── update-master-schedule.yml
│   └── copilot-instructions.md
├── requirements.txt
└── README.md
```

## Input Workbook Rules

- Use `Input/Personal_Schedule.xlsx` as the source of truth.
- Use one sheet per person.
- Do not manually maintain the output workbook as source data.
- Required personal schedule columns:
  - `Date`
  - `Project`
  - `Site`
  - `Schedule / Task`
  - `Start Time`
  - `End Time`
  - `Status`
  - `Remarks`
- Supported optional columns:
  - `Category`
  - `Client`
  - `Location`
  - `Priority`
  - `All Day`
  - `Remote / Site / Office`
  - `Confirmed / Tentative`
  - `Last Updated`
- Date format: `YYYY-MM-DD`
- Time format: `HH:MM`
- Supported status values: `Planned`, `Confirmed`, `Completed`, `Cancelled`, `Leave`
- Unknown `Project` or `Site` values are normalized to `UNKNOWN`.

## Setup

```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Usage

Run validation and build the output workbook:

```bash
python -m scripts.validate_schedule
python -m scripts.build_master_schedule
```

The wrapper commands below are also available and will refresh the generated workbook:

```bash
python -m scripts.build_monthly_view
python -m scripts.build_weekly_view
python -m scripts.detect_conflicts
python -m scripts.calculate_availability
```

## Generated Data

- `data/normalized_schedule.json`: validated records for downstream analysis
- `data/validation_report.json`: `ERROR`, `WARNING`, `INFO` validation log
- `data/build_summary.json`: generation metadata

## Manager Query Guidance

When answering schedule questions from repository data, always provide:

- `Conclusion`
- `Reason`
- `Numbers`
- `Source`
- `Data Date`

Never infer schedules, people, projects, or availability beyond what exists in the workbook and configuration.

## Tests

```bash
pytest
```
