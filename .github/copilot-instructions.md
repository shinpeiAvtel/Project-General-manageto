# Project General manage - Copilot Instructions

Date baseline: 2026-09-10

## Source of truth
- Individual person sheets inside `Input/Personal_Schedule.xlsx` are the source of truth.
- `Output/Master_Schedule.xlsx` is generated output and must not be treated as manually editable source data.
- Do not invent schedules, people, projects, dates, times, status, availability, or project assignments.
- When source data is missing, report `UNKNOWN`, `MISSING`, or `NOT PROVIDED`.

## Required answer format for schedule questions
1. Conclusion
2. Reason
3. Numbers
4. Source
5. Data Date

## Rules
- Prefer deterministic results from `scripts/build_schedule.py` over LLM inference.
- Always state the relevant date explicitly.
- Identify the source sheet and source row when available.
- A detected time overlap means only `TIME_OVERLAP`; never conclude that the person cannot handle both appointments unless source data states that.
- Do not infer availability outside the configured workday in `config/schedule_config.yaml`.
- Do not mix data across people or projects.
- Cancelled schedules are excluded from conflict and availability calculations.
- Do not overwrite individual person sheets from generated outputs.

## Common manager questions
- Today's schedule for all staff
- Tomorrow's schedule for all staff
- Weekly schedule for a selected person
- People available on a specified date
- People assigned to a specified project
- Time conflicts
- Leave schedule
- Site assignments this week
- Monthly schedule overview

For these questions, use repository data only and cite the workbook sheet/row when possible.
