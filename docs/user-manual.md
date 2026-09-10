# User Manual

## Updating Personal Schedules

1. Open `Input/Personal_Schedule.xlsx`
2. Add or update rows only in personal sheets
3. Keep the shared sheet header format
4. Save the workbook without editing generated output sheets

## Required Input Format

- `Date`: `YYYY-MM-DD`
- `Start Time`: `HH:MM`
- `End Time`: `HH:MM`
- `Status`: `Planned`, `Confirmed`, `Completed`, `Cancelled`, `Leave`

## Optional Inputs

- `All Day` should be `TRUE` for all-day entries
- Leave `Project` or `Site` blank only when `UNKNOWN` is acceptable after normalization

## Regenerating Outputs

Run:

```bash
python -m scripts.validate_schedule
python -m scripts.build_master_schedule
```
