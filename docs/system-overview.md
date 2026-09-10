# System Overview

## Purpose

The system consolidates personal Excel schedule sheets into manager-facing outputs while keeping each personal sheet as the source of truth.

## Data Flow

1. Read `Input/Personal_Schedule.xlsx`
2. Detect personal sheets dynamically
3. Validate schedule rows
4. Normalize valid rows into `data/normalized_schedule.json`
5. Generate `Output/Master_Schedule.xlsx`
6. Produce monthly, weekly, conflict, and availability views

## Key Design Rules

- Personal sheet names map directly to `Person`
- System sheets are excluded from personal schedule ingestion
- Output sheets are auto-generated and should not be edited manually
- Source traceability is kept with `Source Sheet` and `Source Row`
- Availability depends on configured workday values only
