# Analysis Rules

- End Date超過かつProgress<100: `DELAYED`
- Start Date超過かつProgress=0: `NOT_STARTED_LATE`
- 同一担当者の期間重複: `POTENTIAL_CONFLICT`
- Milestone=falseで重要カテゴリ: `MILESTONE_CANDIDATE`
