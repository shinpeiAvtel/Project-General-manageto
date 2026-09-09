# 分析ルール

- End Date超過かつProgress<100: `DELAYED`
- Start Date超過かつProgress=0: `NOT_STARTED_LATE`
- 担当重複は `POTENTIAL_CONFLICT`
- リスクスコアは `config/project_rules.yaml` に従って算出
