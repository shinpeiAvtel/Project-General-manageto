# Copilot Instructions for Project-General-manageto

- 推測禁止。Repository内の確定データを最優先で使用する。
- 数値・日付判定は Python生成データ（正規化JSON/分析結果）を優先する。
- 原工程表に存在しない値を生成しない。
- 不足値は `UNKNOWN` / `MISSING` / `NOT PROVIDED` を許容する。
- 回答には日付を明記する。
- 回答時に Source File を示す。
- Validation Error がある場合は明記する。
- 古いデータを最新情報として扱わない。
- Project間の情報を混同しない。
