# Project General manageto

複数プロジェクトの工程表（Excel）を一元管理し、正規化・検証・統合・分析を自動化する Project Management System です。

## システム目的
- 担当者が `Input/` にアップロードした工程表を自動集約
- 共通フォーマットへ正規化
- 全案件統合の Master Schedule 作成
- 遅延・担当重複・更新漏れ・リスクを自動分析

## Repository構成
- `Input/`: 各担当者の工程表アップロード先
- `Output/`: Master Schedule・分析レポート出力
- `scripts/`: 取り込み/検証/分析/出力生成
- `data/`: 正規化データ・分析中間データ
- `docs/`: 運用ドキュメント
- `tests/`: pytest
- `.github/`: Actions / Copilot指示

## 初回セットアップ
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Python環境構築
Python 3.11 以上を推奨します。

## Inputへのアップロード方法
以下の階層で `.xlsx` を配置します。

```text
Input/<担当者名>/<ProjectName>/<file>.xlsx
```

例:
- `Input/Shiraishi/NRT82/NRT82_schedule.xlsx`
- `Input/Tanaka/HND10/HND10_schedule.xlsx`

## Excelテンプレート使用方法
`templates/Project_Schedule_Template.xlsx` をコピーして使用してください。

## ローカル実行方法
```bash
python -m scripts.scan_input
```

## GitHub Actions動作
`Input/**/*.xlsx` または関連スクリプト更新時に `.github/workflows/schedule-analysis.yml` が動作し、
取り込み〜検証〜正規化〜分析〜出力生成〜pytest を実行します。
成果物は Artifact として取得可能です。

## Output確認方法
- `Output/Master_Schedule/Master_Schedule.xlsx`
- `Output/Reports/Upcoming_14_Days.xlsx`
- `Output/Reports/Manager_Summary.md`
- `Output/Reports/Validation_Report.json`

## Copilot Agentへの質問方法
正規化/分析済みデータを前提に、例として次を質問できます。
- 現在遅延しているプロジェクト
- 今週〜来週の重要工程
- 担当者の重複アサイン候補
- 更新漏れ案件

## Validation Error対応方法
`Validation_Report.json` と Master の `Validation Errors` シートを確認し、
元Excelを修正後に再実行してください。自動補完は行いません。

## Outputの運用方針
Phase 1 では Output をCI Artifact配布中心とし、リポジトリへの自動コミットは未設定です。

## 将来拡張（Phase 2候補）
- Smartphone Upload Web UI
- FastAPI
- Project Dashboard
- Gantt Chart
- PDF工程表解析
- CSV対応
- Microsoft Project連携
- Primavera P6連携
- Outlook / Calendar連携
- Slack / Teams通知
- Weekly Project Report
- Automatic Reminder
- Baseline vs Actual比較
- Project Portfolio Dashboard
