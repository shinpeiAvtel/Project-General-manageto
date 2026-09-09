# Project General manageto

複数プロジェクトの工程表（Excel）を一元管理し、正規化・統合・自動分析してMaster Scheduleを生成するProject Management Systemです。

## システム目的
- Input配下へ担当者が工程表をアップロード
- Pythonで取り込み/検証/正規化/分析
- 遅延・競合・更新漏れ・リスクを可視化
- Copilot/Agentが根拠付きで回答可能なデータ基盤を提供

## Repository構成
- `Input/` アップロード領域
- `Output/` 生成物（Master Schedule/Reports）
- `scripts/` 取り込み〜分析〜出力
- `data/` 中間データ（raw/normalized/master）
- `config/` スキーマ/ルール
- `docs/` 運用文書
- `tests/` pytest
- `.github/workflows/` 自動実行

## 初回セットアップ
```bash
python -m venv .venv
source .venv/bin/activate
pip install -r requirements.txt
```

## Inputへのアップロード方法
`Input/<担当者名>/<ProjectName>/<file>.xlsx` 形式で配置してください。

## Excelテンプレート使用方法
`templates/Project_Schedule_Template.xlsx` を利用し、固定ヘッダーを変更しないでください。

## ローカル実行方法
```bash
python scripts/scan_input.py
python scripts/import_excel.py
python scripts/validate_schedule.py
python scripts/normalize_schedule.py
python scripts/analyze_schedule.py
python scripts/generate_master_schedule.py
python scripts/generate_reports.py
```

## GitHub Actions動作
`Input/**/*.xlsx` のpushで `schedule-analysis.yml` が実行され、分析結果をArtifactとして取得できます。

## Output確認方法
- `Output/Master_Schedule/Master_Schedule.xlsx`
- `Output/Reports/Upcoming_14_Days.xlsx`
- `Output/Reports/Manager_Summary.md`
- `Output/Reports/Validation_Report.json`

## Copilot Agentへの質問方法
分析対象は `Output` と `data/master/analysis.json` を参照し、回答時に日付とSource Fileを明記してください。

## Validation Error対応方法
`Validation_Report.json` の `severity/code/message/source_*` を確認し、元Excelを修正して再実行してください。

## Outputの運用方針
Phase 1ではActions Artifact配布を基本とし、Outputの恒久commitは運用判断で切替可能です。

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
