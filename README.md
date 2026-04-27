# life-pepper-report

役員向け MD 週次報告を自動生成し、Google Docs と Slack に配信するスクリプトです。
毎週水曜 9:00 JST に GitHub Actions で自動実行されます（祝日・年末年始はスキップ）。

---

## ファイル構成

```
.
├── main.py                          # メインスクリプト
├── alias.json                       # エイリアス変換テーブル（顧客名・担当者名）
├── requirements.txt                 # 必要な Python ライブラリ
├── .github/workflows/weekly_report.yml
├── config.json                      # ローカル実行用（.gitignore 済み）
└── credentials.json                 # サービスアカウント鍵（.gitignore 済み）
```

---

## ローカル実行

```bash
pip install -r requirements.txt
python3 main.py
```

---

## GitHub Actions セットアップ

### 1. GOOGLE_CREDENTIALS_B64 の作成

```bash
base64 -i credentials.json | pbcopy   # macOS: クリップボードにコピー
```

### 2. Secrets の登録

リポジトリの **Settings → Secrets and variables → Actions → New repository secret** から以下を登録してください。

#### Google 認証

| Secret 名 | 値 |
|---|---|
| `GOOGLE_CREDENTIALS_B64` | `base64 -i credentials.json` の出力結果 |

#### Claude API

| Secret 名 | 値 |
|---|---|
| `CLAUDE_API_KEY` | config.json の `claude_api_key` |

#### Slack

| Secret 名 | 値 |
|---|---|
| `SLACK_BOT_TOKEN` | config.json の `slack_bot_token` |
| `SLACK_MD_LEADER_CHANNEL_ID` | config.json の `slack_md_leader_channel_id` |
| `SLACK_ASSIGN_CHANNEL_ID` | config.json の `slack_assign_channel_id` |
| `SLACK_TEST_MODE` | `false`（本番） または `true`（テスト） |

#### Google Drive / Docs

| Secret 名 | 値 |
|---|---|
| `GOOGLE_REPORT_FOLDER_ID` | config.json の `google_report_folder_id` |

#### Google Spreadsheets

| Secret 名 | 値 |
|---|---|
| `SPREADSHEET_CASES_URL` | config.json の `spreadsheets.cases_data.url` |
| `SPREADSHEET_CASES_SHEET` | config.json の `spreadsheets.cases_data.sheet_name` |
| `SPREADSHEET_BILLING_URL` | config.json の `spreadsheets.billing_data.url` |
| `SPREADSHEET_BILLING_SHEET` | config.json の `spreadsheets.billing_data.sheet_name` |
| `SPREADSHEET_ROADMAP_URL` | config.json の `spreadsheets.roadmap_data.url` |
| `SPREADSHEET_ROADMAP_SHEET` | config.json の `spreadsheets.roadmap_data.sheet_name` |
| `SPREADSHEET_KPI_URL` | config.json の `spreadsheets.kpi_data.url` |
| `SPREADSHEET_KPI_SHEET` | config.json の `spreadsheets.kpi_data.sheet_name` |

#### Gamma（任意）

| Secret 名 | 値 |
|---|---|
| `GAMMA_API_KEY` | config.json の `gamma_api_key` |
| `GAMMA_LATEST_URL` | config.json の `gamma_latest_url` |

### 3. 手動実行

**Actions → Weekly MD Report → Run workflow** から手動実行できます。
