"""
役員向け定例報告 自動生成スクリプト

処理の流れ:
1. Google Sheetsから最新データを取得
2. alias.jsonを参照してコードネームに変換
3. Claude APIで分析・レポート生成
4. 逆変換して実名に復元
5. Markdownファイルとして出力
"""

import base64
import json
import os
import re
from datetime import date, datetime, timedelta

import gspread
from google.oauth2.service_account import Credentials
import anthropic


# ─────────────────────────────────────────
# 設定・データ読み込み
# ─────────────────────────────────────────

def load_config(config_path: str = "config.json") -> dict:
    with open(config_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_config_from_env() -> dict:
    """GitHub Actions 用: 環境変数から config を構築する。"""
    return {
        "spreadsheets": {
            "cases_data": {
                "url":        os.environ["SPREADSHEET_CASES_URL"],
                "sheet_name": os.environ["SPREADSHEET_CASES_SHEET"],
            },
            "billing_data": {
                "url":        os.environ["SPREADSHEET_BILLING_URL"],
                "sheet_name": os.environ["SPREADSHEET_BILLING_SHEET"],
            },
            "roadmap_data": {
                "url":        os.environ["SPREADSHEET_ROADMAP_URL"],
                "sheet_name": os.environ["SPREADSHEET_ROADMAP_SHEET"],
            },
            "kpi_data": {
                "url":        os.environ["SPREADSHEET_KPI_URL"],
                "sheet_name": os.environ["SPREADSHEET_KPI_SHEET"],
            },
        },
        "claude_api_key":             os.environ["CLAUDE_API_KEY"],
        "gamma_api_key":              os.environ.get("GAMMA_API_KEY", ""),
        "credentials_file":           "credentials.json",
        "output_dir":                 "reports",
        "slack_md_leader_channel_id": os.environ["SLACK_MD_LEADER_CHANNEL_ID"],
        "slack_assign_channel_id":    os.environ["SLACK_ASSIGN_CHANNEL_ID"],
        "gamma_latest_url":           os.environ.get("GAMMA_LATEST_URL", ""),
        "slack_bot_token":            os.environ["SLACK_BOT_TOKEN"],
        "slack_test_mode":            os.environ.get("SLACK_TEST_MODE", "false").lower() == "true",
        "google_report_folder_id":    os.environ["GOOGLE_REPORT_FOLDER_ID"],
    }


def load_alias(alias_path: str = "alias.json") -> dict:
    with open(alias_path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_alias_map(alias_data: dict) -> dict:
    """alias.jsonの全カテゴリを1つのフラットな辞書に結合する"""
    alias_map = {}
    for key, mapping in alias_data.items():
        if key.startswith("_"):
            continue
        if isinstance(mapping, dict):
            alias_map.update(mapping)
    return alias_map


# ─────────────────────────────────────────
# Google Sheets
# ─────────────────────────────────────────

def get_google_sheets_client(credentials_path: str) -> gspread.Client:
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
    ]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    return gspread.authorize(creds)


def extract_spreadsheet_id(url: str) -> str:
    match = re.search(r"/spreadsheets/d/([a-zA-Z0-9-_]+)", url)
    if match:
        return match.group(1)
    raise ValueError(f"無効なGoogle SheetsのURL: {url}")


def fetch_sheet_data(client: gspread.Client, url: str, sheet_name: str) -> list[list]:
    spreadsheet_id = extract_spreadsheet_id(url)
    spreadsheet = client.open_by_key(spreadsheet_id)
    worksheet = spreadsheet.worksheet(sheet_name)
    return worksheet.get_all_values()


def calculate_mrr_summary(billing_rows: list[list], today: date, target_mrr: int = 24_000_000) -> str:
    """案件管理DBから当月（または直近確定月）のMRR（粗利ベース）を集計してサマリ文字列を返す。

    シート構造:
      row[0] = 集計行（数式）
      row[1] = ヘッダー行
      row[2:] = データ行
    主要列（0始まり）:
      col[7]  = 単発/月額（完全一致「月額」のみ対象）
      col[14] = 粗利見込（税抜）← MRR集計の基準値
      col[17] = 請求(開始)月（YYYY/MM 形式）← 正しい請求月列
    """
    if len(billing_rows) < 3:
        return "データなし（行数不足）"

    data_rows = billing_rows[2:]

    def parse_amount(s: str) -> float:
        s = str(s).replace(",", "").replace("¥", "").replace(" ", "").strip()
        try:
            return float(s)
        except ValueError:
            return 0.0

    def to_yyyymm(m: str) -> int:
        """YYMM or YYYY/MM → int YYYYMM。解析できない場合は 0 を返す"""
        m = str(m).strip().replace("/", "").replace("-", "")
        if len(m) == 4:   # YYMM（例: "2604"）
            return int("20" + m)
        if len(m) == 6:   # YYYYMM（例: "202604"）
            return int(m)
        return 0

    current_yyyymm = int(today.strftime("%Y%m"))

    # 前月を計算
    if current_yyyymm % 100 == 1:
        prev_yyyymm = (current_yyyymm // 100 - 1) * 100 + 12
    else:
        prev_yyyymm = current_yyyymm - 1

    # col[14]（粗利見込）がある月額案件を請求(開始)月ごとに集計
    from collections import defaultdict
    month_totals: dict = defaultdict(float)
    for r in data_rows:
        if len(r) > 17 and r[7].strip() == "月額":
            nm = to_yyyymm(r[17])
            gross = parse_amount(r[14]) if len(r) > 14 else 0.0
            if nm > 0 and gross > 0:
                month_totals[nm] += gross

    if not month_totals:
        return "月額案件の粗利データなし"

    # 前月確定値を優先。なければ当月、それもなければ直近確定月にフォールバック
    if prev_yyyymm in month_totals:
        target_yyyymm = prev_yyyymm
        yr, mo = target_yyyymm // 100, target_yyyymm % 100
        month_label = f"{yr}年{mo:02d}月"
        note = ""
    elif current_yyyymm in month_totals:
        target_yyyymm = current_yyyymm
        month_label = today.strftime("%Y年%m月")
        note = ""
    else:
        past = {ym: v for ym, v in month_totals.items() if ym < current_yyyymm}
        target_yyyymm = max(past.keys()) if past else max(month_totals.keys())
        yr, mo = target_yyyymm // 100, target_yyyymm % 100
        month_label = f"{yr}年{mo:02d}月"
        note = "（直近確定月）"

    # 対象月の行を抽出してMRR（粗利ベース）集計
    target_rows = [
        r for r in data_rows
        if len(r) > 17 and r[7].strip() == "月額" and to_yyyymm(r[17]) == target_yyyymm
    ]

    total_gross = sum(
        parse_amount(r[14]) for r in target_rows if len(r) > 14
    )
    diff = total_gross - target_mrr
    sign = "+" if diff >= 0 else ""

    return (
        f"【MRR実績サマリ（粗利ベース） {month_label}{note}】\n"
        f"- MRR（粗利ベース・税抜）: {total_gross / 10_000:,.0f}万円 / "
        f"目標2,400万円（差異：{sign}{diff / 10_000:,.0f}万円）\n"
        f"- 集計対象件数: {len(target_rows)}件"
    )


# ─────────────────────────────────────────
# Google Meet 文字起こし取得
# ─────────────────────────────────────────

# 対象MTGのタイトル・本文に含まれる識別キーワード
MEETING_TAGS = [
    "VUP/SU 定例MTG",
    "MDプロジェクト",
    "ソリューションリーダーズ",
    "MDリーダー定例",
]

# 議事録前処理プロンプト（第1段階）
MEETING_EXTRACT_PROMPT = """以下の会議議事録から、経営判断・チーム運営に関わる情報のみを抽出してください。

抽出する情報：
1. 決定事項（「〜することになった」「〜で合意」「〜に決定」に該当するもの）
2. 担当者が明示されたアクションアイテム（誰が・何を・いつまでに）
3. 未解決の課題・リスク（懸念として挙げられたが結論が出ていないもの）
4. 数値・KPIに関する言及（目標・実績・差異）

抽出しない情報：
- 挨拶・雑談・場つなぎの発言
- 既に完了済みの報告のみの内容
- 文脈なく単独では意味をなさない発言

必ず以下のフォーマットのみで出力してください（他のテキスト不要）：

MEETING_TITLE: 会議名
MEETING_DATE: 日付（不明の場合は空欄）
DECISIONS:
・決定事項（なければ行ごと省略可）
ACTIONS:
・[担当者名] アクション内容（期限：XX）
ISSUES:
・課題・リスク（なければ行ごと省略可）
NUMBERS:
・数値・KPI言及（なければ行ごと省略可）

会議議事録：
{transcript}"""


def _get_docs_drive_clients(credentials_path: str):
    """Drive / Docs APIクライアントを返す。失敗時は (None, None)。"""
    try:
        from googleapiclient.discovery import build
    except ImportError:
        return None, None
    scopes = [
        "https://www.googleapis.com/auth/spreadsheets.readonly",
        "https://www.googleapis.com/auth/drive.readonly",
        "https://www.googleapis.com/auth/documents.readonly",
    ]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    drive = build("drive", "v3", credentials=creds)
    docs = build("docs", "v1", credentials=creds)
    return drive, docs


def _extract_doc_text(docs_service, doc_id: str) -> str:
    """Google Doc の全テキストを返す。"""
    doc = docs_service.documents().get(documentId=doc_id).execute()
    lines = []
    for elem in doc.get("body", {}).get("content", []):
        para = elem.get("paragraph")
        if para:
            text = "".join(
                r.get("textRun", {}).get("content", "")
                for r in para.get("elements", [])
            ).rstrip("\n")
            if text:
                lines.append(text)
    return "\n".join(lines)


def _parse_meeting_sections(title: str, text: str) -> dict:
    """テキストを 概要 / 次のステップ / 詳細 セクションに分割して返す。"""
    sections = {"title": title, "概要": "", "次のステップ": "", "詳細": ""}
    SECTION_NAMES = {"概要", "次のステップ", "詳細"}
    current, buf = None, []
    for line in text.splitlines():
        if line.strip() in SECTION_NAMES:
            if current:
                sections[current] = "\n".join(buf).strip()
            current, buf = line.strip(), []
        elif current:
            buf.append(line)
    if current:
        sections[current] = "\n".join(buf).strip()
    return sections


def fetch_meeting_transcripts(credentials_path: str, days: int = 7) -> list[dict]:
    """直近N日以内に更新されたMTG議事録Docを取得・パースして返す。

    Drive API / Docs API が未有効またはアクセス権なしの場合は空リストを返す（graceful degradation）。
    """
    drive_service, docs_service = None, None
    try:
        drive_service, docs_service = _get_docs_drive_clients(credentials_path)
    except Exception as e:
        print(f"      [SKIP] Google API初期化エラー: {e}")
        return []

    if drive_service is None:
        print("      [SKIP] google-api-python-client 未インストール")
        return []

    since = (datetime.utcnow() - timedelta(days=days)).strftime("%Y-%m-%dT%H:%M:%SZ")
    found: dict[str, str] = {}  # {doc_id: doc_name}

    for keyword in MEETING_TAGS:
        # Drive APIのfullText検索（タイトルまたは本文に含まれるもの）
        q = (
            f"mimeType='application/vnd.google-apps.document' "
            f"and modifiedTime > '{since}' "
            f"and (fullText contains '{keyword}' or name contains '{keyword}')"
        )
        try:
            res = drive_service.files().list(
                q=q,
                fields="files(id, name, modifiedTime)",
                pageSize=10,
                orderBy="modifiedTime desc",
            ).execute()
            for f in res.get("files", []):
                found[f["id"]] = f["name"]
        except Exception as e:
            print(f"      [WARN] Drive検索エラー ({keyword}): {e}")

    print(f"      [INFO] 対象MTG議事録: {len(found)}件")

    results = []
    for doc_id, doc_name in found.items():
        try:
            text = _extract_doc_text(docs_service, doc_id)
            results.append(_parse_meeting_sections(doc_name, text))
            print(f"        取得: {doc_name}")
        except Exception as e:
            print(f"        [WARN] {doc_name} 取得エラー: {e}")

    return results


def _summarize_meeting_with_claude(
    client: anthropic.Anthropic,
    meeting: dict,
    alias_map: dict,
) -> dict:
    """議事録1件をClaude APIで構造化JSONに変換する。失敗時はNoneを返す。"""
    parts = [f"会議名: {meeting['title']}"]
    for section in ("概要", "次のステップ", "詳細"):
        body = meeting.get(section, "").strip()
        if body:
            parts.append(f"【{section}】\n{body}")

    full_text = "\n\n".join(parts)
    # 実名をコードネームに変換してからAPIに送る
    anon_text = apply_alias(full_text, alias_map)

    prompt = MEETING_EXTRACT_PROMPT.format(transcript=anon_text)
    try:
        resp = client.messages.create(
            model="claude-opus-4-6",
            max_tokens=1024,
            messages=[{"role": "user", "content": prompt}],
        )
        raw = resp.content[0].text.strip()
        return _parse_structured_meeting_text(raw, meeting["title"])
    except Exception as e:
        print(f"        [WARN] 前処理失敗 ({meeting['title']}): {e}")
    return None


def _parse_structured_meeting_text(text: str, fallback_title: str) -> dict:
    """構造化テキスト形式の議事録サマリーを辞書に変換する。"""
    result = {
        "meeting_title": fallback_title,
        "meeting_date": "",
        "decisions": [],
        "actions": [],
        "issues": [],
        "numbers": [],
    }
    current_section = None
    SECTION_MAP = {
        "DECISIONS":    "decisions",
        "ACTIONS":      "actions",
        "ISSUES":       "issues",
        "NUMBERS":      "numbers",
    }
    for line in text.splitlines():
        line_stripped = line.strip()
        if not line_stripped:
            continue
        if line_stripped.startswith("MEETING_TITLE:"):
            result["meeting_title"] = line_stripped.split(":", 1)[1].strip() or fallback_title
            current_section = None
        elif line_stripped.startswith("MEETING_DATE:"):
            result["meeting_date"] = line_stripped.split(":", 1)[1].strip()
            current_section = None
        elif line_stripped.rstrip(":") in SECTION_MAP:
            current_section = SECTION_MAP[line_stripped.rstrip(":")]
        elif current_section and line_stripped.startswith("・"):
            item = line_stripped.lstrip("・").strip()
            if item:
                if current_section == "actions":
                    # [担当者] アクション（期限：XX） 形式をパース
                    m = re.match(r"\[(.+?)\]\s*(.+?)(?:（期限[：:]\s*(.+?)）)?$", item)
                    if m:
                        result["actions"].append({
                            "who": m.group(1).strip(),
                            "what": m.group(2).strip(),
                            "by_when": m.group(3).strip() if m.group(3) else "",
                        })
                    else:
                        result["actions"].append({"who": "", "what": item, "by_when": ""})
                else:
                    result[current_section].append(item)
    return result


def _format_meeting_summaries(summaries: list[dict]) -> str:
    """構造化済みJSONサマリーをレポートプロンプト用テキストに整形する。"""
    parts = []
    for s in summaries:
        lines = [f"### {s.get('meeting_title', '不明')}"]
        if s.get("meeting_date"):
            lines.append(f"日付: {s['meeting_date']}")

        if s.get("decisions"):
            lines.append("**決定事項**")
            for d in s["decisions"]:
                lines.append(f"- {d}")

        if s.get("actions"):
            lines.append("**アクションアイテム**")
            for a in s["actions"]:
                if isinstance(a, dict):
                    who = a.get("who", "")
                    what = a.get("what", "")
                    by_when = a.get("by_when", "")
                    lines.append(f"- {who}：{what}（期限：{by_when}）")
                else:
                    lines.append(f"- {a}")

        if s.get("issues"):
            lines.append("**未解決の課題・リスク**")
            for i in s["issues"]:
                lines.append(f"- {i}")

        if s.get("numbers"):
            lines.append("**数値・KPI**")
            for n in s["numbers"]:
                lines.append(f"- {n}")

        # フォールバック用生テキスト（抽出失敗時のみ）
        if s.get("_raw"):
            lines.append(s["_raw"])

        if len(lines) > 1:  # タイトル行だけの場合は除外
            parts.append("\n".join(lines))

    return "\n\n---\n\n".join(parts) if parts else "（有効な議事録データなし）"


def format_meeting_transcripts(
    meetings: list[dict],
    claude_client: anthropic.Anthropic = None,
    alias_map: dict = None,
) -> str:
    """パース済みMTG議事録をプロンプト用文字列に整形する。

    claude_client が指定されている場合は2段階処理：
      第1段階: 議事録ごとにClaude APIで決定事項・アクション・課題・数値を抽出（JSON）
      第2段階: 全件のJSONを統合してレポート用テキストに整形
    指定がない場合は生テキストのままフォールバック。
    """
    if not meetings:
        return "（議事録データなし：Drive/Docs APIが未設定、または直近7日以内の対象ドキュメントなし）"

    if claude_client is None:
        # フォールバック：生テキスト整形
        parts = []
        for m in meetings:
            block = f"### {m['title']}\n"
            if m.get("概要"):
                block += f"**概要**\n{m['概要']}\n\n"
            if m.get("次のステップ"):
                block += f"**次のステップ**\n{m['次のステップ']}\n\n"
            if m.get("詳細"):
                detail = m["詳細"][:2000]
                if len(m["詳細"]) > 2000:
                    detail += "\n（...以下省略）"
                block += f"**詳細**\n{detail}\n"
            parts.append(block.strip())
        return "\n\n---\n\n".join(parts)

    # ── 第1段階：議事録ごとに前処理 ──
    alias_map = alias_map or {}
    summaries = []
    for m in meetings:
        print(f"        前処理中: {m['title'][:50]}")
        result = _summarize_meeting_with_claude(claude_client, m, alias_map)
        if result:
            summaries.append(result)
        else:
            # 失敗時：最小限の生テキストをフォールバックとして保持
            summaries.append({
                "meeting_title": m["title"],
                "meeting_date": "",
                "decisions": [],
                "actions": [],
                "issues": [],
                "numbers": [],
                "_raw": apply_alias(
                    (m.get("概要") or "") + "\n" + (m.get("次のステップ") or ""),
                    alias_map,
                ).strip(),
            })

    # ── 第2段階：統合サマリーを整形 ──
    return _format_meeting_summaries(summaries)


# ─────────────────────────────────────────
# エイリアス変換・逆変換
# ─────────────────────────────────────────

def apply_alias(text: str, alias_map: dict) -> str:
    """実名 → コードネームに変換（長い名前から順に処理して部分一致を防ぐ）"""
    for real_name in sorted(alias_map.keys(), key=len, reverse=True):
        codename = alias_map[real_name]
        text = text.replace(real_name, codename)
    return text


def reverse_alias(text: str, alias_map: dict) -> str:
    """コードネーム → 実名に逆変換"""
    reverse_map = {v: k for k, v in alias_map.items()}
    for codename in sorted(reverse_map.keys(), key=len, reverse=True):
        real_name = reverse_map[codename]
        text = text.replace(codename, real_name)
    return text


def anonymize_data(data: list[dict], alias_map: dict) -> str:
    """リストデータをJSON文字列に変換してエイリアス適用"""
    data_str = json.dumps(data, ensure_ascii=False, indent=2)
    return apply_alias(data_str, alias_map)


# ─────────────────────────────────────────
# Claude APIによるレポート生成
# ─────────────────────────────────────────

REPORT_PROMPT_TEMPLATE = """
あなたはプロの経営コンサルタントです。
以下のデータをもとに、役員とチームリーダー向けの週次報告を作成してください。

【読者】
- 役員（吉田典広）：経営判断に必要な数字とリスクだけ見たい
- MDチームリーダー4名：自分たちの行動に直結する情報だけ見たい

【絶対に守るルール】
1. 「〜と考えられます」「〜の可能性があります」は禁止。断言する
2. データがない項目は「未集計」と書いて飛ばす。推測で埋めない
3. アクションは「誰が・何を・いつまでに」の形式のみ
4. 課題・アクション・案件は件数上限なし。データにある全件を記載する
5. 各セクションのタイトルは必ず「## N.」の形式で記載すること（## 5.5. も含む）

【レポート構成】

# MD週次報告 {date}

{data_sources}

## 1. 数字サマリー（事実のみ）
| 指標 | 実績 | 目標 | 差異 | 判定 |
|------|------|------|------|------|
| MRR（粗利） | {mrr}万円 | 2,400万円 | {diff}万円 | {判定} |
| 継続率 | 未集計 | 90% | - | - |
| アップセル | 未集計 | 15件 | - | - |

※判定基準：90%以上=✅ 75-89%=⚠️ 75%未満=🚨

## 2. Unit別・個人別KPI達成状況（全員分）
KPIデータに含まれる全員・全Unitの達成状況を記載。
未達の人員は**太字**で強調。目標値・実績値・達成率を明記。
データがなければ「未集計」と記載。

## 3. 今週の最重要課題（全件・優先度順）
件数上限なし。データから読み取れる全ての重要課題を優先度順に列挙。
各課題は以下の形式で書く：

**[課題名]**
- 課題の事実：〇〇
- 放置した場合の影響：〇〇
- 必要な意思決定：〇〇
- 根拠データ：〇〇

## 4. 今週の意思決定・アクション（全件）
件数上限なし。全ての意思決定・アクションを記載。
担当者名・アクション内容・期限・根拠データを明記。

## 5. 注視案件（契約リスク・全件）
件数上限なし。契約満了・継続不確定・解約リスクのある全案件を記載。
表形式で：顧客名 | 満了日 | リスクレベル | 担当者 | 現状一言

## 5.5. 粗利率26%未満の案件（要改善・全件）
粗利率が26%未満の全案件を件数上限なく記載。
表形式で：顧客名 | 粗利率 | 担当者 | 改善アクション

## 6. 会議決定事項（議事録より）
直近7日以内の会議で決定した事項・アクション・課題を全件記載。
会議名 | 決定事項 | 担当者 | 期限

## 7. 先週との変化（差分のみ）
良くなったこと・悪くなったことを件数制限なく全て記載。

---

以下がデータです：

MRRサマリー：
{mrr_summary}

KPIデータ：
{kpi_data}

案件データ（要注意フラグ付き）：
{cases_summary}

ロードマップ進捗：
{roadmap_summary}

MTG議事録（直近7日）：
{meeting_transcripts}
"""


def _parse_mrr_for_prompt(mrr_summary: str) -> tuple:
    """mrr_summaryから表示用のMRR値・差異・判定記号を抽出する。"""
    mrr_val = "未集計"
    diff_val = "-"
    judgment = "-"

    m_mrr = re.search(r"([\d,]+)万円\s*/\s*目標", mrr_summary)
    if m_mrr:
        mrr_str = m_mrr.group(1).replace(",", "")
        try:
            mrr_int = int(mrr_str)
            mrr_val = f"{mrr_int:,}"
            diff_int = mrr_int - 2400
            sign = "+" if diff_int >= 0 else ""
            diff_val = f"{sign}{diff_int:,}"
            rate = mrr_int / 2400
            judgment = "✅" if rate >= 0.90 else ("⚠️" if rate >= 0.75 else "🚨")
        except ValueError:
            pass

    return mrr_val, diff_val, judgment


def generate_report_with_claude(
    client: anthropic.Anthropic,
    mrr_summary: str,
    roadmap_data_anon: str,
    kpi_data_anon: str,
    cases_data_anon: str,
    meeting_data_anon: str,
    today_str: str,
    data_sources: str = "",
) -> str:
    mrr_val, diff_val, judgment = _parse_mrr_for_prompt(mrr_summary)

    prompt = REPORT_PROMPT_TEMPLATE.format(
        date=today_str,
        mrr=mrr_val,
        diff=diff_val,
        判定=judgment,
        mrr_summary=mrr_summary,
        kpi_data=kpi_data_anon,
        cases_summary=cases_data_anon,
        roadmap_summary=roadmap_data_anon,
        meeting_transcripts=meeting_data_anon,
        data_sources=data_sources,
    )

    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=8192,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


# ─────────────────────────────────────────
# Slack 連携
# ─────────────────────────────────────────

def fetch_slack_messages(token: str, channel_id: str, days: int = 7) -> list[dict]:
    """Slack チャンネルから直近N日以内のメッセージを取得する。
    失敗時は空リストを返す（graceful degradation）。
    """
    try:
        from slack_sdk import WebClient
        from slack_sdk.errors import SlackApiError
    except ImportError:
        print("      [SKIP] slack_sdk 未インストール")
        return []

    if not token or token.startswith("xoxb-YOUR"):
        print(f"      [SKIP] slack_bot_token 未設定 (channel: {channel_id})")
        return []

    client = WebClient(token=token)
    oldest = str((datetime.utcnow() - timedelta(days=days)).timestamp())

    messages = []
    cursor = None
    try:
        while True:
            kwargs = {"channel": channel_id, "oldest": oldest, "limit": 200}
            if cursor:
                kwargs["cursor"] = cursor
            resp = client.conversations_history(**kwargs)
            messages.extend(resp.get("messages", []))
            meta = resp.get("response_metadata", {})
            cursor = meta.get("next_cursor")
            if not cursor:
                break
    except SlackApiError as e:
        print(f"      [WARN] Slack取得エラー (ch:{channel_id}): {e.response['error']}")
        return []

    # bot/システムメッセージを除外し、古い順に返す
    user_msgs = [m for m in messages if m.get("type") == "message" and not m.get("bot_id") and m.get("text")]
    user_msgs.sort(key=lambda m: float(m.get("ts", 0)))
    return user_msgs


def _resolve_user_mentions(text: str, alias_map: dict) -> str:
    """<@UXXXXXXX|表示名> 形式のメンションを表示名だけに変換する。"""
    return re.sub(r"<@[A-Z0-9]+\|([^>]+)>", r"\1", text)


def format_slack_messages(messages: list[dict], channel_name: str, alias_map: dict) -> str:
    """Slackメッセージをプロンプト用テキストに整形する。"""
    if not messages:
        return f"（{channel_name}：直近7日以内のメッセージなし、またはAPI未設定）"

    lines = [f"【{channel_name} 直近7日 抜粋（{len(messages)}件）】"]
    for m in messages[-40:]:  # 最新40件に絞る
        ts = datetime.utcfromtimestamp(float(m.get("ts", 0)))
        ts_str = (ts + timedelta(hours=9)).strftime("%m/%d %H:%M")
        text = _resolve_user_mentions(m.get("text", ""), alias_map)
        text = apply_alias(text, alias_map)
        # 長すぎるメッセージは切る
        if len(text) > 300:
            text = text[:300] + "…"
        lines.append(f"[{ts_str}] {text}")

    return "\n".join(lines)


SLACK_HEADLINE_PROMPT = """あなたはMDチームの役員秘書です。
以下の週次レポートをもとに、Slackの見出しメッセージを作成してください。

ルール：
- 読んで3秒で状況が分かること
- 数字は必ず入れる
- 課題は「何が問題で、誰が何をすべきか」まで1行で書く
- アクションは「誰が・何を・いつまでに」の形式のみ
- 絵文字を効果的に使う
- 全体で15行以内

フォーマット：
📊 *MD週次報告 [日付]*

💰 MRR [実績]万円 / 目標2,400万円（[達成率]%・[差異]万円[不足/超過]）

🚨 *今週の最重要課題*
1. [課題名]：[誰が][何をすべきか]（期限）
2. [課題名]：[誰が][何をすべきか]（期限）
3. [課題名]：[誰が][何をすべきか]（期限）

✅ *今週の意思決定・アクション*
- [担当者] [アクション]（[期限]）
- [担当者] [アクション]（[期限]）
- [担当者] [アクション]（[期限]）

⚠️ *注視案件*
- [顧客名]（[満了日]）：[リスクレベル] - [現状一言]

📎 詳細レポート（スライド）：{gamma_url}

レポートデータ：{report_content}"""


def generate_slack_headline(
    client: anthropic.Anthropic,
    report_md: str,
    gamma_url: str,
) -> str:
    """Claude APIでSlack投稿用の見出しメッセージを生成する。"""
    prompt = SLACK_HEADLINE_PROMPT.format(
        gamma_url=gamma_url if gamma_url else "（Gamma未生成）",
        report_content=report_md,
    )
    message = client.messages.create(
        model="claude-opus-4-6",
        max_tokens=2048,
        messages=[{"role": "user", "content": prompt}],
    )
    return message.content[0].text


def post_slack_report(
    token: str, channel_id: str, headline: str, test_mode: bool = False
) -> bool:
    """MDリーダーチャンネルに見出しメッセージを投稿する。
    test_mode=True のとき実際には投稿せずプレビューを標準出力する。
    """
    if test_mode:
        print("      [TEST MODE] Slack投稿をスキップ。見出しプレビュー:")
        print("─" * 60)
        print(headline)
        print("─" * 60)
        return True

    try:
        from slack_sdk import WebClient
        from slack_sdk.errors import SlackApiError
    except ImportError:
        print("      [SKIP] slack_sdk 未インストール")
        return False

    if not token or token.startswith("xoxb-YOUR"):
        print("      [SKIP] slack_bot_token 未設定")
        return False

    client = WebClient(token=token)
    try:
        client.chat_postMessage(channel=channel_id, text=headline, mrkdwn=True)
        print(f"      Slack投稿完了: #{channel_id}")
        return True
    except SlackApiError as e:
        print(f"      [WARN] Slack投稿エラー: {e.response['error']}")
        return False


# ─────────────────────────────────────────
# Google Docs 出力
# ─────────────────────────────────────────

def _get_drive_write_client(credentials_path: str):
    """Drive 書き込み用クライアントを返す。"""
    from googleapiclient.discovery import build as gapi_build
    scopes = ["https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file(credentials_path, scopes=scopes)
    return gapi_build("drive", "v3", credentials=creds)


def _parse_bold(text: str) -> tuple:
    """**太字** マークアップを除去し (プレーン文字列, [(start,end), ...]) を返す。"""
    plain = ""
    bold_ranges = []
    i = 0
    while i < len(text):
        if text[i:i+2] == "**":
            j = text.find("**", i + 2)
            if j == -1:
                plain += text[i:]
                break
            start = len(plain)
            inner = text[i+2:j]
            plain += inner
            bold_ranges.append((start, start + len(inner)))
            i = j + 2
        else:
            plain += text[i]
            i += 1
    return plain, bold_ranges


def _bold_req(start: int, end: int) -> dict:
    return {
        "updateTextStyle": {
            "range": {"startIndex": start, "endIndex": end},
            "textStyle": {"bold": True},
            "fields": "bold",
        }
    }


def _heading_req(start: int, end: int, style: str) -> dict:
    return {
        "updateParagraphStyle": {
            "range": {"startIndex": start, "endIndex": end},
            "paragraphStyle": {"namedStyleType": style},
            "fields": "namedStyleType",
        }
    }


def _markdown_to_docs_requests(md_text: str) -> list:
    """Markdownテキストを Google Docs batchUpdate リクエストリストに変換する。
    サポート: # h1, ## h2, ### h3, **bold**, 通常テキスト, 空行。
    表（|...|）はプレーンテキストとして挿入する。
    """
    reqs = []
    pos = 1  # Docs のインデックスは 1 始まり（doc body 先頭）

    for raw_line in md_text.splitlines():
        # 見出しレベル判定
        h3 = re.match(r"^### (.+)$", raw_line)
        h2 = re.match(r"^## (.+)$", raw_line)
        h1 = re.match(r"^# (.+)$", raw_line)

        if h3:
            plain, bold_ranges = _parse_bold(h3.group(1))
            text = plain + "\n"
            reqs.append({"insertText": {"location": {"index": pos}, "text": text}})
            reqs.append(_heading_req(pos, pos + len(text), "HEADING_3"))
            for s, e in bold_ranges:
                reqs.append(_bold_req(pos + s, pos + e))
            pos += len(text)
        elif h2:
            plain, bold_ranges = _parse_bold(h2.group(1))
            text = plain + "\n"
            reqs.append({"insertText": {"location": {"index": pos}, "text": text}})
            reqs.append(_heading_req(pos, pos + len(text), "HEADING_2"))
            for s, e in bold_ranges:
                reqs.append(_bold_req(pos + s, pos + e))
            pos += len(text)
        elif h1:
            plain, bold_ranges = _parse_bold(h1.group(1))
            text = plain + "\n"
            reqs.append({"insertText": {"location": {"index": pos}, "text": text}})
            reqs.append(_heading_req(pos, pos + len(text), "HEADING_1"))
            for s, e in bold_ranges:
                reqs.append(_bold_req(pos + s, pos + e))
            pos += len(text)
        else:
            plain, bold_ranges = _parse_bold(raw_line)
            text = plain + "\n"
            reqs.append({"insertText": {"location": {"index": pos}, "text": text}})
            for s, e in bold_ranges:
                reqs.append(_bold_req(pos + s, pos + e))
            pos += len(text)

    return reqs


def _markdown_to_html(md_text: str) -> str:
    """MarkdownをHTML文字列に変換する（Google Docs Upload用）。"""
    def convert_inline(text: str) -> str:
        return re.sub(r'\*\*(.+?)\*\*', r'<strong>\1</strong>', text)

    lines = md_text.splitlines()
    out = ['<html><body>']
    i = 0
    in_table = False

    while i < len(lines):
        line = lines[i]
        stripped = line.strip()

        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped[1:-1].split('|')]
            if not in_table:
                out.append('<table border="1" cellpadding="4" cellspacing="0">')
                in_table = True
                nxt = lines[i + 1].strip() if i + 1 < len(lines) else ''
                if re.match(r'^\|[\s\-:|]+\|$', nxt):
                    out.append('<tr>' + ''.join(f'<th>{convert_inline(c)}</th>' for c in cells) + '</tr>')
                    i += 2
                    continue
            out.append('<tr>' + ''.join(f'<td>{convert_inline(c)}</td>' for c in cells) + '</tr>')
            i += 1
            continue

        if in_table:
            out.append('</table>')
            in_table = False

        if re.match(r'^-{3,}$', stripped):
            out.append('<hr>')
        elif re.match(r'^#### ', line):
            out.append(f'<h4>{convert_inline(line[5:])}</h4>')
        elif re.match(r'^### ', line):
            out.append(f'<h3>{convert_inline(line[4:])}</h3>')
        elif re.match(r'^## ', line):
            out.append(f'<h2>{convert_inline(line[3:])}</h2>')
        elif re.match(r'^# ', line):
            out.append(f'<h1>{convert_inline(line[2:])}</h1>')
        elif stripped == '':
            out.append('<br>')
        else:
            out.append(f'<p>{convert_inline(line)}</p>')

        i += 1

    if in_table:
        out.append('</table>')

    out.append('</body></html>')
    return '\n'.join(out)


def save_to_google_docs(
    report_md: str,
    doc_name: str,
    credentials_path: str,
    folder_id: str,
) -> str:
    """レポートをDrive API経由でGoogle Docとして保存し、URLを返す。失敗時は空文字を返す。"""
    try:
        from googleapiclient.http import MediaInMemoryUpload
        drive_service = _get_drive_write_client(credentials_path)

        file_metadata = {
            "name": doc_name,
            "mimeType": "application/vnd.google-apps.document",
        }
        if folder_id:
            file_metadata["parents"] = [folder_id]

        html_content = _markdown_to_html(report_md)
        media = MediaInMemoryUpload(
            html_content.encode("utf-8"),
            mimetype="text/html",
        )

        file = drive_service.files().create(
            body=file_metadata,
            media_body=media,
            fields="id, webViewLink",
            supportsAllDrives=True,
        ).execute()

        url = file.get("webViewLink", "")
        print(f"      Google Docs保存完了: {url}")
        return url

    except Exception as e:
        print(f"      [WARN] Google Docs保存エラー: {e}")
        if hasattr(e, 'resp'):
            print(f"      HTTPステータス: {e.resp.status}")
        if hasattr(e, 'error_details'):
            print(f"      error_details: {e.error_details}")
        if hasattr(e, 'reason'):
            print(f"      reason: {e.reason}")
        return ""


# ─────────────────────────────────────────
# メイン処理
# ─────────────────────────────────────────

def _restore_credentials_from_env() -> None:
    """GitHub Actions 用: GOOGLE_CREDENTIALS_B64 から credentials.json を復元する。"""
    b64 = os.environ.get("GOOGLE_CREDENTIALS_B64", "")
    if b64:
        with open("credentials.json", "w", encoding="utf-8") as f:
            f.write(base64.b64decode(b64).decode("utf-8"))


def _build_data_sources_text(meetings: list, assign_msgs: list, leader_msgs: list) -> str:
    """参照データソースセクションのMarkdownテキストを構築する。"""
    def _slack_range(msgs: list) -> str:
        if not msgs:
            return "取得なし（0件）"
        tss = [float(m.get("ts", 0)) for m in msgs]
        start = (datetime.utcfromtimestamp(min(tss)) + timedelta(hours=9)).strftime("%m月%d日")
        end = (datetime.utcfromtimestamp(max(tss)) + timedelta(hours=9)).strftime("%m月%d日")
        return f"{start}〜{end}（{len(msgs)}件）"

    lines = ["### 📋 参照データソース", ""]
    lines.append("**会議議事録**")
    if meetings:
        for m in meetings:
            lines.append(f"- {m['title'].strip()}")
    else:
        lines.append("- （直近7日以内の対象議事録なし）")
    lines += [
        "",
        "**Slack**",
        f"- アサインmentチャンネル：{_slack_range(assign_msgs)}",
        f"- MDリーダーチャンネル：{_slack_range(leader_msgs)}",
    ]
    return "\n".join(lines)


def main():
    # GitHub Actions: credentials.json を環境変数から復元
    _restore_credentials_from_env()

    # 設定読み込み（GitHub Actions は環境変数、ローカルは config.json）
    print("[1/12] 設定ファイルを読み込み中...")
    config = load_config_from_env() if os.environ.get("GITHUB_ACTIONS") == "true" else load_config()
    alias_data = load_alias()
    alias_map = build_alias_map(alias_data)

    today = date.today()
    today_str = today.strftime("%Y年%m月%d日")

    slack_token     = config.get("slack_bot_token", "")
    assign_ch       = config.get("slack_assign_channel_id", "")
    leader_ch       = config.get("slack_md_leader_channel_id", "")
    slack_test_mode = config.get("slack_test_mode", False)

    # Claude クライアントを先行初期化（議事録前処理でも使用）
    claude_client = anthropic.Anthropic(api_key=config["claude_api_key"])

    # Google Sheets 接続
    print("[2/12] Google Sheets APIに接続中...")
    gs_client = get_google_sheets_client(config["credentials_file"])

    # データ取得
    print("[3/12] ロードマップデータを取得中...")
    roadmap_data = fetch_sheet_data(
        gs_client,
        config["spreadsheets"]["roadmap_data"]["url"],
        config["spreadsheets"]["roadmap_data"]["sheet_name"],
    )

    print("[4/12] KPIデータを取得中...")
    kpi_data = fetch_sheet_data(
        gs_client,
        config["spreadsheets"]["kpi_data"]["url"],
        config["spreadsheets"]["kpi_data"]["sheet_name"],
    )

    print("[5/12] 案件データを取得中...")
    cases_data = fetch_sheet_data(
        gs_client,
        config["spreadsheets"]["cases_data"]["url"],
        config["spreadsheets"]["cases_data"]["sheet_name"],
    )

    print("[6/12] 請求データを取得・MRR集計中...")
    billing_data = fetch_sheet_data(
        gs_client,
        config["spreadsheets"]["billing_data"]["url"],
        config["spreadsheets"]["billing_data"]["sheet_name"],
    )
    mrr_summary = calculate_mrr_summary(billing_data, today)
    print(f"      {mrr_summary.splitlines()[0]}")

    print("[7/12] MTG議事録を取得・前処理中（直近7日以内）...")
    meetings = fetch_meeting_transcripts(config["credentials_file"])
    meetings = [m for m in meetings if "1on1" not in m["title"]]
    if meetings:
        print(f"      Claude APIで各議事録を前処理中（{len(meetings)}件）...")
    meeting_text = format_meeting_transcripts(meetings, claude_client, alias_map)

    print("[8/12] Slackメッセージを取得中（直近7日以内）...")
    assign_msgs = fetch_slack_messages(slack_token, assign_ch)
    leader_msgs = fetch_slack_messages(slack_token, leader_ch)
    print(f"      アサインch: {len(assign_msgs)}件 / MDリーダーch: {len(leader_msgs)}件")
    assign_text = format_slack_messages(assign_msgs, "アサインメントch", alias_map)
    leader_text = format_slack_messages(leader_msgs, "MDリーダーch", alias_map)

    data_sources_text = _build_data_sources_text(meetings, assign_msgs, leader_msgs)

    # エイリアス変換（匿名化）
    print("[9/12] エイリアス変換（実名 → コードネーム）...")
    roadmap_data_anon = anonymize_data(roadmap_data, alias_map)
    kpi_data_anon = anonymize_data(kpi_data, alias_map)
    cases_data_anon = anonymize_data(cases_data, alias_map)
    meeting_data_anon = apply_alias(meeting_text, alias_map)
    data_sources_anon = apply_alias(data_sources_text, alias_map)
    # Slack textはformat_slack_messages内でapply_alias済み

    # Slackデータをプロンプト用にまとめる
    slack_context = f"{assign_text}\n\n{leader_text}"

    # Claude APIでレポート生成
    print("[10/12] Claude APIでレポートを生成中...")
    report_anon = generate_report_with_claude(
        claude_client,
        mrr_summary,
        roadmap_data_anon,
        kpi_data_anon,
        cases_data_anon,
        meeting_data_anon + "\n\n" + slack_context,
        today_str,
        data_sources=data_sources_anon,
    )

    # 逆変換（コードネーム → 実名）
    print("        逆変換（コードネーム → 実名）...")
    report = reverse_alias(report_anon, alias_map)

    # Markdownファイルに出力
    output_dir = config.get("output_dir", "reports")
    os.makedirs(output_dir, exist_ok=True)
    output_path = os.path.join(output_dir, f"report_{today.strftime('%Y%m%d')}.md")

    with open(output_path, "w", encoding="utf-8") as f:
        f.write(report)

    print(f"      Markdownレポート: {output_path}")

    # Google Docsに保存
    print("[11/12] Google Docsにレポートを保存中...")
    doc_name = f"MD週次報告 {today_str}"
    docs_url = save_to_google_docs(
        report,
        doc_name,
        config["credentials_file"],
        config.get("google_report_folder_id", ""),
    )

    # Slack 配信
    print("[12/12] Slackにレポートを配信中...")
    print("      Claude APIでSlack見出しを生成中...")
    slack_headline = generate_slack_headline(claude_client, report, docs_url)
    post_slack_report(slack_token, leader_ch, slack_headline, test_mode=slack_test_mode)

    # 完了サマリー
    print(f"\n完了！")
    print(f"  Markdownレポート: {output_path}")
    if docs_url:
        print(f"  Google Docs: {docs_url}")


if __name__ == "__main__":
    main()
