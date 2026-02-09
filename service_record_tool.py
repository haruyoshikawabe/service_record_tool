# ===== Google Colab 用：サービス支援記録 作成ツール =====
# 依存: openpyxl（Colabには標準で入っている）

from pathlib import Path
from typing import List, Dict
import csv
from openpyxl import load_workbook
from google.colab import files

TEMPLATE_SHEET = "Format"
ATTEND_VALUE = "出席"

CELL_MAP = {
    "office": "B3",
    "date": "B4",
    "user": "G4",
    "time": "B5",
    "method": "G5",
    "program": "A9",
    "dayreport": "A11",
    "temp": "B13",
    "slack": "B15",
}

# ---------- Utility ----------

def normalize_date(s: str) -> str:
    return (s or "").strip().replace("-", "/")

def normalize_time(start: str, end: str) -> str:
    start, end = (start or "").strip(), (end or "").strip()
    return "" if not start and not end else f"{start}～{end}"

def safe_sheet_name(name: str) -> str:
    for c in [":","/","\\","?","*","[","]"]:
        name = name.replace(c, "_")
    return name[:31]

def detect_encoding(path: Path) -> str:
    for enc in ("cp932","shift_jis","utf-8-sig","utf-8"):
        try:
            path.read_text(encoding=enc)
            return enc
        except Exception:
            pass
    raise RuntimeError(f"文字コード判定失敗: {path}")

def read_csv(path: Path) -> List[Dict[str,str]]:
    enc = detect_encoding(path)
    with path.open(encoding=enc, newline="") as f:
        reader = csv.DictReader(f)
        return [{k.strip(): (v or "").strip() for k,v in r.items()} for r in reader]

def build_program(r: Dict[str,str]) -> str:
    out = []
    def add(p, d):
        if p or d:
            out.append(p + ("\n"+d if p and d else d))
    add(r.get("午前のプログラム",""), r.get("午前のプログラム詳細",""))
    add(r.get("午後1のプログラム",""), r.get("午後1のプログラム詳細",""))
    add(r.get("午後2のプログラム",""), r.get("午後2のプログラム詳細",""))
    add(r.get("終日のプログラム",""), r.get("終日のプログラム詳細",""))
    return "\n".join(out)

def normalize_method(raw: str) -> str:
    raw = raw or ""
    if "在宅" in raw and "通所" not in raw:
        return "利用者宅"
    return "事業所"

# ---------- Main ----------

print("① caseMonth CSV をアップロードしてください")
case_files = files.upload()

print("② userCaseDaily CSV をアップロードしてください")
daily_files = files.upload()

print("③ テンプレ Excel をアップロードしてください")
template_files = files.upload()

case_path = Path(next(iter(case_files)))
daily_path = Path(next(iter(daily_files)))
template_path = Path(next(iter(template_files)))

case_rows = read_csv(case_path)
daily_rows = read_csv(daily_path)

# daily を日付キーで辞書化
daily_by_date = {}
date_col = "日付" if "日付" in daily_rows[0] else list(daily_rows[0].keys())[0]
for r in daily_rows:
    daily_by_date[normalize_date(r.get(date_col,""))] = r

wb = load_workbook(template_path)
assert TEMPLATE_SHEET in wb.sheetnames, "Formatシートが見つかりません"

tpl = wb[TEMPLATE_SHEET]

for r in case_rows:
    if r.get("出欠等") != ATTEND_VALUE:
        continue

    date = normalize_date(r.get("年月日",""))
    if not date:
        continue

    daily = daily_by_date.get(date, {})

    sheet_name = safe_sheet_name(date.replace("/","")[:8] + "_" + r.get("氏名",""))
    ws = wb.copy_worksheet(tpl)
    ws.title = sheet_name

    ws[CELL_MAP["office"]].value = r.get("事業所名","")
    ws[CELL_MAP["date"]].value = date
    ws[CELL_MAP["user"]].value = r.get("氏名","")
    ws[CELL_MAP["time"]].value = normalize_time(r.get("実績開始時間",""), r.get("実績終了時間",""))
    ws[CELL_MAP["method"]].value = normalize_method(r.get("実績記録票備考欄",""))
    ws[CELL_MAP["program"]].value = build_program(daily)
    ws[CELL_MAP["dayreport"]].value = r.get("日報","")
    temp = daily.get("体温","")
    ws[CELL_MAP["temp"]].value = "未検温" if not temp else f"{temp}℃"
    ws[CELL_MAP["slack"]].value = (r.get("備考") or r.get("実績記録票備考欄") or "")[:1500]

out = Path("/content/output.xlsx")
wb.save(out)

print("完了: output.xlsx を生成しました")
files.download(out)
