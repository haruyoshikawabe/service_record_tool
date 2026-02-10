import csv
from pathlib import Path
from typing import Dict, List
from openpyxl import load_workbook

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

def normalize_date(s: str) -> str:
    return (s or "").strip().replace("-", "/")

def normalize_time(start: str, end: str) -> str:
    start, end = (start or "").strip(), (end or "").strip()
    return "" if (not start and not end) else f"{start}～{end}"

def safe_sheet_name(name: str) -> str:
    for c in [":","/","\\","?","*","[","]"]:
        name = name.replace(c, "_")
    name = name.strip()
    return name[:31]

def detect_encoding(path: Path) -> str:
    for enc in ("cp932", "shift_jis", "utf-8-sig", "utf-8"):
        try:
            with path.open("r", encoding=enc, newline="") as f:
                f.read(4096)
            return enc
        except Exception:
            pass
    raise RuntimeError(f"文字コード判定失敗: {path}")

def read_csv_dicts(path: Path) -> List[Dict[str, str]]:
    enc = detect_encoding(path)
    with path.open("r", encoding=enc, newline="") as f:
        reader = csv.DictReader(f)
        return [{(k or "").strip(): (v or "").strip() for k,v in r.items()} for r in reader]

def build_program(d: Dict[str, str]) -> str:
    out: List[str] = []
    def add(p, detail):
        p = (p or "").strip()
        detail = (detail or "").strip()
        if p or detail:
            out.append(p + ("\n" + detail if (p and detail) else detail))
    add(d.get("午前のプログラム",""), d.get("午前のプログラム詳細",""))
    add(d.get("午後1のプログラム",""), d.get("午後1のプログラム詳細",""))
    add(d.get("午後2のプログラム",""), d.get("午後2のプログラム詳細",""))
    add(d.get("終日のプログラム",""), d.get("終日のプログラム詳細",""))
    return "\n".join(out)

def normalize_method(raw: str) -> str:
    raw = raw or ""
    if "在宅" in raw and "通所" not in raw:
        return "利用者宅"
    return "事業所"

def main():
    here = Path(__file__).resolve().parent

    templates = sorted(here.glob("*.xlsx"))
    case_files = sorted(here.glob("caseMonth_*.csv"))
    daily_files = sorted(here.glob("userCaseDaily_*.csv"))

    if not templates:
        raise RuntimeError("テンプレxlsxが見つかりません。exeと同じフォルダにテンプレxlsxを置いてください。")
    if not case_files:
        raise RuntimeError("caseMonth_*.csvが見つかりません。exeと同じフォルダに置いてください。")
    if not daily_files:
        raise RuntimeError("userCaseDaily_*.csvが見つかりません。exeと同じフォルダに置いてください。")

    template_path = templates[0]
    out_path = here / "output.xlsx"

    case_rows: List[Dict[str, str]] = []
    for p in case_files:
        case_rows.extend(read_csv_dicts(p))
    daily_rows: List[Dict[str, str]] = []
    for p in daily_files:
        daily_rows.extend(read_csv_dicts(p))

    if not case_rows:
        raise RuntimeError("caseMonthが空です。")
    if not daily_rows:
        raise RuntimeError("userCaseDailyが空です。")

    date_col = "日付" if "日付" in daily_rows[0] else list(daily_rows[0].keys())[0]
    daily_by_date: Dict[str, Dict[str, str]] = {}
    for r in daily_rows:
        daily_by_date[normalize_date(r.get(date_col, ""))] = r

    wb = load_workbook(template_path)
    if TEMPLATE_SHEET not in wb.sheetnames:
        raise RuntimeError(f"テンプレに '{TEMPLATE_SHEET}' シートがありません。")
    tpl = wb[TEMPLATE_SHEET]

    required = ["事業所名", "氏名", "年月日", "出欠等", "実績開始時間", "実績終了時間"]
    for c in required:
        if c not in case_rows[0]:
            raise RuntimeError(f"caseMonthに必須列がありません: {c}")

    created = 0
    for r in case_rows:
        if (r.get("出欠等", "") or "").strip() != ATTEND_VALUE:
            continue

        date = normalize_date(r.get("年月日", ""))
        if not date:
            continue

        daily = daily_by_date.get(date, {})

        sheet_base = f"{date.replace('/','')[:8]}_{(r.get('氏名','') or '').strip()}"
        sheet_name = safe_sheet_name(sheet_base)

        if sheet_name in wb.sheetnames:
            k = 2
            while True:
                cand = safe_sheet_name(f"{sheet_base}_{k}")
                if cand not in wb.sheetnames:
                    sheet_name = cand
                    break
                k += 1

        ws = wb.copy_worksheet(tpl)
        ws.title = sheet_name

        ws[CELL_MAP["office"]].value = r.get("事業所名", "")
        ws[CELL_MAP["date"]].value = date
        ws[CELL_MAP["user"]].value = r.get("氏名", "")
        ws[CELL_MAP["time"]].value = normalize_time(r.get("実績開始時間", ""), r.get("実績終了時間", ""))
        ws[CELL_MAP["method"]].value = normalize_method(r.get("実績記録票備考欄", ""))
        ws[CELL_MAP["program"]].value = build_program(daily)
        ws[CELL_MAP["dayreport"]].value = r.get("日報", "")

        temp = (daily.get("体温", "") or "").strip()
        ws[CELL_MAP["temp"]].value = "未検温" if temp == "" else f"{temp}℃"

        slack = (r.get("備考") or r.get("実績記録票備考欄") or "").strip()
        ws[CELL_MAP["slack"]].value = (slack[:1500] + "・・・・") if len(slack) > 1500 else slack

        created += 1

    wb.save(out_path)

if __name__ == "__main__":
    main()
