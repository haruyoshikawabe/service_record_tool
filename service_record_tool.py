import csv
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils.exceptions import InvalidFileException

TEMPLATE_SHEET = "Format"

CELL_MAP = {
    "office": "B3",
    "date": "B4",
    "user": "G4",
    "time": "B5",
    "method": "G5",
    "program": "A9",
    "dayreport": "A11",
    "temp": "B13",
    "slack": "A16",
}

ATTEND_VALUE = "出席"
ABSENT_SKIP_VALUE = "欠席時対応"

MSG_NOT_USERCASEDAILY = "userCaseDailyではありません。"
MSG_NOT_CASE_MONTH_DAILY = "caseMonth（またはcaseDaily）ではありません。"
MSG_NOT_CSV = "csvファイルではありません。"
MSG_MONTH_MISMATCH = "userCaseDailyとcaseMonth（caseDaily）の年月が合いません。"
MSG_CASE_NOT_SELECTED = "caseMonth（またはcaseDaily）が未選択です。"
MSG_USER_NOT_SELECTED = "userCaseDailyが未選択です。"
MSG_OUTDIR_NOT_SELECTED = "出力先が未選択です。"
MSG_FILE_IN_USE = "ファイルにアクセスできません。別のプロセスが使用中です。"
MSG_TEMPLATE_NOT_FOUND = "テンプレxlsxが見つかりません。"


def get_base_folder() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def is_csv(path: Path) -> bool:
    return path.suffix.lower() == ".csv"


def looks_like_userCaseDaily(path: Path) -> bool:
    return "userCaseDaily" in path.name


def looks_like_caseMonth_or_caseDaily(path: Path) -> bool:
    name = path.name
    return ("caseMonth" in name) or ("caseDaily" in name)


def detect_encoding(path: Path) -> str:
    for enc in ("cp932", "shift_jis", "utf-8-sig", "utf-8"):
        try:
            with path.open("r", encoding=enc, newline="") as f:
                f.read(4096)
            return enc
        except Exception:
            pass
    raise RuntimeError("文字コード判定失敗")


def read_csv_dicts(path: Path) -> List[Dict[str, str]]:
    enc = detect_encoding(path)
    with path.open("r", encoding=enc, newline="") as f:
        reader = csv.DictReader(f)
        rows = []
        for r in reader:
            rows.append({(k or "").strip(): (v or "").strip() for k, v in r.items()})
        return rows


def normalize_date(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    m = re.match(r"^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$", s)
    if m:
        return f"{int(m.group(1))}/{int(m.group(2))}/{int(m.group(3))}"
    return s.replace("-", "/")


def to_yyyymmdd(date_norm: str) -> str:
    m = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", date_norm)
    if not m:
        return ""
    return f"{int(m.group(1))}{int(m.group(2)):02d}{int(m.group(3)):02d}"


def parse_time_flexible(s: str) -> Optional[Tuple[int, int]]:
    s = (s or "").strip()
    if not s:
        return None
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    m = re.search(r"(\d{1,2})時(\d{1,2})分", s)
    if m:
        return int(m.group(1)), int(m.group(2))
    return None


def format_time_range_jp(start: str, end: str) -> str:
    ps = parse_time_flexible(start)
    pe = parse_time_flexible(end)
    if ps and pe:
        return f"{ps[0]}時{ps[1]:02d}分～{pe[0]}時{pe[1]:02d}分"
    return ""


def pick_date_column(daily_rows):
    for c in ["日付", "年月日", "支援実施日"]:
        if c in daily_rows[0]:
            return c
    return list(daily_rows[0].keys())[0]


def build_program(d):
    parts = []
    keys = [
        ("午前のプログラム", "午前のプログラム詳細"),
        ("午後1のプログラム", "午後1のプログラム詳細"),
        ("午後2のプログラム", "午後2のプログラム詳細"),
        ("終日のプログラム", "終日のプログラム詳細"),
    ]
    for k1, k2 in keys:
        p = (d.get(k1) or "").strip()
        d2 = (d.get(k2) or "").strip()
        if p or d2:
            parts.append(p + ("\n" + d2 if p and d2 else d2))
    return "\n".join(parts)


def ask_paths():
    root = tk.Tk()
    root.withdraw()

    user = filedialog.askopenfilename(title="userCaseDailyを選択", filetypes=[("CSV", "*.csv")])
    if not user:
        messagebox.showerror("エラー", MSG_USER_NOT_SELECTED)
        return None, None, None
    user_path = Path(user)

    case = filedialog.askopenfilename(title="caseMonthを選択", filetypes=[("CSV", "*.csv")])
    if not case:
        messagebox.showerror("エラー", MSG_CASE_NOT_SELECTED)
        return None, None, None
    case_path = Path(case)

    out = filedialog.askdirectory(title="出力先を選択")
    if not out:
        messagebox.showerror("エラー", MSG_OUTDIR_NOT_SELECTED)
        return None, None, None

    return user_path, case_path, Path(out)


def load_template(base):
    for p in base.glob("*.xlsx"):
        return p
    raise FileNotFoundError(MSG_TEMPLATE_NOT_FOUND)


def generate(user_csv: Path, case_csv: Path, outdir: Path):

    case_rows = read_csv_dicts(case_csv)
    daily_rows = read_csv_dicts(user_csv)

    if not case_rows:
        raise RuntimeError("caseMonthが空です")

    # ===== 出力ファイル名確定 =====
    first_attend = None
    for r in case_rows:
        if (r.get("出欠等") or "").strip() == ATTEND_VALUE:
            first_attend = r
            break

    if first_attend is None:
        raise RuntimeError("出席データがありません")

    name_for_file = (first_attend.get("氏名") or "").strip()
    if not name_for_file:
        name_for_file = "名前未設定"

    d0 = normalize_date(first_attend.get("年月日", ""))
    ymd0 = to_yyyymmdd(d0)
    yyyymm = ymd0[:6] if ymd0 else "YYYYMM"

    out_name = f"{name_for_file}_{yyyymm}_サービス支援記録.xlsx"
    out_path = outdir / out_name
    # ============================

    template_path = load_template(get_base_folder())
    wb = load_workbook(template_path)
    tpl = wb[TEMPLATE_SHEET]

    date_col = pick_date_column(daily_rows)
    daily_by_date = {normalize_date(r.get(date_col, "")): r for r in daily_rows}

    for r in case_rows:
        if (r.get("出欠等") or "").strip() != ATTEND_VALUE:
            continue

        date_norm = normalize_date(r.get("年月日", ""))
        yyyymmdd = to_yyyymmdd(date_norm)
        if not yyyymmdd:
            continue

        ws = wb.copy_worksheet(tpl)
        ws.title = yyyymmdd

        daily = daily_by_date.get(date_norm, {})

        ws[CELL_MAP["office"]].value = r.get("事業所名", "")
        ws[CELL_MAP["date"]].value = date_norm
        ws[CELL_MAP["user"]].value = r.get("氏名", "")
        ws[CELL_MAP["time"]].value = format_time_range_jp(
            r.get("実績開始時間", ""),
            r.get("実績終了時間", "")
        )
        ws[CELL_MAP["method"]].value = r.get("実績記録票備考欄", "")
        ws[CELL_MAP["program"]].value = build_program(daily)
        ws[CELL_MAP["dayreport"]].value = r.get("日報", "")
        ws[CELL_MAP["temp"]].value = daily.get("体温", "")
        ws[CELL_MAP["slack"]].value = daily.get("本人との連絡", "")

    del wb[TEMPLATE_SHEET]

    wb.save(out_path)
    return out_path


def main():
    user_path, case_path, outdir = ask_paths()
    if not user_path:
        return
    try:
        out_path = generate(user_path, case_path, outdir)
        messagebox.showinfo("完了", f"保存しました\n{out_path}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    main()
