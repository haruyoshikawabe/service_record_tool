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
    "method": "G5",     # ← 中央揃え対象
    "program": "A9",
    "dayreport": "A11",
    "temp": "B13",
    "slack": "A16",
}

ATTEND_VALUE = "出席"
ABSENT_SKIP_VALUE = "欠席時対応"

MSG_TEMPLATE_NOT_FOUND = "java.io.FileNotFoundException.Sample_Format.xlsx(指定されたファイルが見つかりません。)"


# ===============================
# 対応時間フォーマット
# ===============================

def parse_time_flexible(s: str) -> Optional[Tuple[int, int]]:
    s = (s or "").strip()
    if not s:
        return None
    m = re.search(r"(\d{1,2}):(\d{2})", s)
    if m:
        h = int(m.group(1))
        mi = int(m.group(2))
        if 0 <= h <= 23 and 0 <= mi <= 59:
            return (h, mi)
    return None


def format_time_range_jp(start: str, end: str) -> str:
    ps = parse_time_flexible(start)
    pe = parse_time_flexible(end)

    if ps and pe:
        sh, sm = ps
        eh, em = pe
        return f"{sh}時{sm:02d}分～{eh}時{em:02d}分"

    return f"{start}～{end}"


# ===============================
# Sampleシート削除
# ===============================

def remove_sample_sheets(wb):
    targets = [s for s in wb.sheetnames if "sample" in s.lower()]
    for t in targets:
        del wb[t]


# ===============================
# メイン処理
# ===============================

def generate(user_csv: Path, case_csv: Path, outdir: Path) -> Path:

    template_path = Path(sys.executable).parent / "Sample_Format.xlsx"
    if not template_path.exists():
        raise FileNotFoundError(MSG_TEMPLATE_NOT_FOUND)

    case_rows = list(csv.DictReader(open(case_csv, encoding="cp932")))
    daily_rows = list(csv.DictReader(open(user_csv, encoding="cp932")))

    wb = load_workbook(template_path)
    remove_sample_sheets(wb)

    tpl = wb[TEMPLATE_SHEET]

    for r in case_rows:
        if r.get("出欠等") != ATTEND_VALUE:
            continue

        ws = wb.copy_worksheet(tpl)
        ws.title = r.get("年月日").replace("/", "")[:8]

        ws[CELL_MAP["office"]].value = r.get("事業所名")
        ws[CELL_MAP["date"]].value = r.get("年月日")
        ws[CELL_MAP["user"]].value = r.get("氏名")

        # 対応時間
        ws[CELL_MAP["time"]].value = format_time_range_jp(
            r.get("実績開始時間"),
            r.get("実績終了時間")
        )

        # 対応手段（G5）＋中央揃え
        method_cell = ws[CELL_MAP["method"]]
        method_cell.value = r.get("実績記録票備考欄")
        method_cell.alignment = Alignment(horizontal="center", vertical="center")

        ws[CELL_MAP["dayreport"]].value = r.get("日報")

    remove_sample_sheets(wb)

    out_path = outdir / "output.xlsx"
    wb.save(out_path)

    return out_path


def main():
    root = tk.Tk()
    root.withdraw()

    user_path = filedialog.askopenfilename(title="userCaseDailyを選択")
    case_path = filedialog.askopenfilename(title="caseDailyを選択")
    outdir = filedialog.askdirectory(title="出力先フォルダ")

    if not user_path or not case_path or not outdir:
        return

    generate(Path(user_path), Path(case_path), Path(outdir))
    messagebox.showinfo("完了", "保存しました。")


if __name__ == "__main__":
    main()
