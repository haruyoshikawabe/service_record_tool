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
MSG_TEMPLATE_NOT_FOUND = "テンプレxlsxが見つかりません。exeと同じフォルダにテンプレxlsxを置いてください。"


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


def extract_yyyymm_from_filename(path: Path) -> Optional[str]:
    m = re.search(r"_(\d{6})(\d{2})?", path.name)
    return m.group(1) if m else None


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
        rows: List[Dict[str, str]] = []
        for r in reader:
            rows.append({(k or "").strip(): (v or "").strip() for k, v in r.items()})
        return rows


def normalize_date(s: str) -> str:
    s = (s or "").strip()
    if not s:
        return ""
    m = re.match(r"^\s*(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s*$", s)
    if m:
        y = int(m.group(1))
        mo = int(m.group(2))
        d = int(m.group(3))
        return f"{y}/{mo}/{d}"
    return s.replace("-", "/").strip()


def to_yyyymmdd(date_norm: str) -> str:
    """
    normalize_date 済み（例: 2026/1/8）を yyyymmdd（例: 20260108）へ。
    """
    m = re.match(r"^(\d{4})/(\d{1,2})/(\d{1,2})$", (date_norm or "").strip())
    if not m:
        return ""
    y = int(m.group(1))
    mo = int(m.group(2))
    d = int(m.group(3))
    return f"{y}{mo:02d}{d:02d}"


def safe_sheet_name(name: str) -> str:
    for c in [":", "/", "\\", "?", "*", "[", "]"]:
        name = name.replace(c, "_")
    return name.strip()[:31]


def parse_time_flexible(s: str) -> Optional[Tuple[int, int]]:
    s = (s or "").strip()
    if not s:
        return None
    patterns = [
        r"(\d{1,2}):(\d{2})(?::\d{2})?",
        r"(\d{1,2})時(\d{1,2})分",
    ]
    for pat in patterns:
        m = re.search(pat, s)
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

    def fmt_one(p: Optional[Tuple[int, int]], raw: str) -> str:
        if p:
            h, m = p
            return f"{h}時{m:02d}分"
        return (raw or "").strip()

    left = fmt_one(ps, start)
    right = fmt_one(pe, end)

    if not left and not right:
        return ""
    if left and right:
        return f"{left}～{right}"
    return left or right


def remove_sample_sheets(wb) -> None:
    targets = [name for name in wb.sheetnames if "sample" in name.lower()]
    for name in targets:
        del wb[name]


def pick_date_column(daily_rows: List[Dict[str, str]]) -> str:
    candidates = ["日付", "年月日", "支援実施日"]
    keys = list(daily_rows[0].keys())
    for c in candidates:
        if c in keys:
            return c
    return keys[0]


def pick_daily_contact_only(daily: Dict[str, str]) -> str:
    candidates = [
        "本人との連絡",
        "本人との連絡（チャット）",
        "本人との連絡（Slack）",
        "連絡事項",
        "連絡",
    ]
    for c in candidates:
        v = (daily.get(c) or "").strip()
        if v:
            return v
    return ""


def build_program(d: Dict[str, str]) -> str:
    out: List[str] = []

    def add(p, detail):
        p = (p or "").strip()
        detail = (detail or "").strip()
        if p or detail:
            out.append(p + ("\n" + detail if (p and detail) else detail))

    add(d.get("午前のプログラム", ""), d.get("午前のプログラム詳細", ""))
    add(d.get("午後1のプログラム", ""), d.get("午後1のプログラム詳細", ""))
    add(d.get("午後2のプログラム", ""), d.get("午後2のプログラム詳細", ""))
    add(d.get("終日のプログラム", ""), d.get("終日のプログラム詳細", ""))
    return "\n".join([x for x in out if x.strip()])


def normalize_method(raw: str) -> str:
    raw = raw or ""
    if "在宅" in raw and "通所" not in raw:
        return "利用者宅"
    return "事業所"


def format_contact_text(raw: str) -> str:
    text = (raw or "").strip()
    if not text:
        return ""
    token_pat = r"(\b\d{1,2}:\d{2}\b|\b\d{1,2}時間前\b)"
    parts = re.split(token_pat, text)
    if len(parts) == 1:
        body = text
        return (body[:30] + "・・・・") if len(body) > 30 else body

    lines: List[str] = []
    i = 0
    while i < len(parts):
        seg = parts[i] or ""
        if re.fullmatch(r"\b\d{1,2}:\d{2}\b|\b\d{1,2}時間前\b", seg):
            t = seg
            msg = (parts[i + 1] if i + 1 < len(parts) else "").strip()
            if len(msg) > 30:
                msg = msg[:30] + "・・・・"
            lines.append(f"{t} {msg}".rstrip())
            i += 2
        else:
            i += 1

    return "\n".join([ln for ln in lines if ln])


def ask_paths() -> Tuple[Optional[Path], Optional[Path], Optional[Path]]:
    root = tk.Tk()
    root.withdraw()

    user_path_str = filedialog.askopenfilename(
        title="userCaseDailyを選択",
        filetypes=[("CSV", "*.csv"), ("All files", "*.*")]
    )
    if not user_path_str:
        messagebox.showerror("エラー", MSG_USER_NOT_SELECTED)
        return None, None, None
    user_path = Path(user_path_str)

    if not is_csv(user_path):
        messagebox.showerror("エラー", MSG_NOT_CSV)
        return None, None, None
    if not looks_like_userCaseDaily(user_path):
        messagebox.showerror("エラー", MSG_NOT_USERCASEDAILY)
        return None, None, None

    case_path_str = filedialog.askopenfilename(
        title="caseMonth（またはcaseDaily）を選択",
        filetypes=[("CSV", "*.csv"), ("All files", "*.*")]
    )
    if not case_path_str:
        messagebox.showerror("エラー", MSG_CASE_NOT_SELECTED)
        return None, None, None
    case_path = Path(case_path_str)

    if not is_csv(case_path):
        messagebox.showerror("エラー", MSG_NOT_CSV)
        return None, None, None
    if not looks_like_caseMonth_or_caseDaily(case_path):
        messagebox.showerror("エラー", MSG_NOT_CASE_MONTH_DAILY)
        return None, None, None

    outdir_str = filedialog.askdirectory(title="出力先フォルダを選択")
    if not outdir_str:
        messagebox.showerror("エラー", MSG_OUTDIR_NOT_SELECTED)
        return None, None, None
    outdir = Path(outdir_str)

    return user_path, case_path, outdir


def ensure_same_month(user_path: Path, case_path: Path) -> str:
    u = extract_yyyymm_from_filename(user_path) or ""
    c = extract_yyyymm_from_filename(case_path) or ""
    if u and c and (u != c):
        raise ValueError(MSG_MONTH_MISMATCH)
    # 命名規則で必須なので、どちらか取れた方を返す（取れなければ後段で補完）
    return c or u


def build_output_filename(case_rows: List[Dict[str, str]], yyyymm: str) -> str:
    """
    命名規則：'名前'_'年月'_サービス支援記録.xlsx
    """
    name = (case_rows[0].get("氏名") or "").strip() or "名前未設定"
    return f"{name}_{yyyymm}_サービス支援記録.xlsx"


def load_template_or_fail(base: Path) -> Path:
    candidates = [
        base / "Sample_Format.xlsx",
        base / "Sample Format.xlsx",
        base / "サービス支援記録ーSample Format(河辺陽成).xlsx",
        base / "サービス支援記録-Sample Format(河辺陽成).xlsx",
    ]
    for p in candidates:
        if p.exists():
            return p
    for p in base.glob("*.xlsx"):
        return p
    raise FileNotFoundError(MSG_TEMPLATE_NOT_FOUND)


def generate(user_csv: Path, case_csv: Path, outdir: Path) -> Path:
    base = get_base_folder()
    template_path = load_template_or_fail(base)

    yyyymm = ensure_same_month(user_csv, case_csv)

    case_rows = read_csv_dicts(case_csv)
    daily_rows = read_csv_dicts(user_csv)
    if not case_rows:
        raise RuntimeError("caseMonth（caseDaily）が空です。")
    if not daily_rows:
        raise RuntimeError("userCaseDailyが空です。")

    # 年月がファイル名から取れない場合、caseMonth先頭日から補完
    if not yyyymm:
        d0 = normalize_date(case_rows[0].get("年月日", ""))
        m = re.match(r"^(\d{4})/(\d{1,2})", d0)
        yyyymm = f"{m.group(1)}{int(m.group(2)):02d}" if m else "YYYYMM"

    out_name = build_output_filename(case_rows, yyyymm)
    out_path = outdir / out_name

    if out_path.exists():
        msg = f"このフォルダーには『{out_name}』が存在します。上書きしますか？"
        if not messagebox.askyesno("確認", msg):
            raise RuntimeError("キャンセルしました。")

    try:
        wb = load_workbook(template_path)
    except (InvalidFileException, Exception) as e:
        raise RuntimeError(f"テンプレ読み込み失敗: {e}")

    remove_sample_sheets(wb)

    if TEMPLATE_SHEET not in wb.sheetnames:
        raise RuntimeError(f"テンプレに '{TEMPLATE_SHEET}' シートがありません。")
    tpl = wb[TEMPLATE_SHEET]

    date_col = pick_date_column(daily_rows)
    daily_by_date: Dict[str, Dict[str, str]] = {}
    for r in daily_rows:
        key = normalize_date(r.get(date_col, ""))
        if key:
            daily_by_date[key] = r

    required = ["事業所名", "氏名", "年月日", "出欠等", "実績開始時間", "実績終了時間"]
    for c in required:
        if c not in case_rows[0]:
            raise RuntimeError(f"caseMonth（caseDaily）に必須列がありません: {c}")

    for r in case_rows:
        status = (r.get("出欠等", "") or "").strip()
        if status == ABSENT_SKIP_VALUE:
            continue
        if status != ATTEND_VALUE:
            continue

        date_norm = normalize_date(r.get("年月日", ""))
        if not date_norm:
            continue

        yyyymmdd = to_yyyymmdd(date_norm)
        if not yyyymmdd:
            continue

        daily = daily_by_date.get(date_norm, {})

        # シート名：年月日(yyyymmdd)を必ず先頭にする
        person = (r.get("氏名", "") or "").strip()
        sheet_base = f"{yyyymmdd}_{person}"
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
        ws[CELL_MAP["date"]].value = date_norm
        ws[CELL_MAP["user"]].value = person

        ws[CELL_MAP["time"]].value = format_time_range_jp(
            r.get("実績開始時間", ""),
            r.get("実績終了時間", "")
        )

        method_cell = ws[CELL_MAP["method"]]
        method_cell.value = normalize_method(r.get("実績記録票備考欄", ""))
        method_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

        ws[CELL_MAP["program"]].value = build_program(daily)

        ws[CELL_MAP["dayreport"]].value = r.get("日報", "")

        temp = (daily.get("体温", "") or "").strip()
        ws[CELL_MAP["temp"]].value = "未検温" if temp == "" else f"{temp}℃"

        daily_contact = pick_daily_contact_only(daily)
        cm_note = (r.get("備考") or r.get("実績記録票備考欄") or "").strip()
        raw_contact = daily_contact or cm_note
        ws[CELL_MAP["slack"]].value = format_contact_text(raw_contact)

    remove_sample_sheets(wb)

    try:
        wb.save(out_path)
    except PermissionError:
        raise PermissionError(MSG_FILE_IN_USE)

    return out_path


def main():
    root = tk.Tk()
    root.withdraw()

    user_path, case_path, outdir = ask_paths()
    if user_path is None and case_path is None and outdir is None:
        return

    try:
        out_path = generate(user_path, case_path, outdir)
        messagebox.showinfo("完了", f"保存しました。\n{out_path}")
    except Exception as e:
        messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    main()
