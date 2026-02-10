import csv
import re
import sys
from pathlib import Path
from typing import Dict, List, Optional, Tuple

from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException

import tkinter as tk
from tkinter import filedialog, messagebox

TEMPLATE_SHEET = "Format"

# 既存テンプレのセル配置（あなたの確定：A16）
CELL_MAP = {
    "office": "B3",
    "date": "B4",
    "user": "G4",
    "time": "B5",     # 対応時間
    "method": "G5",
    "program": "A9",
    "dayreport": "A11",
    "temp": "B13",
    "slack": "A16",   # 本人との連絡
}

ATTEND_VALUE = "出席"
ABSENT_SKIP_VALUE = "欠席時対応"  # テストケースNo.20

# テストケースに合わせた固定メッセージ
MSG_NOT_USERCASEDAILY = "userCaseDailyではありません。"
MSG_NOT_CASEDAILY = "caseDailyではありません。"
MSG_NOT_CSV = "csvファイルではありません。"
MSG_MONTH_MISMATCH = "userCaseDailyとcaseDailyの日時が合いません。"
MSG_CASE_NOT_SELECTED = "caseDailyが未選択です。"
MSG_USER_NOT_SELECTED = "userCaseDailyが未選択です。"
MSG_OUTDIR_NOT_SELECTED = "出力先が未選択です。"
MSG_FILE_IN_USE = "ファイルにアクセスできません。別のプロセスが使用中です。"

# テストケースNo.19の想定結果に寄せる（表記揺れもあるが、ここは合わせに行く）
MSG_TEMPLATE_NOT_FOUND = "java.io.FileNotFoundException.Sample_Format.xlsx(指定されたファイルが見つかりません。)"


def get_base_folder() -> Path:
    # PyInstaller(onefile)対策：exeのフォルダを基準にする
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def is_csv(path: Path) -> bool:
    return path.suffix.lower() == ".csv"


def looks_like_userCaseDaily(path: Path) -> bool:
    name = path.name
    return ("userCaseDaily" in name)


def looks_like_caseDaily(path: Path) -> bool:
    name = path.name
    # 現場で caseMonth_... という名前もあり得るが、テストケースは caseDaily を要求しているため
    # まずは caseDaily を優先。必要なら "caseMonth" を許容に変える。
    return ("caseDaily" in name) or ("caseMonth" in name)


def extract_yyyymm_from_filename(path: Path) -> Optional[str]:
    m = re.search(r"_(\d{6})", path.name)
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
    return (s or "").strip().replace("-", "/")


def parse_hhmm(s: str) -> Optional[Tuple[int, int]]:
    s = (s or "").strip()
    m = re.match(r"^(\d{1,2}):(\d{2})$", s)
    if not m:
        return None
    h = int(m.group(1))
    mi = int(m.group(2))
    return (h, mi)


def format_time_jp(start: str, end: str) -> str:
    ps = parse_hhmm(start)
    pe = parse_hhmm(end)
    if not ps or not pe:
        # 入力が HH:MM でない場合はそのまま “～” でつなぐ（落とさない）
        start = (start or "").strip()
        end = (end or "").strip()
        return "" if (not start and not end) else f"{start}～{end}"
    sh, sm = ps
    eh, em = pe
    # テストケースNo.15の「"時"分～"時"分」に合わせる
    return f"{sh}時{sm:02d}分～{eh}時{em:02d}分"


def safe_sheet_name(name: str) -> str:
    for c in [":", "/", "\\", "?", "*", "[", "]"]:
        name = name.replace(c, "_")
    return name.strip()[:31]


def pick_date_column(daily_rows: List[Dict[str, str]]) -> str:
    # userCaseDailyの日付列候補
    candidates = ["日付", "年月日", "支援実施日"]
    keys = list(daily_rows[0].keys())
    for c in candidates:
        if c in keys:
            return c
    return keys[0]


def pick_daily_note(daily: Dict[str, str]) -> str:
    # userCaseDailyの備考列候補
    candidates = ["備考", "備考欄", "本人との連絡", "連絡", "連絡事項"]
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
    return "\n".join(out)


def normalize_method(raw: str) -> str:
    raw = raw or ""
    if "在宅" in raw and "通所" not in raw:
        return "利用者宅"
    return "事業所"


def format_contact_text(raw: str) -> str:
    """
    テストケースNo.16対応：
    - 時刻(HH:MM)単位で改行
    - 各行は「HH:MM 」＋本文
    - 時刻の後ろ30文字以降は「・・・・」で省略表示
    """
    text = (raw or "").strip()
    if not text:
        return ""

    # まず時刻トークンで分割（時刻を保持）
    # 例: "10:35 xxx 11:10 yyy" -> ["", "10:35", " xxx ", "11:10", " yyy"]
    parts = re.split(r"(\b\d{1,2}:\d{2}\b)", text)
    if len(parts) == 1:
        # 時刻が無い場合は全体を30文字省略ルールだけ適用
        body = text
        return (body[:30] + "・・・・") if len(body) > 30 else body

    lines: List[str] = []
    i = 0
    while i < len(parts):
        seg = parts[i]
        if re.fullmatch(r"\b\d{1,2}:\d{2}\b", seg or ""):
            t = seg
            msg = (parts[i + 1] if i + 1 < len(parts) else "").strip()
            # 30文字で省略
            disp = msg
            if len(disp) > 30:
                disp = disp[:30] + "・・・・"
            lines.append(f"{t} {disp}".rstrip())
            i += 2
        else:
            i += 1

    # 時刻ごとに改行
    return "\n".join([ln for ln in lines if ln])


def ask_paths() -> Tuple[Optional[Path], Optional[Path], Optional[Path]]:
    root = tk.Tk()
    root.withdraw()

    # userCaseDaily選択
    user_path_str = filedialog.askopenfilename(title="userCaseDailyを選択", filetypes=[("CSV", "*.*")])
    if not user_path_str:
        return None, None, None
    user_path = Path(user_path_str)

    if not is_csv(user_path):
        messagebox.showerror("エラー", MSG_NOT_CSV)
        return None, None, None
    if not looks_like_userCaseDaily(user_path):
        messagebox.showerror("エラー", MSG_NOT_USERCASEDAILY)
        return None, None, None

    # caseDaily選択
    case_path_str = filedialog.askopenfilename(title="caseDailyを選択", filetypes=[("CSV", "*.*")])
    if not case_path_str:
        # userは選んだがcaseは未選択
        messagebox.showerror("エラー", MSG_CASE_NOT_SELECTED)
        return None, None, None
    case_path = Path(case_path_str)

    if not is_csv(case_path):
        messagebox.showerror("エラー", MSG_NOT_CSV)
        return None, None, None
    if not looks_like_caseDaily(case_path):
        messagebox.showerror("エラー", MSG_NOT_CASEDAILY)
        return None, None, None

    # 出力先フォルダ
    outdir_str = filedialog.askdirectory(title="出力先フォルダを選択")
    if not outdir_str:
        messagebox.showerror("エラー", MSG_OUTDIR_NOT_SELECTED)
        return None, None, None
    outdir = Path(outdir_str)

    return user_path, case_path, outdir


def ensure_same_month(user_path: Path, case_path: Path) -> None:
    u = extract_yyyymm_from_filename(user_path)
    c = extract_yyyymm_from_filename(case_path)
    if u and c and (u != c):
        raise ValueError(MSG_MONTH_MISMATCH)


def build_output_filename(case_rows: List[Dict[str, str]], yyyymm: Optional[str]) -> str:
    # テストケースNo.6: ('名前'_'年月'_サービス支援記録.xlsx)
    name = (case_rows[0].get("氏名") or "").strip() or "名前未設定"
    if not yyyymm:
        # 年月が取れない場合は先頭日付から生成
        d = normalize_date(case_rows[0].get("年月日", ""))
        m = re.match(r"^(\d{4})/(\d{1,2})", d)
        yyyymm = f"{m.group(1)}{int(m.group(2)):02d}" if m else "YYYYMM"
    return f"{name}_{yyyymm}_サービス支援記録.xlsx"


def load_template_or_fail(base: Path) -> Path:
    tpl = base / "Sample_Format.xlsx"
    if not tpl.exists():
        raise FileNotFoundError(MSG_TEMPLATE_NOT_FOUND)
    return tpl


def generate(user_csv: Path, case_csv: Path, outdir: Path) -> Path:
    base = get_base_folder()
    template_path = load_template_or_fail(base)

    # 月一致チェック（テストケースNo.9）
    ensure_same_month(user_csv, case_csv)

    # CSV読み込み
    case_rows = read_csv_dicts(case_csv)
    daily_rows = read_csv_dicts(user_csv)
    if not case_rows:
        raise RuntimeError("caseDailyが空です。")
    if not daily_rows:
        raise RuntimeError("userCaseDailyが空です。")

    yyyymm = extract_yyyymm_from_filename(case_csv) or extract_yyyymm_from_filename(user_csv)
    out_name = build_output_filename(case_rows, yyyymm)
    out_path = outdir / out_name

    # 上書き確認（テストケースNo.13）
    if out_path.exists():
        # 表の文言に寄せる（引用符が独特だが、ここは日本語メッセージで揃える）
        msg = f"このフォルダーには’{out_name}’は存在します。上書きしますか？"
        if not messagebox.askyesno("確認", msg):
            raise RuntimeError("キャンセルしました。")

    # テンプレ読み込み
    try:
        wb = load_workbook(template_path)
    except (InvalidFileException, Exception) as e:
        raise RuntimeError(f"テンプレ読み込み失敗: {e}")

    if TEMPLATE_SHEET not in wb.sheetnames:
        raise RuntimeError(f"テンプレに '{TEMPLATE_SHEET}' シートがありません。")
    tpl = wb[TEMPLATE_SHEET]

    # Sampleシート削除（テストケースNo.18）
    if "Sample" in wb.sheetnames:
        del wb["Sample"]

    # userCaseDailyの日付列
    date_col = pick_date_column(daily_rows)
    daily_by_date: Dict[str, Dict[str, str]] = {}
    for r in daily_rows:
        daily_by_date[normalize_date(r.get(date_col, ""))] = r

    # 必須列チェック（最低限）
    required = ["事業所名", "氏名", "年月日", "出欠等", "実績開始時間", "実績終了時間"]
    for c in required:
        if c not in case_rows[0]:
            raise RuntimeError(f"caseDailyに必須列がありません: {c}")

    created = 0
    for r in case_rows:
        status = (r.get("出欠等", "") or "").strip()

        # 出席のみ作る（テストケースNo.14/20）
        if status == ABSENT_SKIP_VALUE:
            continue
        if status != ATTEND_VALUE:
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

        # 対応時間（テストケースNo.15）
        ws[CELL_MAP["time"]].value = format_time_jp(r.get("実績開始時間", ""), r.get("実績終了時間", ""))

        ws[CELL_MAP["method"]].value = normalize_method(r.get("実績記録票備考欄", ""))
        ws[CELL_MAP["program"]].value = build_program(daily)
        ws[CELL_MAP["dayreport"]].value = r.get("日報", "")

        # 体温（テストケースNo.17）
        temp = (daily.get("体温", "") or "").strip()
        ws[CELL_MAP["temp"]].value = "未検温" if temp == "" else f"{temp}℃"

        # 本人との連絡（テストケースNo.16）
        # daily備考優先 → なければcase側備考
        daily_note = pick_daily_note(daily)
        cm_note = (r.get("備考") or r.get("実績記録票備考欄") or "").strip()
        raw_contact = daily_note or cm_note
        ws[CELL_MAP["slack"]].value = format_contact_text(raw_contact)

        created += 1

    # 保存（テストケースNo.5/7/8）
    try:
        wb.save(out_path)
    except PermissionError:
        raise PermissionError(MSG_FILE_IN_USE)

    return out_path


def main():
    root = tk.Tk()
    root.withdraw()

    # ファイル未選択系（テストケースNo.10/11/12）
    user_path, case_path, outdir = ask_paths()
    if user_path is None and case_path is None and outdir is None:
        # ask_paths内でエラー表示済み or キャンセル済み
        return
    if case_path is None:
        messagebox.showerror("エラー", MSG_CASE_NOT_SELECTED)
        return
    if user_path is None:
        messagebox.showerror("エラー", MSG_USER_NOT_SELECTED)
        return
    if outdir is None:
        messagebox.showerror("エラー", MSG_OUTDIR_NOT_SELECTED)
        return

    try:
        out_path = generate(user_path, case_path, outdir)
        messagebox.showinfo("完了", f"保存しました。\n{out_path}")
    except FileNotFoundError as e:
        messagebox.showerror("エラー", str(e))
    except ValueError as e:
        # 月不一致など
        messagebox.showerror("エラー", str(e))
    except PermissionError as e:
        messagebox.showerror("エラー", str(e))
    except Exception as e:
        messagebox.showerror("エラー", str(e))


if __name__ == "__main__":
    main()
