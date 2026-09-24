"""
Utility functions for file lock handling and monthly archiving.
"""

import os
import sys
import time
import shutil
import glob
import re
import pandas as pd

try:
    if hasattr(sys.stdout, "reconfigure"):
        sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    if hasattr(sys.stderr, "reconfigure"):
        sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass


MONTH_MAP = {
    1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "May", 6: "Jun",
    7: "Jul", 8: "Aug", 9: "Sep", 10: "Oct", 11: "Nov", 12: "Dec"
}


def ensure_file_writable(filepath: str, description: str = "File Excel", max_retries: int = 15) -> str:
    """
    Ensure the target file is writable (not locked by Microsoft Excel or WPS Office).
    If locked, prompts the user to close it.
    Returns the writable filepath, or a fallback filepath if still locked after max retries.
    """
    if not os.path.exists(filepath):
        return filepath

    is_interactive = bool(sys.stdin and hasattr(sys.stdin, "isatty") and sys.stdin.isatty())
    effective_retries = max_retries if is_interactive else min(max_retries, 2)

    first_warn = True
    for attempt in range(effective_retries):
        try:
            with open(filepath, "r+b"):
                pass
            return filepath
        except (PermissionError, OSError):
            if first_warn:
                print("\n" + "=" * 70)
                print(f" [!] CẢNH BÁO / WARNING: {description}")
                print(f"     File '{os.path.basename(filepath)}' đang được mở trong Excel / WPS Office!")
                print(f"     Vui lòng ĐÓNG file Excel này lại để phần mềm có thể cập nhật dữ liệu mới.")
                print("=" * 70)
                first_warn = False

            try:
                if is_interactive:
                    input(f"   >> Sau khi đã ĐÓNG file '{os.path.basename(filepath)}', nhấn phím ENTER để tiếp tục... ")
                else:
                    time.sleep(1)
            except (EOFError, OSError):
                time.sleep(1)

    # If still locked after retries, create a safe fallback
    base, ext = os.path.splitext(filepath)
    fallback = f"{base}_New{ext}"
    print(f"\n[-] Chú ý: File '{os.path.basename(filepath)}' vẫn đang bị khóa bởi Excel.")
    print(f"    Dữ liệu được lưu tạm thời vào: '{os.path.basename(fallback)}'")
    print(f"    (Hãy đóng Excel và đổi tên file sau khi hoàn tất).\n")
    return fallback


def promote_fallback_if_writable(filepath: str) -> bool:
    """If a _New fallback file exists and target filepath is now writable, replace target with fallback."""
    base, ext = os.path.splitext(filepath)
    fallback = f"{base}_New{ext}"
    if os.path.exists(fallback):
        try:
            with open(filepath, "r+b"):
                pass
            shutil.copy2(fallback, filepath)
            os.remove(fallback)
            print(f"[+] Successfully promoted '{os.path.basename(fallback)}' -> '{os.path.basename(filepath)}'")
            return True
        except Exception:
            return False
    return False



def detect_file_month_and_year(filepath: str) -> tuple:
    """
    Detect month abbreviation (e.g. 'Jun') and 4-digit year (e.g. '2026')
    from file contents (Audit Date / InspDate) or filename.
    Returns: (month_str, year_str)
    """
    detected_month = None
    detected_year = None

    # 1. Read first rows from file to check date columns (highest priority)
    try:
        df_sample = pd.read_excel(filepath, nrows=50)
        date_cols = [c for c in df_sample.columns if "date" in str(c).lower()]
        for c in date_cols:
            dates = pd.to_datetime(df_sample[c], errors="coerce").dropna()
            if not dates.empty:
                first_date = dates.iloc[0]
                detected_month = MONTH_MAP.get(first_date.month)
                detected_year = str(first_date.year)
                if detected_month and detected_year:
                    return detected_month, detected_year
    except Exception:
        pass

    filename = os.path.basename(filepath)

    # 2. Try filename regex for YYYY-MM-DD or YYYY_MM_DD
    m = re.search(r"(\d{4})[-_](\d{2})[-_](\d{2})", filename)
    if m:
        y_val = m.group(1)
        month_num = int(m.group(2))
        if 1 <= month_num <= 12:
            return MONTH_MAP[month_num], y_val

    # 3. Try month names in filename (e.g. QAStation_Jul_...)
    for m_num, m_name in MONTH_MAP.items():
        if re.search(rf"[-_\s]{m_name}[-_\s\.]", filename, re.IGNORECASE):
            detected_month = m_name
            break

    # Look for 4-digit year in filename
    y_m = re.search(r"(20\d{2})", filename)
    if y_m:
        detected_year = y_m.group(1)

    return (detected_month or "Jun", detected_year or "2026")


def detect_file_month(filepath: str) -> str:
    """Detect month abbreviation (e.g. 'Jun') from file contents or filename."""
    m, _ = detect_file_month_and_year(filepath)
    return m


def detect_file_year(filepath: str) -> str:
    """Detect 4-digit year (e.g. '2026') from file contents or filename."""
    _, y = detect_file_month_and_year(filepath)
    return y


def auto_archive_previous_months(input_dir: str, active_month: str):
    """
    Ensure files in input_dir belonging to previous/other months are archived
    into input_dir/Archive/{Month}/ so only the active month's files remain in the root.
    """
    archive_base = os.path.join(input_dir, "Archive")

    # 1. Handle existing loose subfolders like 'Aug' that aren't 'Archive'
    for item in os.listdir(input_dir):
        item_path = os.path.join(input_dir, item)
        if os.path.isdir(item_path) and item.lower() != "archive":
            # If it's a month folder like 'Aug', move it into Archive/Aug
            dest_dir = os.path.join(archive_base, item)
            os.makedirs(dest_dir, exist_ok=True)
            for f in os.listdir(item_path):
                s_fp = os.path.join(item_path, f)
                d_fp = os.path.join(dest_dir, f)
                if os.path.isfile(s_fp):
                    shutil.move(s_fp, d_fp)
            try:
                os.rmdir(item_path)
            except Exception:
                pass

    # 2. Check root files
    files = glob.glob(os.path.join(input_dir, "*.xlsx")) + glob.glob(os.path.join(input_dir, "*.xls"))
    for fp in files:
        f_month = detect_file_month(fp)
        if f_month.lower() != active_month.lower():
            dest_dir = os.path.join(archive_base, f_month)
            os.makedirs(dest_dir, exist_ok=True)
            dest_file = os.path.join(dest_dir, os.path.basename(fp))
            try:
                shutil.move(fp, dest_file)
                print(f"   [+] Đã lưu trữ file tháng cũ '{os.path.basename(fp)}' vào: Archive/{f_month}/")
            except Exception as e:
                print(f"   [-] Không thể di chuyển file: {e}")
