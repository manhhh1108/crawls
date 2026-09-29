"""
Kiem tra tat ca file xlsx co dung 2 row 'Danh muc sp' va 'Thuong hieu' hay khong.

Phan loai:
  - OK: row 4 = 'Danh muc sp' co value, row 5 = 'Thuong hieu' co value
  - MISSING_ROWS: chua co row 4/5 (file format cu, can chay add_columns.py)
  - EMPTY_VALUE: co row 4/5 nhung value rong (folder chua co Danh muc trong checklist)
  - WRONG_LABEL: row 4/5 sai label (corrupted)

Usage:
  python check_columns.py             # kiem tra tat ca
  python check_columns.py --details   # in chi tiet 50 file dau cua moi loai
"""
import os
import sys
import glob
import time
from collections import defaultdict
from openpyxl import load_workbook

OUTPUT_DIR = "SICK_Products"


def check_file(path):
    """
    Tra ve: ('OK' | 'MISSING_ROWS' | 'EMPTY_VALUE' | 'WRONG_LABEL', detail)
    """
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active

        rows = []
        for r, row in enumerate(ws.iter_rows(min_row=1, max_row=10, values_only=True), 1):
            rows.append(row)

        if len(rows) < 5:
            wb.close()
            return ("MISSING_ROWS", f"file co {len(rows)} rows")

        row4 = rows[3]  # row index 4 (1-based)
        row5 = rows[4]  # row index 5

        label4 = str(row4[0] or "").strip() if row4 else ""
        label5 = str(row5[0] or "").strip() if row5 else ""

        wb.close()

        if label4 != "Danh mục sp" or label5 != "Thương hiệu":
            if "Danh mục sp" not in (label4, label5) and "Thương hiệu" not in (label4, label5):
                return ("MISSING_ROWS", f"row4='{label4}', row5='{label5}'")
            return ("WRONG_LABEL", f"row4='{label4}', row5='{label5}'")

        # Kiem tra value
        vi4 = str(row4[1] or "").strip() if row4 and len(row4) > 1 else ""
        en4 = str(row4[3] or "").strip() if row4 and len(row4) > 3 else ""
        vi5 = str(row5[1] or "").strip() if row5 and len(row5) > 1 else ""
        en5 = str(row5[3] or "").strip() if row5 and len(row5) > 3 else ""

        if not vi4 and not en4:
            return ("EMPTY_VALUE", "Danh muc sp trong (ca VI va EN)")
        if not vi5 and not en5:
            return ("EMPTY_VALUE", "Thuong hieu trong")

        return ("OK", "")
    except Exception as e:
        return ("ERROR", str(e))


def main():
    show_details = "--details" in sys.argv

    folders = sorted(d for d in os.listdir(OUTPUT_DIR) if os.path.isdir(os.path.join(OUTPUT_DIR, d)))
    print(f"Scanning {len(folders)} folders...")

    by_status = defaultdict(list)  # status -> [(folder, file, detail)]
    total = 0
    last_print = time.time()

    for fi, folder in enumerate(folders):
        folder_path = os.path.join(OUTPUT_DIR, folder)
        files = [f for f in glob.glob(os.path.join(folder_path, "*.xlsx"))
                 if not os.path.basename(f).startswith("~$")]
        for f in files:
            total += 1
            status, detail = check_file(f)
            by_status[status].append((folder, os.path.basename(f), detail))

            now = time.time()
            if now - last_print > 1 or total % 200 == 0:
                ok_count = len(by_status.get("OK", []))
                print(f"  [{total} files] OK: {ok_count} | issues: {total - ok_count} | now: {folder}", flush=True)
                last_print = now

    print()
    print("=" * 70)
    print(f"TONG: {total} files")
    print("=" * 70)
    for status in ["OK", "MISSING_ROWS", "EMPTY_VALUE", "WRONG_LABEL", "ERROR"]:
        items = by_status.get(status, [])
        if items:
            pct = len(items) * 100 // max(total, 1)
            print(f"  {status:15s}: {len(items):>6} files ({pct}%)")

    # Detail breakdown by folder for issues
    print()
    for status in ["MISSING_ROWS", "EMPTY_VALUE", "WRONG_LABEL", "ERROR"]:
        items = by_status.get(status, [])
        if not items:
            continue
        folder_count = defaultdict(int)
        for folder, _, _ in items:
            folder_count[folder] += 1
        print(f"\n=== {status} ({len(items)} files) - top folders ===")
        for folder, count in sorted(folder_count.items(), key=lambda x: -x[1])[:15]:
            print(f"  {count:>5}  {folder}")
        if show_details:
            print(f"  --- chi tiet {min(20, len(items))} file dau ---")
            for folder, fname, detail in items[:20]:
                print(f"    {folder}/{fname}  [{detail}]")


if __name__ == "__main__":
    main()
