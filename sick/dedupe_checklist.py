"""
Xoa cac dong duplicate trong SICK_Web_Products_Statistics.xlsx (theo Category Code).

Quy tac giu lai (theo do uu tien):
  1. Neu chi 1 dong co data o cot E (Danh muc SP Tieng Viet) hoac F (Tieng Anh) -> giu dong do
  2. Neu nhieu dong cung co data -> giu dong dau tien (so row nho nhat)
  3. Neu khong dong nao co data -> giu dong dau tien

QUAN TRONG:
  - KHONG tu y them/sua data o cot E va F
  - Backup file goc ra .bak truoc khi sua
  - Idempotent: chay nhieu lan an toan

Usage:
  python dedupe_checklist.py --dry-run    # preview se xoa nhung gi
  python dedupe_checklist.py              # thuc hien xoa
"""
import sys
import shutil
from collections import defaultdict
from openpyxl import load_workbook

CHECKLIST = "SICK_Web_Products_Statistics.xlsx"


def has_data(value):
    return bool(str(value or "").strip())


def main():
    dry_run = "--dry-run" in sys.argv

    if not dry_run:
        backup = CHECKLIST + ".bak"
        shutil.copy(CHECKLIST, backup)
        print(f"Backup goc: {backup}")
        print()

    wb = load_workbook(CHECKLIST)
    ws = wb.active

    # Group rows by Category Code (cot C)
    by_code = defaultdict(list)  # code -> [(row_idx, stt, name, vi_cat, en_cat)]
    for r in range(3, ws.max_row + 1):
        stt = ws.cell(row=r, column=1).value
        name = ws.cell(row=r, column=2).value
        code = ws.cell(row=r, column=3).value
        vi_cat = ws.cell(row=r, column=5).value
        en_cat = ws.cell(row=r, column=6).value

        if not stt or not name or not code:
            continue
        s = str(stt)
        if s.count(".") < 2:  # leaf only
            continue

        code_clean = str(code).strip().lower()
        by_code[code_clean].append((r, str(stt), str(name).strip(), vi_cat, en_cat))

    # Tim cac group co duplicate
    duplicates = {k: v for k, v in by_code.items() if len(v) > 1}
    print(f"Tim thay {len(duplicates)} family code bi duplicate.")
    print()

    rows_to_delete = []  # list of (row_idx, stt, name, reason)
    rows_to_keep = []    # list of (row_idx, stt, name, reason)

    for code, rows in duplicates.items():
        # Quyet dinh row nao giu
        rows_with_data = [r for r in rows if has_data(r[3]) or has_data(r[4])]

        if len(rows_with_data) == 1:
            keeper = rows_with_data[0]
            reason_keep = "co data, dong khac trong"
        elif len(rows_with_data) > 1:
            keeper = rows_with_data[0]  # giu dong dau co data
            reason_keep = f"co data, dong dau (cua {len(rows_with_data)} dong co data)"
        else:
            keeper = rows[0]  # khong co dong nao co data, giu dong dau
            reason_keep = "khong dong nao co data, giu dong dau"

        rows_to_keep.append((keeper[0], keeper[1], keeper[2], reason_keep))

        for r in rows:
            if r[0] == keeper[0]:
                continue
            reason_del = "duplicate, da co dong khac giu lai"
            rows_to_delete.append((r[0], r[1], r[2], reason_del))

    # Sap xep theo row giam dan de delete tu duoi len (tranh shift index)
    rows_to_delete.sort(key=lambda x: -x[0])

    # In ke hoach
    print("=" * 90)
    print(f"GIU LAI ({len(rows_to_keep)} dong):")
    print("=" * 90)
    for r, stt, name, reason in sorted(rows_to_keep, key=lambda x: x[0]):
        print(f"  Row {r:>4} STT={stt:9s} {name[:40]:40s}  [{reason}]")

    print()
    print("=" * 90)
    print(f"XOA ({len(rows_to_delete)} dong):")
    print("=" * 90)
    for r, stt, name, reason in sorted(rows_to_delete, key=lambda x: x[0]):
        print(f"  Row {r:>4} STT={stt:9s} {name[:40]:40s}  [{reason}]")

    print()
    if dry_run:
        print(f"DRY RUN - khong xoa. Bo --dry-run de thuc hien.")
        return

    # Thuc hien xoa - tu duoi len
    print(f"Dang xoa {len(rows_to_delete)} dong...")
    for r, _, _, _ in rows_to_delete:
        ws.delete_rows(r, 1)

    wb.save(CHECKLIST)
    print(f"Da xoa va luu: {CHECKLIST}")
    print(f"Backup: {CHECKLIST}.bak")


if __name__ == "__main__":
    main()
