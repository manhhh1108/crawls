"""
Them 2 row moi (Danh muc sp + Thuong hieu) vao tat ca file xlsx da crawl.

Lay danh muc tu cot E (Danh muc SP Tieng Viet) va F (Danh muc SP tieng Anh)
trong file SICK_Web_Products_Statistics.xlsx.

Brand: "SICK Sensor" cho tat ca.

Idempotent: chay nhieu lan khong gay loi (skip neu da co row "Danh muc sp").

Usage:
  python add_columns.py                    # chay tat ca
  python add_columns.py --family WTT12-S   # chi 1 dong san pham (test)
  python add_columns.py --dry-run          # preview, khong sua
"""
import os
import re
import sys
import glob
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

CHECKLIST = "SICK_Web_Products_Statistics.xlsx"
OUTPUT_DIR = "SICK_Products"
BRAND = "SICK Sensor"

# Same styling as original writeXlsxWithPython
LABEL_FILL = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
THIN = Side(border_style="thin", color="BFBFBF")
BORDER = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
BOLD = Font(bold=True)
WRAP_TOP = Alignment(wrap_text=True, vertical="top")


def safe_folder(name):
    """Match toSafeFilename() from crawl-sick-product.js."""
    s = re.sub(r"\s+", " ", str(name)).strip()
    s = re.sub(r"[^a-zA-Z0-9_-]+", "_", s)
    s = re.sub(r"^_+|_+$", "", s)
    return s


def load_category_map():
    """
    Doc checklist, build mapping: folder_name -> (vi_cat, en_cat).
    Neu folder xuat hien nhieu lan trong checklist (vd cung family code o nhieu category),
    UU TIEN giu lai gia tri co Danh muc SP (khong trong) thay vi gia tri trong sau cung.
    """
    wb = load_workbook(CHECKLIST, data_only=True)
    ws = wb.active
    mapping = {}
    for row in ws.iter_rows(min_row=3, values_only=True):
        if len(row) < 6:
            continue
        stt, name, code, count, vi_cat, en_cat = row[:6]
        if not stt or not name:
            continue
        s = str(stt)
        if s.count(".") < 2:  # chi leaf level
            continue
        folder = safe_folder(name)
        new_vi = str(vi_cat or "").strip()
        new_en = str(en_cat or "").strip()
        existing = mapping.get(folder)
        # Neu folder da co va lan nay trong, GIU lai gia tri cu (co data)
        if existing and existing[0] and not new_vi:
            continue
        mapping[folder] = (new_vi, new_en)
    return mapping


def add_rows_to_file(path, vi_cat, en_cat, dry_run=False):
    """
    Insert 2 rows tai vi tri 4 (sau row 'Ma', truoc row 'URL').
    Return: 'updated' | 'skipped' | 'value_updated' | 'error'

    Logic:
    - Neu row 4 chua co label 'Danh muc sp' -> insert 2 rows moi
    - Neu row 4 da co label nhung value TRONG va checklist co value -> chi update value
    - Neu row 4 da co label va value KHONG TRONG -> skip
    """
    wb = load_workbook(path)
    ws = wb.active

    if ws.cell(row=4, column=1).value == "Danh mục sp":
        # Da co rows. Kiem tra value co trong khong.
        current_vi = str(ws.cell(row=4, column=2).value or "").strip()
        if current_vi or not vi_cat:
            return "skipped"

        # Update value chi (khong insert)
        if dry_run:
            return "would_update_value"
        ws.cell(row=4, column=2, value=vi_cat)
        ws.cell(row=4, column=4, value=en_cat)
        wb.save(path)
        return "value_updated"

    if dry_run:
        return "would_update"

    # Insert 2 rows tai pos 4
    ws.insert_rows(4, 2)

    # Row 4: Danh muc sp
    ws.cell(row=4, column=1, value="Danh mục sp")
    ws.merge_cells(start_row=4, start_column=2, end_row=4, end_column=3)
    ws.cell(row=4, column=2, value=vi_cat)
    ws.merge_cells(start_row=4, start_column=4, end_row=4, end_column=5)
    ws.cell(row=4, column=4, value=en_cat)

    # Row 5: Thuong hieu
    ws.cell(row=5, column=1, value="Thương hiệu")
    ws.merge_cells(start_row=5, start_column=2, end_row=5, end_column=3)
    ws.cell(row=5, column=2, value=BRAND)
    ws.merge_cells(start_row=5, start_column=4, end_row=5, end_column=5)
    ws.cell(row=5, column=4, value=BRAND)

    # Apply formatting
    for r in (4, 5):
        for c in range(1, 6):
            cell = ws.cell(row=r, column=c)
            cell.border = BORDER
            cell.alignment = WRAP_TOP
        ws.cell(row=r, column=1).font = BOLD
        ws.cell(row=r, column=1).fill = LABEL_FILL

    wb.save(path)
    return "updated"


def main():
    dry_run = "--dry-run" in sys.argv
    family_filter = None
    if "--family" in sys.argv:
        idx = sys.argv.index("--family")
        if idx + 1 < len(sys.argv):
            family_filter = sys.argv[idx + 1]

    cat_map = load_category_map()
    print(f"Loaded {len(cat_map)} category mappings tu checklist")

    if family_filter:
        folders = [d for d in os.listdir(OUTPUT_DIR) if family_filter.lower() in d.lower()]
        print(f"Filter '--family {family_filter}': {len(folders)} folder(s)")
    else:
        folders = sorted(d for d in os.listdir(OUTPUT_DIR) if os.path.isdir(os.path.join(OUTPUT_DIR, d)))

    if dry_run:
        print("DRY RUN - khong sua file")

    total = 0
    updated = 0
    skipped = 0
    no_category = 0
    errors = 0
    no_cat_folders = []

    for folder in folders:
        folder_path = os.path.join(OUTPUT_DIR, folder)
        if not os.path.isdir(folder_path):
            continue

        cat = cat_map.get(folder)
        if not cat:
            # Folder hoan toan khong co trong checklist - skip
            no_category += 1
            no_cat_folders.append(folder)
            continue
        vi_cat, en_cat = cat
        # Neu cot E/F checklist trong, van them row nhung value de trong
        files = glob.glob(os.path.join(folder_path, "*.xlsx"))
        if not files:
            continue

        f_updated = 0
        f_value_updated = 0
        f_skipped = 0
        for f in files:
            total += 1
            try:
                result = add_rows_to_file(f, vi_cat, en_cat, dry_run=dry_run)
                if result in ("updated", "would_update"):
                    updated += 1
                    f_updated += 1
                elif result in ("value_updated", "would_update_value"):
                    f_value_updated += 1
                    updated += 1  # van count la 'updated' tong the
                elif result == "skipped":
                    skipped += 1
                    f_skipped += 1
            except Exception as e:
                print(f"  ERROR {f}: {e}")
                errors += 1

        action = "would update" if dry_run else "updated"
        val_act = "would fill value" if dry_run else "value filled"
        print(f"  {folder:30s} -> {action}: {f_updated:4d}, {val_act}: {f_value_updated:4d}, skipped: {f_skipped:4d}  [{vi_cat[:50]}]")

    print()
    print("=" * 70)
    print(f"Tong xlsx files: {total}")
    print(f"  {'Would update' if dry_run else 'Updated'}: {updated}")
    print(f"  Skipped (da co):  {skipped}")
    print(f"  Folders khong co danh muc: {no_category}")
    print(f"  Errors:    {errors}")

    if no_cat_folders and not family_filter:
        print()
        print("CANH BAO - Cac folder khong tim thay danh muc trong checklist:")
        for f in no_cat_folders[:10]:
            print(f"  {f}")
        if len(no_cat_folders) > 10:
            print(f"  ... va {len(no_cat_folders) - 10} folder khac")


if __name__ == "__main__":
    main()
