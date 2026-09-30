"""
Rename file anh trong SICK_Products de bo phan partNumber.

Truoc:  WTT12S-C2569-1136898.jpg
Sau:    WTT12S-C2569.jpg

Logic:
- Voi moi folder, tim cac xlsx file (vd WTT12S-C2569.xlsx)
- Voi moi image trong folder/images, tim xlsx ma image bat dau bang xlsx_name + '-'
- Rename image thanh xlsx_name + '.jpg'
- Idempotent: skip neu image da co ten dung roi

Usage:
  python rename_images.py                    # rename tat ca
  python rename_images.py --dry-run          # preview
  python rename_images.py --family WTT12-S   # 1 dong (test)
"""
import os
import sys
import glob
import time

OUTPUT_DIR = "SICK_Products"


def main():
    dry_run = "--dry-run" in sys.argv
    family_filter = None
    if "--family" in sys.argv:
        idx = sys.argv.index("--family")
        if idx + 1 < len(sys.argv):
            family_filter = sys.argv[idx + 1]

    if family_filter:
        folders = [d for d in os.listdir(OUTPUT_DIR) if family_filter.lower() in d.lower()]
    else:
        folders = sorted(d for d in os.listdir(OUTPUT_DIR) if os.path.isdir(os.path.join(OUTPUT_DIR, d)))

    total_images = 0
    renamed = 0
    already_named = 0
    no_xlsx_match = 0
    conflicts = 0
    errors = 0
    last_print = time.time()

    for fi, folder in enumerate(folders):
        folder_path = os.path.join(OUTPUT_DIR, folder)
        images_dir = os.path.join(folder_path, "images")
        if not os.path.isdir(images_dir):
            continue

        # Lay danh sach xlsx names trong folder
        xlsx_names = set()
        for f in glob.glob(os.path.join(folder_path, "*.xlsx")):
            base = os.path.basename(f)
            if not base.startswith("~$"):
                xlsx_names.add(base.replace(".xlsx", ""))

        # Sap xep xlsx names theo do dai giam dan de match longest-prefix-first
        sorted_xlsx = sorted(xlsx_names, key=len, reverse=True)

        for img_path in glob.glob(os.path.join(images_dir, "*.jpg")):
            total_images += 1
            img_name = os.path.basename(img_path).replace(".jpg", "")

            # Image da match exact 1 xlsx -> da dung ten
            if img_name in xlsx_names:
                already_named += 1
                continue

            # Tim xlsx ma image bat dau bang xlsx_name + '-'
            matched = None
            for xname in sorted_xlsx:
                if img_name.startswith(xname + "-"):
                    matched = xname
                    break

            if not matched:
                no_xlsx_match += 1
                continue

            new_path = os.path.join(images_dir, f"{matched}.jpg")

            # Neu file dich da ton tai (tu lan rename truoc) -> conflict
            if os.path.exists(new_path) and os.path.abspath(new_path) != os.path.abspath(img_path):
                conflicts += 1
                continue

            if dry_run:
                renamed += 1
            else:
                try:
                    os.rename(img_path, new_path)
                    renamed += 1
                except Exception as e:
                    errors += 1
                    print(f"  ERROR rename {img_path}: {e}")

        now = time.time()
        if now - last_print > 1 or fi % 30 == 0:
            print(f"  [{fi+1}/{len(folders)}] {folder} | renamed: {renamed} | already: {already_named}", flush=True)
            last_print = now

    print()
    print("=" * 70)
    print(f"Tong images:                    {total_images}")
    if dry_run:
        print(f"Se rename:                      {renamed}")
    else:
        print(f"Da rename:                      {renamed}")
    print(f"Da co ten dung (skip):          {already_named}")
    print(f"Khong tim thay xlsx tuong ung:  {no_xlsx_match}")
    print(f"Conflict (file dich da ton tai): {conflicts}")
    print(f"Errors:                         {errors}")


if __name__ == "__main__":
    main()
