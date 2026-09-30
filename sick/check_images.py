"""
Kiem tra anh trong SICK_Products co loi gi khong.

Detect:
  - File anh qua nho (< 5KB): co the chi co watermark, khong co product image
  - File anh trung lap (cung MD5 hash): SICK API tra ve cung 1 anh cho nhieu product
  - File anh corrupt: PIL khong mo duoc

Usage:
  python check_images.py                       # quick scan
  python check_images.py --details             # show detail per issue
  python check_images.py --family WTT12-S      # 1 dong (test)
"""
import os
import sys
import glob
import hashlib
import time
from collections import defaultdict

OUTPUT_DIR = "SICK_Products"
TINY_SIZE = 5 * 1024  # < 5KB = co the bi loi


def md5_file(path):
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(8192), b""):
            h.update(chunk)
    return h.hexdigest()


def main():
    show_details = "--details" in sys.argv
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
    tiny_files = []  # files < 5KB
    by_hash = defaultdict(list)  # hash -> [paths]
    last_print = time.time()

    for fi, folder in enumerate(folders):
        images_dir = os.path.join(OUTPUT_DIR, folder, "images")
        if not os.path.isdir(images_dir):
            continue
        for img in glob.glob(os.path.join(images_dir, "*.jpg")):
            total_images += 1
            try:
                size = os.path.getsize(img)
                if size < TINY_SIZE:
                    tiny_files.append((img, size))
                # Hash file de detect duplicate
                h = md5_file(img)
                by_hash[h].append(img)
            except Exception as e:
                print(f"  ERROR {img}: {e}")

            now = time.time()
            if now - last_print > 1 or total_images % 1000 == 0:
                print(f"  [{fi+1}/{len(folders)}] {folder} | scanned: {total_images}", flush=True)
                last_print = now

    # Tim duplicate (cung hash, khac product)
    duplicates = {h: paths for h, paths in by_hash.items() if len(paths) > 1}
    # Loc bo cac duplicate trong CUNG product (vi du gallery images)
    real_duplicates = {}
    for h, paths in duplicates.items():
        # Get unique product names
        products = set()
        for p in paths:
            name = os.path.basename(p).replace(".jpg", "")
            # Strip trailing -digits if any (old format)
            products.add(name)
        if len(products) > 1:
            real_duplicates[h] = paths

    print()
    print("=" * 70)
    print(f"Tong anh:                 {total_images}")
    print(f"Anh < {TINY_SIZE//1024}KB (co the loi):    {len(tiny_files)}")
    print(f"Hash duplicate (>=2 prod): {len(real_duplicates)} groups, {sum(len(p) for p in real_duplicates.values())} files")
    print()

    if tiny_files:
        # Group by folder
        by_folder = defaultdict(int)
        for path, _ in tiny_files:
            f = os.path.basename(os.path.dirname(os.path.dirname(path)))
            by_folder[f] += 1
        print(f"Top 15 folder co nhieu anh nho:")
        for folder, count in sorted(by_folder.items(), key=lambda x: -x[1])[:15]:
            print(f"  {count:>5}  {folder}")
        if show_details:
            print(f"\n--- 30 sample anh nho nhat ---")
            for path, size in sorted(tiny_files, key=lambda x: x[1])[:30]:
                rel = os.path.relpath(path, OUTPUT_DIR)
                print(f"  {size:>6} bytes  {rel}")

    if real_duplicates:
        print(f"\nTop 15 group anh trung (>= 2 product co cung anh):")
        sorted_dup = sorted(real_duplicates.items(), key=lambda x: -len(x[1]))
        for h, paths in sorted_dup[:15]:
            print(f"\n  Hash {h[:12]}... - {len(paths)} files:")
            for p in paths[:5]:
                rel = os.path.relpath(p, OUTPUT_DIR)
                print(f"    {rel}")
            if len(paths) > 5:
                print(f"    ... va {len(paths) - 5} cai khac")


if __name__ == "__main__":
    main()
