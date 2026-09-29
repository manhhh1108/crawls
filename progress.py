"""
Kiem tra tien do crawl SICK products.

Usage:
  python progress.py              # tong quan + top 20 dong chua xong
  python progress.py --all        # liet ke tat ca dong chua xong
  python progress.py --recent     # them: 10 file moi crawl gan day
  python progress.py --watch      # auto-refresh moi 30 giay
"""
import os
import sys
import glob
import time
from openpyxl import load_workbook

CHECKLIST = "SICK_Web_Products_Statistics.xlsx"
OUTPUT_DIR = "SICK_Products"


def safe_folder(name):
    """Match toSafeFilename() trong crawl-sick-product.js: thay sequence non-alnum bang 1 underscore."""
    import re
    s = re.sub(r"\s+", " ", str(name)).strip()
    s = re.sub(r"[^a-zA-Z0-9_-]+", "_", s)
    s = re.sub(r"^_+|_+$", "", s)
    return s


def collect_progress():
    """Doc checklist va doi chieu voi folder thuc te."""
    wb = load_workbook(CHECKLIST, data_only=True)
    ws = wb.active
    families = []  # (stt, name, actual, expected, has_folder)
    for row in ws.iter_rows(min_row=3, values_only=True):
        stt, series, code, expected = row[0], row[1], row[2], row[3]
        if not stt or not series:
            continue
        s = str(stt)
        if s.count('.') < 2:  # chi lay leaf-level (X.Y.Z)
            continue
        name = str(series).strip()
        folder = os.path.join(OUTPUT_DIR, safe_folder(name))
        has_folder = os.path.exists(folder)
        actual = len(glob.glob(os.path.join(folder, "*.xlsx"))) if has_folder else 0
        exp = int(expected) if expected else 0
        families.append((s, name, actual, exp, has_folder))
    return families


def total_disk_size():
    total = 0
    for root, _, files in os.walk(OUTPUT_DIR):
        for f in files:
            try:
                total += os.path.getsize(os.path.join(root, f))
            except OSError:
                pass
    return total


def fmt_size(b):
    for unit in ['B', 'KB', 'MB', 'GB']:
        if b < 1024:
            return f"{b:.1f} {unit}"
        b /= 1024
    return f"{b:.1f} TB"


def recent_files(n=10):
    """List N file xlsx moi nhat (theo mtime)."""
    files = []
    for path in glob.glob(os.path.join(OUTPUT_DIR, "*", "*.xlsx")):
        try:
            files.append((os.path.getmtime(path), path))
        except OSError:
            pass
    files.sort(reverse=True)
    return files[:n]


def render(families, args):
    total_done = sum(f[2] for f in families)
    total_expected = sum(f[3] for f in families)
    incomplete = [f for f in families if f[2] < f[3]]
    not_started = [f for f in incomplete if f[2] == 0]
    in_progress = [f for f in incomplete if f[2] > 0]
    completed = len(families) - len(incomplete)

    pct = total_done * 100 // max(total_expected, 1)
    images = len(glob.glob(os.path.join(OUTPUT_DIR, "*", "images", "*.jpg")))

    print("=" * 70)
    print(f"  SICK CRAWL PROGRESS  -  {time.strftime('%Y-%m-%d %H:%M:%S')}")
    print("=" * 70)
    print(f"  San pham:    {total_done:,} / {total_expected:,}  ({pct}%)")
    print(f"  Dong SP:     {len(families)} (xong: {completed} | dang crawl: {len(in_progress)} | chua bat dau: {len(not_started)})")
    print(f"  Anh:         {images:,} jpg")
    print(f"  Disk size:   {fmt_size(total_disk_size())}")
    print()

    show_n = len(incomplete) if args.get("all") else 20

    if in_progress:
        print(f"  >>> DONG DANG CRAWL DO ({len(in_progress)} dong):")
        # sort theo so SP con thieu (giam dan) - dong gan xong nhat truoc
        for stt, name, a, e, _ in sorted(in_progress, key=lambda x: x[2] - x[3]):
            pct_one = a * 100 // max(e, 1)
            bar_len = pct_one // 5  # 20 chars max
            bar = "#" * bar_len + "." * (20 - bar_len)
            print(f"    {a:5d}/{e:5d}  [{bar}] {pct_one:3d}%  {name}")
        print()

    if not_started:
        print(f"  >>> DONG CHUA BAT DAU ({len(not_started)} dong, top {min(show_n, len(not_started))} co nhieu SP nhat):")
        for stt, name, a, e, _ in sorted(not_started, key=lambda x: -x[3])[:show_n]:
            print(f"    {e:5d} SP  {name}")
        if len(not_started) > show_n and not args.get("all"):
            print(f"    ... va {len(not_started) - show_n} dong nua (chay '--all' de xem het)")
        print()

    if args.get("recent"):
        print("  >>> 10 FILE MOI NHAT:")
        for mtime, path in recent_files(10):
            ago = time.time() - mtime
            if ago < 60:
                ago_str = f"{int(ago)}s"
            elif ago < 3600:
                ago_str = f"{int(ago/60)}m"
            else:
                ago_str = f"{int(ago/3600)}h"
            rel = os.path.relpath(path, OUTPUT_DIR)
            print(f"    {ago_str:>5} ago  {rel}")
        print()

    if completed == len(families):
        print("  *** HOAN TAT TAT CA ***")


def main():
    args = {
        "all": "--all" in sys.argv,
        "recent": "--recent" in sys.argv,
        "watch": "--watch" in sys.argv,
    }
    if args["watch"]:
        while True:
            os.system("cls" if os.name == "nt" else "clear")
            try:
                render(collect_progress(), args)
            except Exception as e:
                print(f"Loi: {e}")
            print(f"\n(Auto-refresh moi 30s. Ctrl+C de dung.)")
            try:
                time.sleep(30)
            except KeyboardInterrupt:
                break
    else:
        render(collect_progress(), args)


if __name__ == "__main__":
    main()
