import os
import re
import sys
import glob
from collections import defaultdict
from openpyxl import load_workbook

OUTPUT_DIR = "SICK_Products"

# Vietnamese diacritics + unique chars
VI_CHARS = (
    "ăâđêôơưĂÂĐÊÔƠƯ"
    "áàảãạắằẳẵặấầẩẫậ"
    "ÁÀẢÃẠẮẰẲẴẶẤẦẨẪẬ"
    "éèẻẽẹếềểễệ"
    "ÉÈẺẼẸẾỀỂỄỆ"
    "íìỉĩị"
    "ÍÌỈĨỊ"
    "óòỏõọốồổỗộớờởỡợ"
    "ÓÒỎÕỌỐỒỔỖỘỚỜỞỠỢ"
    "úùủũụứừửữự"
    "ÚÙỦŨỤỨỪỬỮỰ"
    "ýỳỷỹỵ"
    "ÝỲỶỸỴ"
)
VI_RE = re.compile(f"[{re.escape(VI_CHARS)}]")

# Tu tieng Anh thong dung trong technical text - khong tinh la "translatable" (tu chuyen nganh giu nguyen)
KEEP_WORDS = {
    "led", "iec", "iso", "din", "pvc", "abs", "pmma", "pom", "pet", "tpe", "pa", "pc",
    "pnp", "npn", "rfid", "nfc", "usb", "rs", "sil", "pl", "mttf", "mttfd", "ppr",
    "atex", "iecex", "ce", "csa", "eac", "ccc", "ul", "fcc", "fda", "rohs", "reach",
    "din", "gl", "dnv", "aplikon",
    "mm", "cm", "kg", "rpm", "lx", "psi", "bar", "min", "max", "typ", "avg",
    "and", "the", "for", "with", "without", "from", "into", "onto",  # common but if many = English
}

# Phrases tieng Anh chac chan KHONG xuat hien trong tieng Viet
ENGLISH_GIVEAWAY_PHRASES = [
    r"\bthe\s+laser\b",
    r"\bthe\s+target\b",
    r"\bthe\s+light\b",
    r"\bthe\s+sensor\b",
    r"\bthe\s+device\b",
    r"\bthe\s+product\b",
    r"\bdistance\s+front\s+of\s+sensor\b",
    r"\b(at|on|in|with|of|for|to|by)\s+the\b",
    r"\bcovered\s+fully\b",
    r"\bbe\s+used\b",
    r"\b(a|an)\s+\w+\s+of\b",
    r"\bcan\s+be\b",
    r"\bare\s+\w+ed\b",
    r"\b\w+ing\s+(of|the|a|on)\b",
]
ENGLISH_GIVEAWAY_RE = re.compile("|".join(ENGLISH_GIVEAWAY_PHRASES), re.IGNORECASE)


def has_untranslated_english(text):
    """Tra ve True neu text co cum tieng Anh dai khong duoc dich."""
    if not text:
        return False
    s = str(text)
    if len(s) < 30:
        return False

    # Phat hien nhanh: phrase tieng Anh chac chan
    if ENGLISH_GIVEAWAY_RE.search(s):
        return True

    # Heuristic: tach text theo ky tu Vietnamese
    # Trong moi chunk khong chua VN char, dem tu English co 4+ chu cai khong phai unit
    chunks = VI_RE.split(s)
    for chunk in chunks:
        # Loai bo so + don vi truoc khi count words
        c = re.sub(r"\d+[.,]?\d*\s*[A-Za-z°/µΩ]+\b", " ", chunk)
        c = re.sub(r"\d+", " ", c)
        words = re.findall(r"[a-zA-Z]{3,}", c)
        # Loc bo cac tu giu nguyen (unit/code/abbrev)
        meaningful = [
            w for w in words
            if w.lower() not in KEEP_WORDS and not w.isupper()  # bo qua tu viet HOA toan bo (likely codes)
        ]
        if len(meaningful) >= 5:
            return True

    return False


def check_file(path):
    """Doc file xlsx, tra ve list cac van de phat hien duoc."""
    issues = []
    try:
        wb = load_workbook(path, data_only=True, read_only=True)
        ws = wb.active

        # Doc tat ca cells
        for row in ws.iter_rows(values_only=True):
            if not row or len(row) < 5:
                continue
            label = str(row[0] or "").strip()
            vi = row[1]
            en = row[3]

            # Mo ta ngan cua san pham (formerly "Title mo ta ngan")
            if label in ("Mô tả ngắn của sản phẩm", "Title mô tả ngắn"):
                if vi and has_untranslated_english(vi):
                    issues.append(("Title mô tả ngắn", str(vi)[:200]))

            # Detail spec rows - col B = name VI, col C = value VI, col D = name EN, col E = value EN
            # Format: row = (col_a, col_b, col_c, col_d, col_e)
            if not label:  # detail spec row (col A empty)
                # row[2] is col C (valueVi), row[4] is col E (valueEn)
                if len(row) >= 5:
                    val_vi = row[2]
                    val_en = row[4]
                    if val_vi and val_en and has_untranslated_english(val_vi):
                        issues.append((str(row[1] or "")[:50], str(val_vi)[:150]))
        wb.close()
    except Exception as e:
        issues.append(("ERROR", str(e)))
    return issues


def main():
    delete_mode = "--delete" in sys.argv
    sample = 0
    if "--sample" in sys.argv:
        idx = sys.argv.index("--sample")
        if idx + 1 < len(sys.argv):
            sample = int(sys.argv[idx + 1])

    family_filter = None
    if "--family" in sys.argv:
        idx = sys.argv.index("--family")
        if idx + 1 < len(sys.argv):
            family_filter = sys.argv[idx + 1]

    if family_filter:
        folders = [d for d in os.listdir(OUTPUT_DIR) if family_filter.lower() in d.lower()]
    else:
        folders = sorted(d for d in os.listdir(OUTPUT_DIR) if os.path.isdir(os.path.join(OUTPUT_DIR, d)))

    total_files = 0
    bad_files = []
    error_files = []
    by_family = defaultdict(int)
    sample_shown = 0

    # Pre-count tong so file de hien % progress
    print(f"Scanning {len(folders)} folders...", flush=True)
    folder_files = []
    for folder in folders:
        folder_path = os.path.join(OUTPUT_DIR, folder)
        if not os.path.isdir(folder_path):
            continue
        files = [f for f in glob.glob(os.path.join(folder_path, "*.xlsx"))
                 if not os.path.basename(f).startswith("~$")]
        if files:
            folder_files.append((folder, files))
    grand_total = sum(len(fs) for _, fs in folder_files)
    print(f"Total {grand_total} xlsx files de check.\n", flush=True)

    import time as _time
    last_print = _time.time()

    for folder, files in folder_files:
        folder_bad = 0
        for f in files:
            total_files += 1
            issues = check_file(f)
            real_issues = [(t, txt) for t, txt in issues if t != "ERROR"]
            errors = [(t, txt) for t, txt in issues if t == "ERROR"]
            if errors:
                error_files.append((f, errors))
                continue
            if real_issues:
                bad_files.append((f, real_issues))
                by_family[folder] += 1
                folder_bad += 1
                if sample > 0 and sample_shown < sample:
                    print(f"\n[SAMPLE] {f}")
                    for tag, text in real_issues[:3]:
                        print(f"   [{tag}] {text}")
                    sample_shown += 1

            # Progress indicator: in moi 1 giay HOAC moi 100 files
            now = _time.time()
            if now - last_print > 1 or total_files % 100 == 0:
                pct = total_files * 100 // max(grand_total, 1)
                print(f"  [{pct:3d}%] {total_files:>5}/{grand_total} files | bad: {len(bad_files):>5} | now: {folder}", flush=True)
                last_print = now

    print()
    print("=" * 70)
    print(f"Tong files scanned: {total_files}")
    print(f"Files co translation loi: {len(bad_files)} ({len(bad_files)*100//max(total_files,1)}%)")
    if error_files:
        print(f"Files khong doc duoc (loi mo): {len(error_files)}")
    print()

    if by_family:
        print("Top 20 dong san pham co nhieu file loi:")
        for folder, count in sorted(by_family.items(), key=lambda x: -x[1])[:20]:
            print(f"  {count:5d}  {folder}")
        print()

    if not bad_files:
        print("KHONG co file nao bi loi translation. Hoan tat!")
        return

    if delete_mode:
        print(f"Xoa {len(bad_files)} file loi...")
        deleted = 0
        for f, _ in bad_files:
            try:
                os.remove(f)
                deleted += 1
            except Exception as e:
                print(f"  Khong xoa duoc {f}: {e}")
        print(f"Da xoa {deleted} file. Chay resume crawl de re-crawl voi translation moi:")
        print()
        print("  node crawl-checklist.js --checklist SICK_Web_Products_Statistics.xlsx \\")
        print("    --out-dir SICK_Products --resume --delay 500 > crawl.log 2>&1")
    else:
        print("(Chay --delete de xoa cac file loi -> resume crawl se re-crawl chung)")
        print("(Chay --sample 20 de xem 20 sample text loi)")


if __name__ == "__main__":
    main()
