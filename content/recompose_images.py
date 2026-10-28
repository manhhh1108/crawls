# -*- coding: utf-8 -*-
"""Ghep lai TAT CA anh dai dien tu ban goc net nhat dang co.

    python recompose_images.py              # ca 6 hang
    python recompose_images.py reer sick

Chay sau fix_images_hd.py: nhung bai khong lay duoc ban HD (trang hang khong
co, hoac tai loi) van con giu ban ghep cu - ban do duoc phong len khung 800 tu
anh 200-300px nen vo. Ghep lai de brandimg tu chon khung nho hon (xem
brandimg.MAX_UPSCALE), anh se net du nho hon mot chut.

Uu tien nguon: <ten>.hd.*  (ban HD do fix_images_hd.py tai)  ->  <ten>.*

Nguon nho hon MIN_SRC thi BO HAN anh (xoa file, xoa img_file trong content.json):
anh 30-100px phong len khung du nho nhat van vo nat, de trong con hon.
"""
import glob
import io
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import brandimg                                        # noqa: E402

MIN_SRC = 200        # CANH DAI NHAT toi thieu cua anh goc thi moi dang ghep
# (do canh dai nhat chu khong phai chieu rong: luoi quang an toan cua Leuze
#  la anh cao hep 31 x 326, do theo chieu rong se loai oan)

BRANDS = {
    "sick":    ("01 SICK", "Noi dung SICK"),
    "pepperl": ("02 Pepperl+Fuchs", "Noi dung Pepperl+Fuchs"),
    "reer":    ("03 Reer", "Noi dung Reer"),
    "ifm":     ("04 IFM", "Noi dung IFM"),
    "perma":   ("05 Perma", "Noi dung Perma"),
    "wenglor": ("06 wenglor", "Noi dung wenglor"),
    "mitsu":   ("07 Mitsubishi Electric", "Noi dung Mitsubishi Electric"),
    "leuze":   ("08 Leuze", "Noi dung Leuze"),
}


def best_source(src_dir, base):
    """Ban HD neu co, khong thi ban thuong."""
    hd = sorted(glob.glob(os.path.join(src_dir, base + ".hd.*")))
    if hd:
        return hd[0]
    plain = [p for p in sorted(glob.glob(os.path.join(src_dir, base + ".*")))
             if ".hd." not in os.path.basename(p)]
    return plain[0] if plain else ""


def restore(arts, src_dir):
    """Dat lai img_file cho bai da bi xoa nham o lan chay truoc.

    Ten file anh duoc sinh deterministic tu ten dong SP (slug, cat 60 ky tu,
    trung thi them -2/-3) nen suy lai duoc, mien la file goc van con trong
    src_img.
    """
    sys.path.insert(0, REPO)
    try:
        from c_pepperl_try import slugify
    except Exception:
        return
    used = set()
    for a in arts:
        base = (slugify(a.get("title_en") or a.get("name") or "")[:60].strip("-")
                or "sp")
        if base in used:
            n = 2
            while "%s-%d" % (base, n) in used:
                n += 1
            base = "%s-%d" % (base, n)
        used.add(base)
        if not a.get("img_file") and best_source(src_dir, base):
            a["img_file"] = base + ".jpg"


def run(key):
    sub, outsub = BRANDS[key]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = os.path.join(REPO, "data", outsub, "images")
    if not (os.path.exists(src_json) and os.path.isdir(src_dir)):
        print("%-9s (chua co du lieu)" % key)
        return
    doc = json.load(io.open(src_json, encoding="utf-8"))
    arts = doc["articles"]
    from PIL import Image
    restore(arts, src_dir)
    ok = miss = tiny = 0
    sizes = {}
    for a in arts:
        name = a.get("img_file") or ""
        if not name:
            continue
        base = os.path.splitext(name)[0]
        src = best_source(src_dir, base)
        if not src:
            miss += 1
            continue
        try:
            if max(Image.open(src).size) < MIN_SRC:
                # anh goc qua nho -> bo, de bai hien "(chua co anh)"
                a["img_file"] = ""
                p = os.path.join(img_dir, base + ".jpg")
                if os.path.exists(p):
                    os.remove(p)
                tiny += 1
                continue
            out = brandimg.compose(src, os.path.join(img_dir, base + ".jpg"))
            w = Image.open(out).size[0]
            sizes[w] = sizes.get(w, 0) + 1
            ok += 1
        except Exception as e:
            miss += 1
            print("   [%s] %s" % (a["name"][:28], repr(e)[:50]))
    with io.open(src_json, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    sz = ", ".join("%dpx x%d" % kv for kv in sorted(sizes.items(), reverse=True)[:4])
    print("%-9s ghep lai %d | bo vi qua nho %d | thieu nguon %d | khung: %s"
          % (key, ok, tiny, miss, sz))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    for k in [a for a in sys.argv[1:] if a in BRANDS] or list(BRANDS):
        run(k)


if __name__ == "__main__":
    main()
