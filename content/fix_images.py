# -*- coding: utf-8 -*-
"""Dung lai thu muc anh cho Pepperl+Fuchs va Reer.

    python fix_images.py            # ca 2 hang
    python fix_images.py reer       # mot hang

Hai loi da gap:
  1. data/Noi dung <hang>/images rong - anh chi duoc ghep luc crawl, build lai
     Excel khong tao lai anh nen xoa thu muc la mat.
  2. Ten anh lay theo slug(ten dong SP) ma 13 nhom P+F / 20 nhom Reer trung ten
     -> cac bai do dung chung MOT file anh. Nay them hau to -2, -3 cho khoi de.

Anh goc tai lai tu URL trong content.json (items[].img) roi ghep nen NTREND.
"""
import io
import json
import os
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import brandimg                                        # noqa: E402
import out_dir                                         # noqa: E402
from c_pepperl_try import UA, slugify                  # noqa: E402

BRANDS = {
    "pepperl": ("02 Pepperl+Fuchs", "Noi dung Pepperl+Fuchs"),
    "reer":    ("03 Reer", "Noi dung Reer"),
}


def first_img(a):
    return next((i.get("img") for i in (a.get("items") or []) if i.get("img")), "")


def run(key):
    sub, outsub = BRANDS[key]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = out_dir.img_dir(outsub)
    for d in (src_dir, img_dir):
        os.makedirs(d, exist_ok=True)
    doc = json.load(io.open(src_json, encoding="utf-8"))
    arts = doc["articles"]

    used, ok, miss = set(), 0, []
    for i, a in enumerate(arts, 1):
        url = first_img(a)
        if not url:
            a["img_file"] = ""
            miss.append(a["name"])
            continue
        # 1 bai P+F co "ten" la ca doan van -> slug dai qua gioi han duong dan
        base = (slugify(a.get("title_en") or a["name"])[:60].strip("-")
                or ("sp-%03d" % i))
        if base in used:                     # ten dong SP trung -> them hau to
            n = 2
            while "%s-%d" % (base, n) in used:
                n += 1
            base = "%s-%d" % (base, n)
        used.add(base)
        src = os.path.join(src_dir, base + (os.path.splitext(url.split("?")[0])[1][:5] or ".png"))
        try:
            if not (os.path.exists(src) and os.path.getsize(src) > 500):
                req = urllib.request.Request(url, headers={"User-Agent": UA})
                data = urllib.request.urlopen(req, timeout=60).read()
                if len(data) < 500:
                    raise ValueError("anh qua nho")
                with open(src, "wb") as f:
                    f.write(data)
            out = brandimg.compose(src, os.path.join(img_dir, base + ".jpg"))
            a["img_file"] = os.path.basename(out)
            ok += 1
        except Exception as e:
            a["img_file"] = ""
            miss.append("%s (%s)" % (a["name"], repr(e)[:40]))
        if i % 25 == 0:
            print("   %d/%d" % (i, len(arts)))
    with io.open(src_json, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    print("%-10s anh: %d/%d -> %s" % (key, ok, len(arts), img_dir))
    if miss:
        print("   khong co anh (%d): %s" % (len(miss), ", ".join(miss[:6])))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    keys = [a for a in sys.argv[1:] if a in BRANDS] or list(BRANDS)
    for k in keys:
        run(k)


if __name__ == "__main__":
    main()
