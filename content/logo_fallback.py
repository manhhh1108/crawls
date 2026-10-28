# -*- coding: utf-8 -*-
"""Bai nao khong co anh san pham thi lay LOGO HANG lam anh dai dien.

    python logo_fallback.py              # tat ca cac hang
    python logo_fallback.py mitsu ifm    # vai hang
    python logo_fallback.py --undo mitsu # go lai, tra ve "(chua co anh)"

Lam giong c_sick.py da lam cho SICK tu dau: dong SP nao trang hang khong co anh
thi ghep logo SICK vao khung NTREND. Nay ap dung cho cac hang con lai.

Hai nguon lam bai khong co anh:
  * trang hang khong dang anh nao cho dong SP do
  * anh goc qua nho nen recompose_images.py da bo (xem MIN_SRC o do) - vd
    Mitsubishi chi co bieu tuong nhom module 34x50 px, phong len la vo nat

Chay sau cung, khi content.json da chot anh that. Sau do chay lai
build_<hang>_excel.py de o "Anh san pham" trong bai hien ten file moi.
"""
import io
import json
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

import brandimg                                        # noqa: E402

ASSETS = os.path.join(HERE, "assets")

# key: (thu muc out, thu muc data, ten file logo trong assets/)
BRANDS = {
    "sick":     ("01 SICK", "Noi dung SICK", "sick-logo.jpg"),
    "pepperl":  ("02 Pepperl+Fuchs", "Noi dung Pepperl+Fuchs", "pepperl-logo.png"),
    "reer":     ("03 Reer", "Noi dung Reer", "reer-logo.png"),
    "ifm":      ("04 IFM", "Noi dung IFM", "ifm-logo.png"),
    "perma":    ("05 Perma", "Noi dung Perma", "perma-logo.png"),
    "wenglor":  ("06 wenglor", "Noi dung wenglor", "wenglor-logo.png"),
    "mitsu":    ("07 Mitsubishi Electric", "Noi dung Mitsubishi Electric",
                 "mitsu-logo.png"),
    "leuze":    ("08 Leuze", "Noi dung Leuze", "leuze-logo.png"),
    "festo":    ("10 Festo", "Noi dung Festo", "festo-logo.png"),
    "espe":     ("15 ESPE Technology", "Noi dung ESPE Technology",
                 "espe-logo.png"),
    "otennlux": ("16 OTENNLUX Lighting", "Noi dung OTENNLUX Lighting",
                 "otennlux-logo.png"),
    "balluff":  ("17 Balluff", "Noi dung Balluff", "balluff-logo.png"),
}


def bases(arts, slugify):
    """Ten file anh cua tung bai, dung dung quy tac cac crawler dang dung.

    Ten sinh ra tu ten dong SP (slug, cat 60 ky tu, trung thi them -2/-3) nen
    phai duyet CA danh sach theo thu tu moi ra dung hau to, khong the tinh
    rieng mot bai.
    """
    out, used = [], set()
    for a in arts:
        b = (slugify(a.get("title_en") or a.get("name") or "")[:60].strip("-")
             or "sp")
        if b in used:
            n = 2
            while "%s-%d" % (b, n) in used:
                n += 1
            b = "%s-%d" % (b, n)
        used.add(b)
        out.append(b)
    return out


def run(key, undo=False):
    sub, outsub, logo = BRANDS[key]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    img_dir = os.path.join(REPO, "data", outsub, "images")
    if not os.path.exists(src_json):
        print("%-9s (chua co content.json)" % key)
        return
    doc = json.load(io.open(src_json, encoding="utf-8"))
    arts = doc["articles"]

    from c_pepperl_try import slugify
    bs = bases(arts, slugify)

    logo_path = os.path.join(ASSETS, logo)
    if not undo:
        thieu = sum(1 for a in arts if not a.get("img_file"))
        if not thieu:
            print("%-9s du anh ca %d bai, khong can logo" % (key, len(arts)))
            return
        if not os.path.exists(logo_path):
            print("%-9s %d bai thieu anh nhung CHUA CO LOGO: %s"
                  % (key, thieu, logo_path))
            return

    if undo:
        n = 0
        for a, b in zip(arts, bs):
            if not a.pop("img_logo", False):
                continue
            a["img_file"] = ""
            p = os.path.join(img_dir, b + ".jpg")
            if os.path.exists(p):
                os.remove(p)
            n += 1
        with io.open(src_json, "w", encoding="utf-8") as f:
            json.dump(doc, f, ensure_ascii=False)
        print("%-9s da go %d anh logo" % (key, n))
        return

    os.makedirs(img_dir, exist_ok=True)
    done, keep, fail = 0, 0, 0
    names = []
    for a, b in zip(arts, bs):
        if a.get("img_file"):
            keep += 1
            continue
        try:
            out = brandimg.compose(logo_path, os.path.join(img_dir, b + ".jpg"))
        except Exception as e:
            fail += 1
            print("   [%s] %s" % ((a.get("title_en") or a["name"])[:30],
                                  repr(e)[:50]))
            continue
        a["img_file"] = os.path.basename(out)
        a["img_logo"] = True
        names.append(a.get("title_en") or a["name"])
        done += 1
    with io.open(src_json, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    print("%-9s dat logo cho %d bai | anh that %d | loi %d"
          % (key, done, keep, fail))
    for n in names[:8]:
        print("      - %s" % n[:60])
    if len(names) > 8:
        print("      ... va %d bai nua" % (len(names) - 8))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    undo = "--undo" in sys.argv
    keys = [a for a in sys.argv[1:] if a in BRANDS] or list(BRANDS)
    for k in keys:
        run(k, undo)


if __name__ == "__main__":
    main()
