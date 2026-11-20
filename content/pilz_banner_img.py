# -*- coding: utf-8 -*-
"""Pilz - lay ANH BANNER dau trang cho nhung muc khong co anh san pham.

    python pilz_banner_img.py            # chay that
    python pilz_banner_img.py --xem      # chi liet ke, khong tai

Chay SAU c_pilz.py, roi chay lai build_pilz_excel.py.

VI SAO PHAI TACH RA MOT SCRIPT RIENG
Muc eshop nao co "0 Products found" thi khong co the san pham nao, nen khong
co anh thu nho - truoc day 18 muc nhu vay deu phai dung logo hang. Nhung dau
trang van co anh banner (`<img class="category-image">`), va khach yeu cau lay
dung anh do (review 04/11).

`c_pilz.grab_image()` doc ra URL anh banner roi, nhung TAI KHONG DUOC: anh
banner nam duoi `/mam/...`, va may chu tra ve mot trang 657 byte dat cookie
`__wc_jsct` bang Javascript roi reload - lop chan bot. Anh san pham duoi
`/medias/` thi curl tai binh thuong.

Nen rieng may anh nay phai di qua Chromium that (shared.browser.Browser): mo
trang eshop mot lan cho Javascript chay xong, roi tai anh trong CUNG phien do.
Khach da dong y cach nay (04/11). Khong gop vao c_pilz.py vi moi lan cao lai
se phai bat Chromium len cho 10 anh, trong khi 162 anh con lai curl tai thua.

Anh hang dang chi 188-218 px - khung giao van ep du 800x800 theo yeu cau, nen
nhin se hoi mem. Do la gioi han cua nguon.
"""
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

import brandimg                                        # noqa: E402
import c_pilz                                          # noqa: E402

CONTENT = c_pilz.CONTENT
SRC_IMG = c_pilz.SRC_IMG
IMG_DIR = c_pilz.IMG_DIR
HOME = c_pilz.SITE + "/en-CN/eshop"


def banner_url(url):
    """URL anh banner cua mot muc, doc tu trang da cache."""
    from shared.fetcher import try_fetch
    s = try_fetch(url, retries=2) or ""
    m = c_pilz.CAT_IMG.search(s)
    g = c_pilz.SRC_ATTR.search(m.group(0)) if m else None
    if not g:
        return ""
    u = g.group(1)
    return u if u.startswith("http") else c_pilz.SITE + u


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    xem = "--xem" in sys.argv

    doc = json.load(io.open(CONTENT, encoding="utf-8"))
    arts = doc["articles"]
    # Bai chua co anh that = bai dang deo logo, hoac chua co anh nao.
    thieu = [a for a in arts if a.get("img_logo") or not a.get("img_file")]
    print("bai chua co anh that: %d/%d" % (len(thieu), len(arts)))

    viec = []
    for a in thieu:
        u = banner_url(a["url_src"])
        ten = (a.get("title_en") or a["name"])[:44]
        if not u:
            print("  %-44s trang hang khong co anh banner" % ten)
            continue
        viec.append((a, u))
        print("  %-44s %s" % (ten, u[-62:]))
    print("co anh banner: %d/%d" % (len(viec), len(thieu)))
    if xem or not viec:
        return

    from shared.browser import Browser
    br = Browser(HOME, headless=True)
    xong, hong = 0, 0
    try:
        for a, u in viec:
            ten = (a.get("title_en") or a["name"])[:44]
            # Giu dung ten file bai dang dung (logo_fallback da dat) de anh
            # that GHI DE len logo, khong de lai file thua trong images/.
            base = (os.path.splitext(a.get("img_file") or "")[0]
                    or re.sub(r"[^a-z0-9-]+", "-", ten.lower()).strip("-"))
            src = os.path.join(SRC_IMG, base + os.path.splitext(
                u.split("?")[0])[1][:5].lower() or ".jpg")
            try:
                if not c_pilz.co_anh(src):
                    r = br.ctx.request.get(u, timeout=60000)
                    if r.status != 200:
                        raise RuntimeError("HTTP %s" % r.status)
                    with open(src, "wb") as f:
                        f.write(r.body())
                if not c_pilz.co_anh(src):
                    raise RuntimeError("tai ve khong phai anh")
                out = brandimg.compose(src, os.path.join(IMG_DIR,
                                                         base + ".jpg"))
            except Exception as e:
                hong += 1
                print("  LOI %-44s %s" % (ten, repr(e)[:50]))
                continue
            a["img_file"] = os.path.basename(out)
            a.pop("img_logo", None)
            xong += 1
            print("  OK  %-44s %s" % (ten, a["img_file"]))
    finally:
        br.close()

    with io.open(CONTENT, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    print("\nthay logo bang anh banner: %d bai | loi %d" % (xong, hong))
    print("con lai dung logo: %d bai"
          % sum(1 for a in arts if a.get("img_logo")))
    print("chay tiep: python build_pilz_excel.py")


if __name__ == "__main__":
    main()
