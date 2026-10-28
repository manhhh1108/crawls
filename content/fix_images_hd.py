# -*- coding: utf-8 -*-
"""Lam net lai anh dai dien cua Pepperl+Fuchs (va cac hang lay nham anh nho).

    python fix_images_hd.py pepperl
    python fix_images_hd.py reer

Van de: anh dai dien lay tu THE SAN PHAM tren trang danh sach, ma the do chi co
anh thu nho:

    P+F      200 x 220      (.../navi/productInfo/pd/b250540a.png)
    Reer     247 x 296      (...-247x296.png.webp)

Ghep len khung 800 x 800 cua NTREND thi phong to 3-4 lan -> vo het.

Anh net nam o cho khac:
    P+F   mo TRANG CHI TIET cua ma SP dau tien, trong do co ban '_hd' goi qua
          bo chuyen anh /_ipx/ voi w_1600  -> 1417 x 1559
    Reer  bo hau to kich thuoc trong ten file WooCommerce ('-247x296') -> 400 x 480

Chay xong nho build lai Excel cua hang do (ten file anh khong doi, chi doi noi
dung nen thuc ra khong bat buoc).
"""
import io
import json
import os
import re
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import brandimg                                        # noqa: E402
import out_dir                                         # noqa: E402

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
MIN_W = 300          # duoi nguong nay coi la anh thu nho, khong dung

BRANDS = {
    "pepperl": ("02 Pepperl+Fuchs", "Noi dung Pepperl+Fuchs"),
    "reer":    ("03 Reer", "Noi dung Reer"),
    "wenglor": ("06 wenglor", "Noi dung wenglor"),
    "leuze":   ("08 Leuze", "Noi dung Leuze"),
    "mitsu":   ("07 Mitsubishi Electric", "Noi dung Mitsubishi Electric"),
}

# P+F: /_ipx/f_webp/q_60/w_740/pffiles/...  ->  /_ipx/f_png/q_90/w_1600/pffiles/...
IPX = re.compile(r"/_ipx/[^/]+/[^/]+/w_\d+/")
PF_SRC = re.compile(r'(?:src|data-src)="(/_ipx/[^"]*?_hd\.[a-z]+)"')


def pf_hd_url(path):
    return "https://www.pepperl-fuchs.com" + IPX.sub("/_ipx/f_png/q_90/w_1600/", path)


def reer_hd_url(url):
    """...-247x296.png.webp -> ...png.webp (ban goc cua WooCommerce)."""
    return re.sub(r"-\d{2,4}x\d{2,4}(?=\.[a-z]+)", "", url)


# wenglor: the tren trang danh sach chi co ban ws194Wx109H, trang chi tiet moi
# co ban ws950Wx534H (cung ten B-PRO-... nhung khac ma ?context=)
WG_BIG = re.compile(r'(?:src|data-src|data-zoom-image|href)="'
                    r'(/medias/B-PRO-[^"]*?ws(\d{3,4})W[^"]*)"')


def get(url, timeout=60):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    return urllib.request.urlopen(req, timeout=timeout).read()


def width(data):
    from PIL import Image
    try:
        return Image.open(io.BytesIO(data)).size[0]
    except Exception:
        return 0


def pf_image(pg, item_url):
    """Mo trang chi tiet ma SP -> URL anh _hd."""
    pg.goto(item_url, wait_until="domcontentloaded", timeout=90000)
    pg.wait_for_timeout(800)
    m = PF_SRC.search(pg.content())
    return pf_hd_url(m.group(1)) if m else ""


def wg_image(html):
    """URL anh SP to nhat tren trang chi tiet cua wenglor."""
    best, bw = "", 0
    for u, w in WG_BIG.findall(html):
        if int(w) > bw:
            best, bw = u, int(w)
    return ("https://www.wenglor.com" + best) if bw > 400 else ""


def run_wenglor():
    sub, outsub = BRANDS["wenglor"]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = out_dir.img_dir(outsub)
    arts = json.load(io.open(src_json, encoding="utf-8"))["articles"]
    sys.path.insert(0, REPO)
    from shared.fetcher import try_fetch
    import html as H
    ok = skip = fail = 0
    for i, a in enumerate(arts, 1):
        name = a.get("img_file") or ""
        if not name:
            skip += 1
            continue
        base = os.path.splitext(name)[0]
        cur = os.path.join(src_dir, base + ".hd.webp")
        try:
            if not (os.path.exists(cur) and os.path.getsize(cur) > 5000):
                iu = next((x.get("url") for x in (a.get("items") or [])
                           if x.get("url")), "")
                if not iu:
                    skip += 1
                    continue
                big = wg_image(H.unescape(try_fetch(iu, retries=2, timeout=60) or ""))
                if not big:
                    fail += 1
                    continue
                data = get(big)
                if width(data) < MIN_W:
                    fail += 1
                    continue
                with open(cur, "wb") as f:
                    f.write(data)
            brandimg.compose(cur, os.path.join(img_dir, base + ".jpg"))
            ok += 1
        except Exception as e:
            fail += 1
            print("   [%s] %s" % (a["name"][:30], repr(e)[:60]))
        if i % 25 == 0:
            print("   %d/%d (net %d, bo qua %d, loi %d)"
                  % (i, len(arts), ok, skip, fail))
            sys.stdout.flush()
    print("wenglor: lam net %d anh | bo qua %d | loi %d" % (ok, skip, fail))


def run_leuze():
    """Leuze khong co ban anh nao to hon 326px (da kiem: /media/ goc va
    /thumbnail/..._800x800 deu ra 326). Chi can ghep lai de brandimg tu chon
    khung nho hon, khong phong to nua."""
    sub, outsub = BRANDS["leuze"]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = out_dir.img_dir(outsub)
    arts = json.load(io.open(src_json, encoding="utf-8"))["articles"]
    ok = skip = 0
    for a in arts:
        name = a.get("img_file") or ""
        src = os.path.join(src_dir, name)
        if not name or not os.path.exists(src):
            skip += 1
            continue
        brandimg.compose(src, os.path.join(img_dir, os.path.splitext(name)[0] + ".jpg"))
        ok += 1
    print("leuze: ghep lai %d anh | bo qua %d" % (ok, skip))


def run_mitsu():
    """Chon anh TO NHAT tren trang thay vi anh dau tien.

    Trang Mitsubishi tron lan icon, so do, anh thumbnail va anh san pham, lai
    khong co lop CSS nao phan biet. Lay anh dau tien (sau khi loai logo/icon)
    thi 36/162 bai trung phai anh thumbnail 30-96px -> ghep ra vo nat. Nay tai
    thu vai ung vien roi giu anh rong nhat.
    """
    sub, outsub = BRANDS["mitsu"]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = out_dir.img_dir(outsub)
    arts = json.load(io.open(src_json, encoding="utf-8"))["articles"]
    sys.path.insert(0, REPO)
    from shared.fetcher import try_fetch
    import c_mitsubishi as cm
    import html as H
    ok = skip = fail = 0
    for i, a in enumerate(arts, 1):
        name = a.get("img_file") or ""
        if not name:
            skip += 1
            continue
        base = os.path.splitext(name)[0]
        cur = os.path.join(src_dir, base + ".hd.png")
        try:
            if not (os.path.exists(cur) and os.path.getsize(cur) > 3000):
                html = try_fetch(a["url_src"], retries=2, timeout=60) or ""
                page = a["url_src"]
                best, bw = None, 0
                tried = 0
                for cand in cm.ALLIMG.findall(html):
                    c = H.unescape(cand)
                    low = c.lower()
                    if any(k in low for k in cm.IMG_SKIP) or low.endswith(".svg"):
                        continue
                    if c.startswith("http"):
                        u = c
                    elif c.startswith("/"):
                        u = cm.SITE + c
                    else:
                        u = page.rsplit("/", 1)[0] + "/" + c.lstrip("./")
                        while "/../" in u:
                            u = re.sub(r"/[^/]+/\.\./", "/", u)
                    tried += 1
                    if tried > 12:
                        break
                    try:
                        data = get(u, timeout=30)
                    except Exception:
                        continue
                    w = width(data)
                    if w > bw:
                        best, bw = data, w
                    if bw >= 900:
                        break
                if not best or bw < 200:
                    fail += 1
                    continue
                with open(cur, "wb") as f:
                    f.write(best)
            brandimg.compose(cur, os.path.join(img_dir, base + ".jpg"))
            ok += 1
        except Exception as e:
            fail += 1
            print("   [%s] %s" % (a["name"][:28], repr(e)[:50]))
        if i % 25 == 0:
            print("   %d/%d (net %d, bo qua %d, loi %d)"
                  % (i, len(arts), ok, skip, fail))
            sys.stdout.flush()
    print("mitsu: lam net %d anh | bo qua %d | loi %d" % (ok, skip, fail))


def run_pepperl():
    sub, outsub = BRANDS["pepperl"]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = out_dir.img_dir(outsub)
    doc = json.load(io.open(src_json, encoding="utf-8"))
    arts = doc["articles"]

    from playwright.sync_api import sync_playwright
    ok = skip = fail = 0
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, channel="chrome")
        pg = b.new_page()
        try:
            for i, a in enumerate(arts, 1):
                name = a.get("img_file") or ""
                if not name:
                    skip += 1
                    continue
                base = os.path.splitext(name)[0]
                cur = os.path.join(src_dir, base + ".hd.png")
                try:
                    if not (os.path.exists(cur) and os.path.getsize(cur) > 5000):
                        iu = next((x.get("url") for x in (a.get("items") or [])
                                   if x.get("url")), "")
                        if not iu:
                            skip += 1
                            continue
                        hd = pf_image(pg, iu)
                        if not hd:
                            fail += 1
                            continue
                        data = get(hd)
                        if width(data) < MIN_W:
                            fail += 1
                            continue
                        with open(cur, "wb") as f:
                            f.write(data)
                    brandimg.compose(cur, os.path.join(img_dir, base + ".jpg"))
                    ok += 1
                except Exception as e:
                    fail += 1
                    print("   [%s] %s" % (a["name"][:30], repr(e)[:60]))
                if i % 20 == 0:
                    print("   %d/%d (net %d, bo qua %d, loi %d)"
                          % (i, len(arts), ok, skip, fail))
                    sys.stdout.flush()
        finally:
            b.close()
    print("pepperl: lam net %d anh | bo qua %d | loi %d" % (ok, skip, fail))


def run_reer():
    sub, outsub = BRANDS["reer"]
    src_json = os.path.join(HERE, "out", sub, "content.json")
    src_dir = os.path.join(HERE, "out", sub, "src_img")
    img_dir = out_dir.img_dir(outsub)
    arts = json.load(io.open(src_json, encoding="utf-8"))["articles"]
    ok = skip = fail = 0
    for i, a in enumerate(arts, 1):
        name = a.get("img_file") or ""
        url = next((x.get("img") for x in (a.get("items") or []) if x.get("img")), "")
        if not name or not url:
            skip += 1
            continue
        base = os.path.splitext(name)[0]
        cur = os.path.join(src_dir, base + ".hd.png")
        try:
            if not (os.path.exists(cur) and os.path.getsize(cur) > 5000):
                data = get(reer_hd_url(url))
                if width(data) < 350:        # Reer goc chi 400px, ha nguong
                    fail += 1
                    continue
                with open(cur, "wb") as f:
                    f.write(data)
            brandimg.compose(cur, os.path.join(img_dir, base + ".jpg"))
            ok += 1
        except Exception as e:
            fail += 1
            print("   [%s] %s" % (a["name"][:30], repr(e)[:60]))
        if i % 40 == 0:
            print("   %d/%d" % (i, len(arts)))
            sys.stdout.flush()
    print("reer: lam net %d anh | bo qua %d | loi %d" % (ok, skip, fail))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    runners = {"pepperl": run_pepperl, "reer": run_reer,
               "wenglor": run_wenglor, "leuze": run_leuze,
               "mitsu": run_mitsu}
    for k in [a for a in sys.argv[1:] if a in BRANDS] or ["pepperl"]:
        runners[k]()


if __name__ == "__main__":
    main()
