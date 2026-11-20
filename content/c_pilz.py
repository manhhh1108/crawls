# -*- coding: utf-8 -*-
"""Pilz (hang 09) - crawl noi dung bai dang cho tung muc eshop.

    python c_pilz.py              # chay tiep nhung dong chua co
    python c_pilz.py 5            # chi 5 dong dau (xem thu)
    python c_pilz.py --force      # lam lai tu dau

Can chay truoc: python cat_map.py pilz

pilz.com.cn/en-CN khong chan curl -> dung shared.fetcher. Moi dong trong danh
muc Ntrend la MOT MUC ESHOP (trang /c/<ma>), tuc mot dong SP gom nhieu ma:

    <h1>Plug-in connectors</h1>
    <div class="pagination-bar-results">20 Products found</div>
    <article class="product__list--item">
        <a class="product__list--thumb" href=...><img src="/medias/...">
        <div class="product__code"><strong>380308</strong></div>
        <a class="product__list--name" ...>PSS67 M12 connector,straight,male,5pole</a>
        <div class="product-keywords">Plug-in connector, 5-pin, ...</div>

Dung ba cot do cho bang ma SP cua template: Model | Ma dat hang | Mo ta.

Trang chi hien 20 ma mot lan, phan trang bang ?q=%3Aproductrelevance&page=N
(tham so pageSize bi bo qua). 31/180 muc co hon 20 ma, muc lon nhat 244 ma.

Mo ta dai theo template ("Lấy full thông số/đặc điểm của sản phẩm") lay tu
KHOI BO LOC ben trai: moi bo loc la mot thong so cua ca dong SP va cac gia tri
co that trong muc do - giong cach da lam cho Festo (facets).

Trang muc khong co doan van gioi thieu nao, nen mo ta ngan dung mau DESCDESC.

ANH: anh thu nho cua ma SP dau tien trong muc. Muc "0 Products found" khong co
the san pham nao -> lay anh banner (`class="category-image"`) o dau trang. Anh
hang dang chi 188-300 px, markup khong co co nao to hon (quet 784 trang trong
cache chi thay mot co `300Wx300H`), nhung khung giao van ep du 800x800 theo
yeu cau cua khach.

Ket qua: out/09 Pilz/content.json -> build_pilz_excel.py dung tiep.
"""
import io
import json
import os
import re
import subprocess
import sys
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

import out_dir                                         # noqa: E402

BRAND = "Pilz"
SITE = "https://www.pilz.com.cn"
OUT = os.path.join(HERE, "out", "09 Pilz")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung Pilz")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

PAGE_SIZE = 20
MAX_PAGES = 40           # tran an toan: 40 trang = 800 ma cho mot muc
MAX_ITEMS = 600
MAX_SPECS = 24           # so bo loc dua vao mo ta dai
MAX_VALS = 10            # so gia tri moi bo loc

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

H1 = re.compile(r"(?s)<h1[^>]*>(.*?)</h1>")
FOUND = re.compile(r'pagination-bar-results">\s*([\d,]+)\s*Products? found')
ITEM = re.compile(r'(?s)<article class="product__list--item">(.*?)</article>')
I_CODE = re.compile(r'(?s)product__code[^"]*"><strong>(.*?)</strong>')
I_NAME = re.compile(r'(?s)<a class="product__list--name"[^>]*>(.*?)</a>')
I_DESC = re.compile(r'(?s)<div class="product-keywords">(.*?)</div>')
I_IMG = re.compile(r'(?s)class="product__list--thumb"[^>]*>.*?<img src="([^"]+)"')
I_HREF = re.compile(r'class="product__list--thumb" href="([^"]+)"')
# Anh banner o dau trang muc. Muc "0 Products found" khong co the san pham nao
# nen khong co anh thu nho, nhung van co anh nay (review 04/11).
CAT_IMG = re.compile(r'<img[^>]*class="category-image"[^>]*>')
SRC_ATTR = re.compile(r'src="([^"]+)"')

FACET = re.compile(r'(?s)<div class="facet js-facet">(.*?)(?=<div class="facet '
                   r'js-facet">|<div id="product-facet-mobile|</form>\s*</div>\s*'
                   r'</div>\s*</div>\s*</div>)')
F_NAME = re.compile(r'(?s)facet__name js-facet-name">(.*?)</div>')
F_VAL1 = re.compile(r'(?s)facet__list__text">(.*?)<span class="facet__value__count"')
F_VAL2 = re.compile(r'(?s)<span class="facet__text">\s*<a[^>]*>(.*?)</a>')


def txt(s):
    return re.sub(r"\s+", " ",
                  H.unescape(re.sub(r"<[^>]+>", " ", s or ""))).strip()


def page_url(url, n):
    """Trang thu n cua muc, ban co tham so sap xep.

    Trang mo tran (khong co ?q=) sap xep KHAC voi cac trang phan trang: o muc
    "Contact expansions" hai ban chi trung 10/20 ma. Nen phai lay CA HAI roi
    gop lai - chi lay mot ben thi ben kia sot ma.
    """
    return "%s?q=%%3Aproductrelevance&page=%d" % (url, n)


def parse_items(html):
    out = []
    for blk in ITEM.findall(html):
        c, n = I_CODE.search(blk), I_NAME.search(blk)
        if not c or not n:
            continue
        d, im = I_DESC.search(blk), I_IMG.search(blk)
        hr = I_HREF.search(blk)
        out.append({"art": txt(c.group(1)), "name": txt(n.group(1)),
                    "desc": txt(d.group(1)) if d else "",
                    "img": im.group(1) if im else "",
                    "url": (SITE + hr.group(1)) if hr else ""})
    return out


def parse_facets(html):
    """[[ten bo loc, "gia tri; gia tri; ..."]] - thong so cua ca dong SP.

    Bo loc "Approval" co gia tri "-" (nghia la khong cong bo) va bo loc nao
    cung keo theo so luong trong ngoac - cat het, chi giu gia tri.
    """
    out = []
    for blk in FACET.findall(html):
        nm = F_NAME.search(blk)
        if not nm:
            continue
        name = txt(nm.group(1))
        vals = []
        for v in F_VAL1.findall(blk) + F_VAL2.findall(blk):
            t = txt(v).rstrip(u"\xa0 ")
            t = re.sub(r"\s*\(\d+\)\s*$", "", t).strip()
            if t and t != "-" and t not in vals:
                vals.append(t)
        if name and vals:
            out.append([name, "; ".join(vals[:MAX_VALS])])
    return out[:MAX_SPECS]


def fetch_line(url):
    from shared.fetcher import try_fetch
    s = try_fetch(url, retries=2)
    if not s:
        return None
    m = FOUND.search(s)
    total = int(m.group(1).replace(",", "")) if m else 0
    # Gop ma cua ca hai ban: trang tran (dang vua tai) + tung trang co ?q=.
    items = parse_items(s)
    npg = min(MAX_PAGES, -(-total // PAGE_SIZE) if total else 1)
    hong = []
    for n in range(0, npg):
        if len(items) >= MAX_ITEMS:
            break
        # Thinh thoang mot trang tra ve HTTP 000 (curl dut ket noi) - khong
        # duoc dung lai vi cac trang sau van lay duoc, chi ghi lai la thieu.
        t = try_fetch(page_url(url, n), retries=4)
        if not t:
            hong.append(n)
            continue
        more = parse_items(t)
        if not more:
            if n == 0:       # ban ?q= rong nhung ban tran co the van co ma
                continue
            break
        items += more
    # Trang hang dang LAP MA that: muc "Accessories PSSuniversal" dem 27 san
    # pham nhung trong do co 2 dong trung y het ca ma lan ten (313111, 313116),
    # nen bo trung xong chi con 25. Khong phai cao sot.
    uniq, seen = [], set()
    for it in items:
        if it["art"] in seen:
            continue
        seen.add(it["art"])
        uniq.append(it)
    h1 = H1.search(s)
    # Anh bai: uu tien anh thu nho cua ma SP dau tien; muc khong co ma nao thi
    # lay anh banner dau trang. 10/18 muc truoc day phai dung logo hang nay co
    # anh that.
    img = next((i["img"] for i in uniq if i["img"]), "")
    if not img:
        m = CAT_IMG.search(s)
        g = SRC_ATTR.search(m.group(0)) if m else None
        img = g.group(1) if g else ""
    return {"h1": txt(h1.group(1)) if h1 else "",
            "total": total or len(uniq),
            "items": uniq,
            "trang_hong": hong,
            "specs": parse_facets(s),
            "img": img}


def is_code_name(s):
    s = (s or "").strip()
    if not s:
        return False
    if " " not in s:
        return bool(re.search(r"[A-Z]{2,}", s)) or bool(re.search(r"\d", s))
    return not re.search(r"[a-z]", s)


def name_vi_of(name, translate_many):
    if is_code_name(name):
        return name
    from vi_engine import lookup
    return lookup(name) or translate_many([name], "key")[0]


def co_anh(p):
    """File tai ve co phai anh that khong (khong phai trang bao loi)."""
    try:
        if os.path.getsize(p) < 1200:
            return False
        from PIL import Image
        with Image.open(p) as im:
            im.verify()
        return True
    except Exception:
        return False


def grab_image(url, line_en, slugify, used):
    """Anh tren pilz.com.cn phai tai bang curl: urllib bi tra ve HTTP 567."""
    if not url:
        return ""
    base = slugify(line_en)[:60].strip("-") or "sp"
    if base in used:
        n = 2
        while "%s-%d" % (base, n) in used:
            n += 1
        base = "%s-%d" % (base, n)
    used.add(base)
    full = url if url.startswith("http") else SITE + url
    src = os.path.join(SRC_IMG, base + ".jpg")
    try:
        if not co_anh(src):
            subprocess.run(
                ["curl", "-sS", "-L", "--max-time", "60", "-o", src,
                 "-A", UA, "-H", "Referer: " + SITE + "/en-CN", full],
                capture_output=True, timeout=90)
        if not co_anh(src):
            # Anh banner nam duoi /mam/ bi may chu chan bang trang do cookie
            # Javascript (tra ve 657 byte HTML). Xoa di de lan sau con thu
            # lai, dung de file rac do lam "anh goc".
            if os.path.exists(src):
                os.remove(src)
            return ""
        import brandimg
        return os.path.basename(
            brandimg.compose(src, os.path.join(IMG_DIR, base + ".jpg")))
    except Exception as e:
        print("     [anh loi]", repr(e)[:60])
        return ""


# Ten muc trong danh muc Ntrend co kem so ma: "Plug-in connectors (20)"
CNT = re.compile(r"\s*\(\d+\)\s*$")


def save(ls, done):
    arts, seen = [], set()
    for l in ls:
        u = l.get("url")
        if u in done and u not in seen:
            seen.add(u)
            arts.append(done[u])
    with io.open(CONTENT, "w", encoding="utf-8") as f:
        json.dump({"brand": "09 Pilz", "articles": arts}, f, ensure_ascii=False)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    force = "--force" in sys.argv
    limit = next((int(a) for a in sys.argv[1:] if a.isdigit()), 0)

    from c_pepperl_try import slugify
    from vi_engine import translate_many
    from translate import _flush

    ls = [l for l in json.load(io.open(URLMAP, encoding="utf-8"))["lines"]
          if l.get("url")]
    if limit:
        ls = ls[:limit]
    uniq, seen = [], set()
    for l in ls:
        if l["url"] in seen:
            continue
        seen.add(l["url"])
        uniq.append(l)
    done = {}
    if os.path.exists(CONTENT) and not force:
        done = {a["url_src"]: a for a in
                json.load(io.open(CONTENT, encoding="utf-8"))["articles"]
                if a.get("url_src")}
    used = {os.path.splitext(a["img_file"])[0] for a in done.values()
            if a.get("img_file")}
    todo = [l for l in uniq if l["url"] not in done]
    print("muc eshop: %d | bo trung cung trang: %d | bai se co: %d"
          % (len(ls), len(ls) - len(uniq), len(uniq)))
    print("da xong: %d | se lam: %d" % (len(done), len(todo)))
    if not todo:
        return

    import prewarm
    raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
    for i, ln in enumerate(todo, 1):
        try:
            d = fetch_line(ln["url"])
        except Exception as e:
            d = None
            print("     LOI", repr(e)[:70])
        if not d:
            print("  cao %3d/%d %-38s KHONG TAI DUOC"
                  % (i, len(todo), ln["name"][:38]))
            continue
        title = d["h1"] or CNT.sub("", ln["name"])
        d["img_file"] = grab_image(d["img"], title, slugify, used)
        d["title"] = title
        raw[ln["url"]] = d

        bag["key"].add(title)
        bag["key"].update(k for k, _v in d["specs"])
        bag["value"].update(v for _k, v in d["specs"] if v)
        bag["value"].update(it["desc"] for it in d["items"] if it["desc"])
        print("  cao %3d/%d %-38s ma=%-4d/%-4d ts=%-2d %s%s"
              % (i, len(todo), title[:38], len(d["items"]), d["total"],
                 len(d["specs"]), "img" if d["img_file"] else "---",
                 "  THIEU %d TRANG" % len(d["trang_hong"])
                 if d["trang_hong"] else ""))
        sys.stdout.flush()
    prewarm.run(bag)

    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, specs, items = d["title"], d["specs"], d["items"]
        keys = sorted({k for k, _v in specs})
        vals = sorted({v for _k, v in specs if v}
                      | {it["desc"] for it in items if it["desc"]})
        tr_k = dict(zip(keys, translate_many(keys, "key"))) if keys else {}
        tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        specs_vi = [[tr_k.get(k, k), tr_v.get(v, v)] for k, v in specs]
        for it in items:
            it["desc_vi"] = tr_v.get(it["desc"], it["desc"])

        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title,
            series="", codes=[it["art"] for it in items],
            img_file=d["img_file"], total=d["total"], items=items,
            specs_en=specs, specs_vi=specs_vi,
            long_en=[], long_vi=[], short_en=[], short_vi=[],
            name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-38s ma=%-4d ts=%-2d %s"
              % (i, len(todo), title[:38], len(items), len(specs),
                 "img" if d["img_file"] else "---"))
        if i % 10 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
