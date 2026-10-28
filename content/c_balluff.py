# -*- coding: utf-8 -*-
"""Balluff (hang 17) - crawl noi dung bai dang cho tung dong san pham.

    python c_balluff.py            # chay tiep nhung dong chua co
    python c_balluff.py 3          # chi 3 dong dau (xem thu)
    python c_balluff.py --force    # lam lai tu dau

Can chay truoc: python cat_map.py balluff

balluff.com dung Cloudflare, tra 403 cho moi request khong phai trinh duyet.
NHUNG bo loc san pham cua trang goi mot REST API nam o TEN MIEN KHAC
(balluff-search-2-api.linksderisar.io) va ten mien do KHONG bi chan -> khong
can Playwright, chi can 2 lan goi cho moi dong SP:

    /categories/<ma>                     ten, description, bulletpoints,
                                         image_url (1280x1280), columns
    /categories/<ma>/products?page=N     order_code, label, list_columns

`columns` moi dong SP mot bo khac nhau (Use / Dimension / Material / Version /
Focal length ...) va cho san nhan tieng Anh de lam dau cot bang ma SP.

Theo 'Template nội dung Balluff.docx':
    Mo ta ngan  = bulletpoints, toi da 5 gach dau dong
    Mo ta dai   = doan gioi thieu phan Feature + thong so gom tu bang ma SP
    Bang ma SP  = Ma dat hang | Model | cac cot trong `columns`

Ket qua: out/17 Balluff/content.json -> build_balluff_excel.py dung tiep.
"""
import gzip
import io
import json
import os
import re
import sys
import time
import urllib.request
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

import out_dir                                         # noqa: E402

BRAND = "Balluff"
SITE = "https://www.balluff.com"
API = "https://balluff-search-2-api.linksderisar.io/api/product-finder"
OUT = os.path.join(HERE, "out", "17 Balluff")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung Balluff")
IMG_DIR = os.path.join(XLSX_DIR, "images")
CACHE = os.path.join(REPO, ".cache", "balluff")
for d in (OUT, SRC_IMG, IMG_DIR, CACHE):
    os.makedirs(d, exist_ok=True)

MAX_COLS = 8
MAX_BULLETS = 5        # template: mo ta ngan toi da 5 gach dau dong
MAX_HL = 8

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
HDRS = {"User-Agent": UA, "Accept": "application/json",
        "Accept-Encoding": "gzip", "Origin": SITE, "Referer": SITE + "/"}


def txt(s):
    s = re.sub(r"(?i)<br\s*/?>|</p>|</li>", " ", s or "")
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s))).strip()


def api(path, tries=3):
    """Goi API, co cache tren dia (API nay khong qua shared.fetcher vi
    shared.fetcher goi curl voi bo header cua trang HTML)."""
    import hashlib
    p = os.path.join(CACHE, hashlib.sha1(path.encode()).hexdigest() + ".gz")
    if os.path.exists(p):
        with gzip.open(p, "rt", encoding="utf-8") as f:
            return json.load(f)
    last = None
    for i in range(tries):
        try:
            r = urllib.request.urlopen(
                urllib.request.Request(API + path, headers=HDRS), timeout=60)
            b = r.read()
            if r.headers.get("Content-Encoding") == "gzip":
                b = gzip.decompress(b)
            d = json.loads(b.decode("utf-8"))
            with gzip.open(p, "wt", encoding="utf-8") as f:
                json.dump(d, f, ensure_ascii=False)
            return d
        except Exception as e:
            last = e
            time.sleep(1.5 * (i + 1))
    raise last


# ---------------------------------------------------------------------------
# 1. Tra ma danh muc theo cay HIEN TAI cua hang
# ---------------------------------------------------------------------------
def tree():
    """/categories tra ve ca cay danh muc (10 linh vuc / 41 nhom / 150 dong SP).

    Can cai nay vi cot `code` trong catalog/out da cu: mot so dong ghi ma LINH
    VUC thay vi ma nhom (A0011 trong khi URL la .../products/G1108), va 14 dong
    co ma da bi hang doi han (F01101 -> F01102). Doi chieu voi cay song la cach
    duy nhat lay lai dung ma.
    """
    out = {}

    def walk(ns, path):
        for n in ns:
            p = path + [n.get("title") or n.get("name") or ""]
            out[n["id"]] = {"type": n.get("type"), "path": p,
                            "name": p[-1],
                            "kids": [c["id"] for c in (n.get("children") or [])]}
            walk(n.get("children") or [], p)

    walk((api("/categories") or {}).get("data") or [], [])
    return out


def norm(s):
    return re.sub(r"[^a-z0-9]+", "", (s or "").lower())


def resolve(line, tr):
    """Ma danh muc dung cho mot dong: duoi URL -> khop ten chinh xac.

    KHONG dung khop gan dung: 'Contact temperature sensors' tung bi keo sang
    'Contrast sensors' (mot dong SP khac da co bai rieng).
    """
    tail = re.sub(r".*/", "", line.get("url") or "")
    if tr.get(tail, {}).get("type") == "product":
        return tail
    want = norm(line.get("name"))
    for cid, n in tr.items():
        if n["type"] == "product" and norm(n["name"]) == want:
            return cid
    return ""


# ---------------------------------------------------------------------------
# 2. Doc mot dong SP
# ---------------------------------------------------------------------------
# Cot khong phai thong so: gia (phai dang nhap moi thay) va ma dat hang (da co
# cot rieng).
SKIP_COL = {"price", "sap_mara_ean11"}


def paras(s):
    """description tra ve mot khoi van ban dai -> cat thanh doan."""
    s = re.sub(r"(?i)<br\s*/?>", "</p>", s or "")
    out = []
    for p in re.split(r"</p>|</div>|</li>|\n{2,}", s):
        t = txt(p)
        if len(t) > 25:
            out.append(t)
    if not out:
        t = txt(s)
        # description cua Balluff hay dinh lien hai cau ("...application.The
        # sensor brackets...") -> tach lai o cho cham + chu hoa.
        for part in re.split(r"(?<=[a-z\)])\.(?=[A-Z])", t):
            part = part.strip()
            if len(part) > 25:
                out.append(part if part.endswith(".") else part + ".")
    return out[:6]


def fetch_line(code):
    if not code:
        # Dong SP hang da bo khoi danh muc (vd Contact temperature sensors):
        # van ra bai, chi khong co bang ma SP.
        return {"title": "", "prose": [], "bullets": [], "img": "",
                "cols": [], "items": []}
    c = (api("/categories/%s" % code) or {}).get("data") or {}
    cols = [(x["id"], x.get("name") or x["id"])
            for x in (c.get("columns") or [])
            if x.get("id") and x["id"] not in SKIP_COL]
    cols = cols[:MAX_COLS]

    items, page, last = [], 1, 1
    # Tran an toan: 100 trang x 50 ma = 5000 ma cho mot dong. Dong lon nhat cua
    # Balluff (Double-ended cordsets) co 3418 ma -> tran 40 truoc day cat mat.
    while page <= last and page <= 100:
        d = api("/categories/%s/products?page=%d" % (code, page))
        last = (d.get("meta") or {}).get("last_page") or 1
        for p in d.get("data") or []:
            lc = p.get("list_columns") or {}
            vals = []
            for cid, _name in cols:
                v = lc.get(cid)
                if isinstance(v, (list, tuple)):
                    v = ", ".join(str(x) for x in v)
                elif isinstance(v, dict):
                    v = str(v.get("value") or "")
                vals.append(txt("" if v is None else str(v)))
            items.append({"art": (p.get("order_code") or "").strip(),
                          "name": (p.get("label") or "").strip(),
                          "img": p.get("image_url") or p.get("thumbnail_url") or "",
                          "vals": vals})
        page += 1
    return {"title": (c.get("title") or c.get("name") or "").strip(),
            "prose": paras(c.get("description")),
            "bullets": [txt(b) for b in (c.get("bulletpoints") or []) if txt(b)],
            "img": c.get("image_url") or c.get("image_400_url")
                   or c.get("thumbnail_url") or "",
            "cols": [n for _i, n in cols],
            "items": items}


# ---------------------------------------------------------------------------
# 3. Mo ta dai / ngan
# ---------------------------------------------------------------------------
def groups(cols_en, items):
    import collections
    g = [collections.Counter() for _ in cols_en]
    for it in items:
        for i, v in enumerate(it["vals"][:len(cols_en)]):
            if v:
                g[i][v] += 1
    return [(cols_en[i], c) for i, c in enumerate(g) if c]


def fmt(k, cnt, max_vals=6):
    vals = [v for v, _c in cnt.most_common(max_vals)]
    more = len(cnt) - len(vals)
    return "%s: %s%s" % (k, "; ".join(vals),
                         " (+%d tuy chon khac)" % more if more > 0 else "")


def long_desc(cols_en, items, line_en=""):
    low = (line_en or "").lower()
    out = []
    for k, c in groups(cols_en, items):
        if len(c) == 1 and list(c)[0].lower() in low:
            continue
        out.append(fmt(k, c))
    return out


# ---------------------------------------------------------------------------
# 4. Ten tieng Viet + anh
# ---------------------------------------------------------------------------
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


def grab_image(url, line_en, slugify, used):
    if not url:
        return ""
    base = slugify(line_en)[:60].strip("-") or "sp"
    if base in used:
        n = 2
        while "%s-%d" % (base, n) in used:
            n += 1
        base = "%s-%d" % (base, n)
    used.add(base)
    ext = os.path.splitext(url.split("?")[0])[1][:5] or ".png"
    src = os.path.join(SRC_IMG, base + ext)
    try:
        if not (os.path.exists(src) and os.path.getsize(src) > 500):
            # URL anh cua Balluff co dau cach ("Product Detail 400x400")
            import urllib.parse
            req = urllib.request.Request(
                urllib.parse.quote(url, safe=":/?&=#%"),
                headers={"User-Agent": UA, "Referer": SITE + "/"})
            data = urllib.request.urlopen(req, timeout=60).read()
            if len(data) < 500:
                return ""
            with open(src, "wb") as f:
                f.write(data)
        import brandimg
        return os.path.basename(
            brandimg.compose(src, os.path.join(IMG_DIR, base + ".jpg")))
    except Exception as e:
        print("     [anh loi]", repr(e)[:60])
        return ""


def save(ls, done):
    arts, seen = [], set()
    for l in ls:
        u = l.get("url")
        if u in done and u not in seen:
            seen.add(u)
            arts.append(done[u])
    with io.open(CONTENT, "w", encoding="utf-8") as f:
        json.dump({"brand": "17 Balluff", "articles": arts},
                  f, ensure_ascii=False)


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
    tr = tree()
    nofix = []
    for l in ls:
        c = resolve(l, tr)
        l["code"] = c
        if not c:
            nofix.append(l["name"])
    print("cay danh muc hang: %d node | dong tra duoc ma: %d/%d"
          % (len(tr), len(ls) - len(nofix), len(ls)))
    if nofix:
        print("  hang da bo khoi danh muc (bai chi co ten + danh muc): %s"
              % ", ".join(x[:34] for x in nofix))
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
    print("dong SP: %d | bo trung cung trang: %d | bai se co: %d"
          % (len(ls), len(ls) - len(uniq), len(uniq)))
    print("da xong: %d | se lam: %d" % (len(done), len(todo)))
    if not todo:
        return

    # ---- Pha 1: goi het API + tai anh, gom chuoi roi dich MOT LUOT ----
    import prewarm
    raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
    for i, ln in enumerate(todo, 1):
        try:
            d = fetch_line(ln["code"])
        except Exception as e:
            print("  cao %3d/%d %-38s LOI %s"
                  % (i, len(todo), ln["name"][:38], repr(e)[:50]))
            continue
        title = d["title"] or ln["name"]
        d["title"] = title
        d["img_file"] = grab_image(d["img"], title, slugify, used)
        raw[ln["url"]] = d

        long_en = long_desc(d["cols"], d["items"], title)
        bag["key"].add(title)
        bag["key"].update(d["cols"])
        bag["key"].update(it["name"] for it in d["items"] if it["name"])
        bag["value"].update(long_en)
        bag["value"].update(v for it in d["items"] for v in it["vals"] if v)
        bag["prose"].update(d["bullets"][:MAX_BULLETS] or long_en[:MAX_HL])
        bag["prose"].update(d["prose"])
        print("  cao %3d/%d %-38s ma=%-4d cot=%-2d bullet=%-2d %s"
              % (i, len(todo), title[:38], len(d["items"]), len(d["cols"]),
                 len(d["bullets"]), "img" if d["img_file"] else "---"))
        sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, cols, items = d["title"], d["cols"], d["items"]
        long_en = long_desc(cols, items, title)
        short_en = d["bullets"][:MAX_BULLETS] or long_en[:MAX_HL]

        vals = sorted({v for it in items for v in it["vals"] if v})
        n_en = sorted({it["name"] for it in items if it["name"]})
        tr_vals = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        tr_n = dict(zip(n_en, translate_many(n_en, "key"))) if n_en else {}
        for it in items:
            it["name_vi"] = tr_n.get(it["name"], "")
            it["vals_vi"] = [tr_vals.get(v, v) for v in it["vals"]]

        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title, items=items,
            img_file=d["img_file"], cols_en=cols,
            cols_vi=translate_many(cols, "key") if cols else [],
            long_en=long_en,
            long_vi=translate_many(long_en, "value") if long_en else [],
            short_en=short_en,
            short_vi=translate_many(short_en, "prose") if short_en else [],
            prose_en=d["prose"],
            prose_vi=translate_many(d["prose"], "prose") if d["prose"] else [],
            name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-38s ma=%-4d cot=%-2d ts=%-2d %s"
              % (i, len(todo), title[:38], len(items), len(cols),
                 len(long_en), "img" if d["img_file"] else "---"))
        if i % 20 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
