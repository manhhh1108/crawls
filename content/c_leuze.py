# -*- coding: utf-8 -*-
"""Leuze (hang 08) - crawl noi dung bai dang cho tung dong san pham.

    python c_leuze.py            # chay tiep nhung dong chua co
    python c_leuze.py 3          # chi 3 dong dau (xem thu)
    python c_leuze.py --force    # lam lai tu dau

Can chay truoc: python leuze_map.py   (nhanh danh muc -> URL tu sitemap)

leuze.com chan moi request khong phai trinh duyet (403 ke ca robots.txt) nen
moi trang deu phai render bang Chrome that (Playwright). Bu lai shop chay
Shopware: them ?p=1&limit=200 la ra HET ma SP trong mot lan tai, khong phai
bam phan trang.

Theo 'Template nội dung Leuze.docx':
    Mo ta ngan  = doan gioi thieu duoi <h1> cua trang danh muc
    Mo ta dai   = full thong so gom tu product-selector-features cua cac ma SP
    Bang ma SP  = Model | Ma dat hang (Part No.) | Mo ta | cac cot thong so

Ket qua: out/08 Leuze/content.json  -> build_leuze_excel.py dung tiep.
"""
import io
import json
import os
import re
import sys
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

BRAND = "Leuze"
SITE = "https://www.leuze.com"
OUT = os.path.join(HERE, "out", "08 Leuze")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = os.path.join(REPO, "data", "Noi dung Leuze")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

LIMIT = 200            # Shopware: so ma SP tren mot trang
MAX_HL = 8             # Mo ta ngan: 5-8 thong so (khi trang khong co overview)


def txt(s):
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()


# ---------------------------------------------------------------------------
# 1. Doc trang danh muc
# ---------------------------------------------------------------------------
BOX = re.compile(r'<div class="card product-box(.*?)(?=<div class="card product-box|<div class="cms-element-sidebar-filter|</main>)', re.S)
PNAME = re.compile(r'class="product-name"[^>]*title="([^"]*)"')
PNUM = re.compile(r'class="product-number"[^>]*>(.*?)</span>', re.S)
PDESC = re.compile(r'<div class="product-description">(.*?)</div>', re.S)
PFEAT = re.compile(r'<li title="[^"]*">\s*<b>(.*?)</b>(.*?)</li>', re.S)
PIMG = re.compile(r'<img src="([^"]+)"[^>]*class="product-image')
H1 = re.compile(r"<h1[^>]*>(.*?)</h1>", re.S)
# doan gioi thieu nam ngay duoi <h1> cua trang danh muc
LEAD = re.compile(r"</h1>(.*?)(?:<h2|<div class=\"cms-element|Article list)", re.S)
TOTAL = re.compile(r"Results list:\s*([\d,]+)\s*products found")


def parse_items(html):
    out = []
    for blk in BOX.findall(html):
        n = PNAME.search(blk)
        if not n:
            continue
        num = PNUM.search(blk)
        d = PDESC.search(blk)
        im = PIMG.search(blk)
        feats = [[txt(k).rstrip(":"), txt(v)] for k, v in PFEAT.findall(blk)]
        out.append({"name": txt(n.group(1)),
                    "art": txt(num.group(1)) if num else "",
                    "desc": txt(d.group(1)) if d else "",
                    "img": im.group(1) if im else "",
                    "attrs": [kv for kv in feats if kv[0] and kv[1]]})
    return out


def lead_text(html):
    m = LEAD.search(html)
    if not m:
        return ""
    t = txt(m.group(1))
    return t if 20 <= len(t) <= 900 else ""


class Browser(object):
    def __init__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self.b = self._pw.chromium.launch(headless=False, channel="chrome")
        self.pg = self.b.new_page()

    def html(self, url):
        self.pg.goto(url, wait_until="domcontentloaded", timeout=120000)
        self.pg.wait_for_timeout(1200)
        return self.pg.content()

    def fetch_bytes(self, url):
        """Tai file nhi phan. files.leuze.com tra 403 cho urllib VA cho ca
        request context cua Playwright -> phai dieu hinh mot tab phu sang thang
        URL anh roi lay body cua response."""
        pg = self.b.new_page()
        try:
            r = pg.goto(url, wait_until="commit", timeout=60000)
            return r.body() if r and r.ok else b""
        except Exception:
            return b""
        finally:
            pg.close()

    def close(self):
        self.b.close()
        self._pw.stop()


def fetch_line(br, url):
    base = url.rstrip("/")
    first = br.html("%s?p=1&limit=%d" % (base, LIMIT))
    m = H1.search(first)
    h1 = txt(m.group(1)) if m else ""
    t = TOTAL.search(first)
    total = int(t.group(1).replace(",", "")) if t else None
    lead = lead_text(first)
    items = parse_items(first)
    # qua LIMIT thi doc tiep cac trang sau
    page = 2
    while total and len(items) < total and page <= 20:
        got = parse_items(br.html("%s?p=%d&limit=%d" % (base, page, LIMIT)))
        if not got:
            break
        items += got
        page += 1
    return items, h1, lead, total


# ---------------------------------------------------------------------------
# 2. Mo ta dai / ngan
# ---------------------------------------------------------------------------
def groups(items):
    import collections
    g = collections.defaultdict(collections.Counter)
    for it in items:
        for k, v in it.get("attrs", []):
            g[k][v] += 1
    return sorted(g.items(), key=lambda kv: -sum(kv[1].values()))


def fmt(k, cnt, max_vals=6):
    vals = [v for v, _c in cnt.most_common(max_vals)]
    more = len(cnt) - len(vals)
    return "%s: %s%s" % (k, "; ".join(vals),
                         " (+%d tuy chon khac)" % more if more > 0 else "")


def long_desc(items, line_en=""):
    low = (line_en or "").lower()
    out = []
    for k, c in groups(items):
        if len(c) == 1 and list(c)[0].lower() in low:
            continue
        out.append(fmt(k, c))
    return out


def attr_columns(items, max_cols=8):
    return [k for k, _c in groups(items)][:max_cols]


# ---------------------------------------------------------------------------
# 3. Ten tieng Viet + anh
# ---------------------------------------------------------------------------
def is_code_name(s):
    """Ten dong SP la ma/ten rieng (MLC 500, BCL 348i) hay cum tu thuong
    ('Rod mounts', 'Bar code tapes')?"""
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


def grab_image(br, items, line_en, slugify, used):
    url = next((it.get("img") for it in items if it.get("img")), "")
    if not url:
        return ""
    base = slugify(line_en)[:60].strip("-") or "sp"
    if base in used:
        n = 2
        while "%s-%d" % (base, n) in used:
            n += 1
        base = "%s-%d" % (base, n)
    used.add(base)
    ext = os.path.splitext(url.split("?")[0])[1][:5] or ".jpg"
    src = os.path.join(SRC_IMG, base + ext)
    try:
        if not (os.path.exists(src) and os.path.getsize(src) > 500):
            data = br.fetch_bytes(url)
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
        json.dump({"brand": "08 Leuze", "articles": arts}, f, ensure_ascii=False)


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
    print("dong SP: %d | bo trung cung trang: %d | bai se co: %d"
          % (len(ls), len(ls) - len(uniq), len(uniq)))
    print("da xong: %d | se lam: %d" % (len(done), len(todo)))
    if not todo:
        return

    # ---- Pha 1: cao het trang + tai anh, gom chuoi roi dich MOT LUOT ----
    import prewarm
    raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
    br = Browser()
    try:
        for i, ln in enumerate(todo, 1):
            try:
                items, h1, lead, total = fetch_line(br, ln["url"])
            except Exception as e:
                print("  %3d/%d %-34s LOI %s"
                      % (i, len(todo), ln["name"][:34], str(e)[:40]))
                continue
            title = h1 or ln["name"]
            img = grab_image(br, items, title, slugify, used)
            raw[ln["url"]] = (items, title, lead, img)
            cols = attr_columns(items)
            long_en = long_desc(items, title)
            bag["key"].update(cols)
            bag["key"].add(title)
            bag["key"].update(it["desc"] for it in items if it.get("desc"))
            bag["value"].update(long_en)
            bag["value"].update(v for it in items for _k, v in it.get("attrs", []))
            bag["prose"].update([lead] if lead else long_en[:MAX_HL])
            print("  cao %3d/%d %-34s ma=%-4d %s"
                  % (i, len(todo), title[:34], len(items), "img" if img else "---"))
            sys.stdout.flush()
    finally:
        br.close()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    if True:
        for i, ln in enumerate(todo, 1):
            if ln["url"] not in raw:
                continue
            items, title, lead, img = raw[ln["url"]]
            cols = attr_columns(items)
            long_en = long_desc(items, title)
            short_en = [lead] if lead else long_en[:MAX_HL]

            vals = sorted({v for it in items for _k, v in it.get("attrs", [])})
            d_en = sorted({it.get("desc", "") for it in items if it.get("desc")})
            tr_long = translate_many(long_en, "value") if long_en else []
            tr_short = translate_many(short_en, "prose") if short_en else []
            tr_cols = translate_many(cols, "key") if cols else []
            tr_vals = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
            tr_d = dict(zip(d_en, translate_many(d_en, "key"))) if d_en else {}
            for it in items:
                it["desc_vi"] = tr_d.get(it.get("desc", ""), "")
                it["attrs_vi"] = [[k, tr_vals.get(v, v)]
                                  for k, v in it.get("attrs", [])]

            done[ln["url"]] = dict(ln, url_src=ln["url"], items=items,
                                   img_file=img, cols_en=cols, cols_vi=tr_cols,
                                   long_en=long_en, long_vi=tr_long,
                                   short_en=short_en, short_vi=tr_short,
                                   title_en=title,
                                   name_vi=name_vi_of(title, translate_many))
            print("  %3d/%d %-34s ma=%-4d cot=%-2d ts=%-2d %s"
                  % (i, len(todo), title[:34], len(items), len(cols),
                     len(long_en), "img" if img else "---"))
            if i % 20 == 0:
                _flush(True)
                save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
