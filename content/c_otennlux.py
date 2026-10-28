# -*- coding: utf-8 -*-
"""OTENNLUX Lighting (hang 16) - crawl noi dung bai dang cho tung dong SP.

    python c_otennlux.py            # chay tiep nhung dong chua co
    python c_otennlux.py 3          # chi 3 dong dau (xem thu)
    python c_otennlux.py --force    # lam lai tu dau

Can chay truoc: python cat_map.py otennlux

en.otennlux.com khong chan request thuong -> dung shared.fetcher (curl+cache).
Khac cac hang khac: 156/163 dong danh muc tro den MOT TRANG SAN PHAM (_pNN),
khong phai trang danh muc, nen "dong SP" o day la mot san pham cu the:

    h1                  ten san pham
    div.main            cau gioi thieu ngan ngay duoi h1  -> Mo ta ngan
    ul.ptab-list        Item no. / Color / Product weight ... -> bang thong so
    div.text            bai viet day du (van + bang model + anh minh hoa)

Theo 'Template nội dung OTENNLUX Lighting.docx':
    Mo ta ngan  = overview (div.main)
    Mo ta dai   = full phan thong tin PHIA TREN bang thong so ky thuat
    Thong so    = bang trong div.text + cac thuoc tinh ptab-list

Ket qua: out/16 OTENNLUX Lighting/content.json -> build_otennlux_excel.py.
"""
import io
import json
import os
import re
import sys
import urllib.parse
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

BRAND = "OTENNLUX"
SITE = "https://en.otennlux.com"
OUT = os.path.join(HERE, "out", "16 OTENNLUX Lighting")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = os.path.join(REPO, "data", "Noi dung OTENNLUX Lighting")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_HL = 8

FULLWIDTH = {u"（": "(", u"）": ")", u"，": ",", u"：": ":",
             u"；": ";", u"～": "~", u"／": "/", u"－": "-",
             u"＜": "<", u"＞": ">", u"％": "%", u"、": ", ",
             u"。": ". "}
INLINE = re.compile(r"(?i)</?(span|strong|b|i|u|em|font|a|sub|sup|small)\b[^>]*>")


def plain(s):
    for a, b in FULLWIDTH.items():
        s = s.replace(a, b)
    # trang nay hay dinh chu lai voi nhau ("Weattendedlots of Exhibition")
    return re.sub(r"\s+", " ", s)


def txt(s):
    s = INLINE.sub("", s or "")
    return plain(H.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


# ---------------------------------------------------------------------------
# 1. Boc mot trang
# ---------------------------------------------------------------------------
H1 = re.compile(r"<h1[^>]*>(.*?)</h1>", re.S)
OVERVIEW = re.compile(r'</h1>\s*<div class="main">(.*?)</div>', re.S)
ATTR = re.compile(r"<li><p>(.*?)</p>\s*<span>(.*?)</span></li>", re.S)
DESC = re.compile(r'<div class="text">(.*?)'
                  r'(?:</div>\s*</div>\s*</div>|<div class="pro_about)', re.S)
# Anh dai dien PHAI lay trong khoi div.prom_img. Tren moi trang con co khoi
# "hot products" o thanh ben, cac the anh trong do dung CUNG id
# product_detail_img va nam TRUOC khoi chinh -> tim thang theo id thi 163/163
# bai deu an phai anh cua san pham khac.
PROM = re.compile(r'class="prom_img[^"]*"(.*?)(?:class="prom-right|</h1>)', re.S)
MAIN_IMG = re.compile(r'<img[^>]+src="([^"]+)"')
TABLE = re.compile(r"<table.*?</table>", re.S)
ROW = re.compile(r"<tr.*?</tr>", re.S)
CELL = re.compile(r"<(t[dh])([^>]*)>(.*?)</\1>", re.S)
SPAN = re.compile(r'(rowspan|colspan)\s*=\s*"?(\d+)', re.I)

# Thong tin thuong mai, khong dua vao bai dang ky thuat
SKIP_ATTR = re.compile(r"payment|shipping|lead\s*time|minimum|moq|price\s*term"
                       r"|delivery|packing|port", re.I)

# Doan quang cao lap lai y nguyen o moi trang cua hang - cat bo
BOILER = re.compile(
    r"our (light|product)s? with high quality"
    r"|different certification"
    r"|we ?attended ?lots of exhibition"
    r"|looking forward to meet you at the exhibition"
    r"|otennlux is (a )?manufacturer of"
    r"|we only provide high quality products"
    r"|please contact us to begin the business"
    r"|if you are interested in our products"
    r"|leave a message"
    r"|we have confidence in our products"
    r"|such as emo", re.I)

# Nhan khong co noi dung. Giu lai "Key Features:" / "Application:" vi chung
# chia bo cuc cho mo ta dai, nhung "More information:" (41 bai) va "Other Item:"
# chi la chu dan sang bang/anh phia duoi, dua vao bai thi treo lung.
LABEL_ONLY = re.compile(r"^(more information|other items?|advantage|download"
                        r"|product (details?|images?)|details?)\s*[:：]?$", re.I)


def grid(tbl):
    """<table> -> luoi chu nhat, da tra lai cac o bi gop (rowspan/colspan)."""
    out, hang = [], {}
    for rh in ROW.findall(tbl):
        row = {}
        for c in sorted(hang):
            n, v = hang[c]
            row[c] = v
            if n <= 1:
                del hang[c]
            else:
                hang[c] = [n - 1, v]
        col = 0
        for _tag, attrs, body in CELL.findall(rh):
            while col in row:
                col += 1
            rs = cs = 1
            for k, n in SPAN.findall(attrs):
                if k.lower() == "rowspan":
                    rs = max(1, min(int(n), 80))
                else:
                    cs = max(1, min(int(n), 20))
            v = txt(body)
            for j in range(cs):
                row[col + j] = v
                if rs > 1:
                    hang[col + j] = [rs - 1, v]
            col += cs
        if any(row.values()):
            out.append([row.get(i, "") for i in range(max(row) + 1)])
    if not out:
        return []
    w = max(len(r) for r in out)
    out = [r + [""] * (w - len(r)) for r in out]
    # bo cot rong hoan toan (bang cua ho hay thua 1-2 cot cuoi)
    keep = [i for i in range(w) if any(r[i] for r in out)]
    out = [[r[i] for i in keep] for r in out]
    return [r for r in out if any(r)]


def prose_of(desc_html):
    """Van ban PHIA TREN bang dau tien, da cat doan quang cao.

    Template OTENNLUX: "Lay full phan thong tin phia tren bang thong so ky
    thuat" -> cat o the <table> dau tien; trang nao khong co bang thi lay het
    roi loc theo BOILER.
    """
    head = TABLE.split(desc_html)[0] if TABLE.search(desc_html) else desc_html
    head = INLINE.sub("", head)
    head = re.sub(r"(?i)<br\s*/?>", "</p>", head)
    out = []
    for p in re.split(r"(?i)</p>|</div>|</li>|</h[1-6]>", head):
        t = txt(p)
        if len(t) < 12 or BOILER.search(t) or LABEL_ONLY.match(t):
            continue
        t = re.sub(r"^[•\-\*·]\s*", "", t)   # bo dau bullet •
        if t not in out:
            out.append(t)
    return out[:14]


def attrs_of(html):
    out = []
    for k, v in ATTR.findall(html):
        k, v = txt(k).rstrip(":").strip(), txt(v)
        if not k or not v or SKIP_ATTR.search(k):
            continue
        out.append([k, v])
    return out


def img_full(url):
    """.../xxx_medium.jpg -> .../xxx.jpg (ban 800x800 thay vi 550x550)."""
    return re.sub(r"_(?:small|medium|thumb)(?=\.[a-z]+$)", "", url or "")


def fetch_line(url):
    from shared.fetcher import try_fetch
    s = try_fetch(urllib.parse.quote(url, safe=":/?&=#%"), retries=2)
    if not s:
        return None
    if isinstance(s, bytes):
        s = s.decode("utf-8", "replace")
    m = H1.search(s)
    ov = OVERVIEW.search(s)
    de = DESC.search(s)
    desc = de.group(1) if de else ""
    title = txt(m.group(1)) if m else ""
    pm = PROM.search(s)
    im = MAIN_IMG.search(pm.group(1)) if pm else None
    img = SITE + img_full(im.group(1)) if im else ""
    if not img and title:
        # 7/163 dong tro den trang DANH MUC (_cNN), khong co khoi anh san pham
        # -> lay anh danh muc trong mega-menu, doi chieu bang thuoc tinh alt.
        mm = re.search(r'<img src="(/uploadfile/category/[^"]+)" alt="\s*%s\s*"'
                       % re.escape(title), s, re.I)
        if mm:
            img = SITE + img_full(mm.group(1))
    models = []
    for tbl in TABLE.findall(desc):
        rows = grid(tbl)
        if len(rows) >= 2 and len(rows[0]) >= 2:
            models.append({"head": rows[0], "rows": rows[1:]})
    return {"h1": title,
            "overview": txt(ov.group(1)) if ov else "",
            "attrs": attrs_of(s),
            "prose": prose_of(desc),
            "models": models,
            "img": img}


# ---------------------------------------------------------------------------
# 2. Mo ta dai / ngan
# ---------------------------------------------------------------------------
def code_col(t):
    for i, c in enumerate(t["head"]):
        if re.search(r"item\s*no|model|part\s*no|order\s*code|type", c, re.I):
            return i
    return 0


def model_codes(models):
    out, seen = [], set()
    for t in models:
        i = code_col(t)
        for r in t["rows"]:
            if i < len(r) and r[i]:
                c = re.sub(r"\s+", "", r[i])
                if c not in seen and len(c) <= 40:
                    seen.add(c)
                    out.append(c)
    return out


def model_cells(models):
    out = set()
    for t in models:
        ci = code_col(t)
        for r in t["rows"]:
            for i, c in enumerate(r):
                if c and i != ci:
                    out.add(c)
    return out


def long_desc(attrs, line_en=""):
    low = (line_en or "").lower()
    out = []
    for k, v in attrs:
        if v and v.lower() not in low:
            out.append("%s: %s" % (k, v))
    return out


# ---------------------------------------------------------------------------
# 3. Ten tieng Viet + anh
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
    import urllib.request
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
            req = urllib.request.Request(
                urllib.parse.quote(url, safe=":/?&=#%"),
                headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64;"
                                       " x64) Chrome/131.0.0.0",
                         "Referer": SITE + "/"})
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
        json.dump({"brand": "16 OTENNLUX Lighting", "articles": arts},
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
    for i, ln in enumerate(todo, 1):
        try:
            d = fetch_line(ln["url"])
        except Exception as e:
            d = None
            print("     LOI", repr(e)[:60])
        if not d:
            print("  cao %3d/%d %-40s KHONG TAI DUOC"
                  % (i, len(todo), ln["name"][:40]))
            continue
        title = d["h1"] or ln["name"]
        d["title"] = title
        d["img_file"] = grab_image(d["img"], title, slugify, used)
        raw[ln["url"]] = d

        long_en = long_desc(d["attrs"], title)
        bag["key"].add(title)
        bag["key"].update(k for k, _v in d["attrs"])
        bag["key"].update(c for t in d["models"] for c in t["head"])
        bag["value"].update(long_en)
        bag["value"].update(v for _k, v in d["attrs"])
        bag["value"].update(model_cells(d["models"]))
        bag["prose"].update([d["overview"]] if d["overview"] else [])
        bag["prose"].update(d["prose"])
        print("  cao %3d/%d %-40s ts=%-2d bang=%-2d van=%-2d %s"
              % (i, len(todo), title[:40], len(d["attrs"]), len(d["models"]),
                 len(d["prose"]), "img" if d["img_file"] else "---"))
        sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, attrs, models = d["title"], d["attrs"], d["models"]
        long_en = long_desc(attrs, title)
        short_en = ([d["overview"]] if d["overview"]
                    else d["prose"][:2] or long_en[:MAX_HL])

        keys = sorted({k for k, _v in attrs}
                      | {c for t in models for c in t["head"]})
        vals = sorted({v for _k, v in attrs} | model_cells(models))
        tr_k = dict(zip(keys, translate_many(keys, "key"))) if keys else {}
        tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        attrs_vi = [[tr_k.get(k, k), tr_v.get(v, v)] for k, v in attrs]
        models_vi = []
        for t in models:
            ci = code_col(t)
            models_vi.append({
                "head": [tr_k.get(c, c) for c in t["head"]],
                "rows": [[c if i == ci else tr_v.get(c, c)
                          for i, c in enumerate(r)] for r in t["rows"]]})

        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title, img_file=d["img_file"],
            codes=model_codes(models),
            attrs_en=attrs, attrs_vi=attrs_vi,
            models_en=models, models_vi=models_vi,
            long_en=long_en,
            long_vi=translate_many(long_en, "value") if long_en else [],
            short_en=short_en,
            short_vi=translate_many(short_en, "prose") if short_en else [],
            prose_en=d["prose"],
            prose_vi=translate_many(d["prose"], "prose") if d["prose"] else [],
            name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-40s ts=%-2d ma=%-3d %s"
              % (i, len(todo), title[:40], len(attrs),
                 len(done[ln["url"]]["codes"]),
                 "img" if d["img_file"] else "---"))
        if i % 20 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
