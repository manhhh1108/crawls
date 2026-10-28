# -*- coding: utf-8 -*-
"""Mitsubishi Electric (hang 07) - crawl noi dung bai dang cho tung dong SP.

    python c_mitsubishi.py            # chay tiep nhung dong chua co
    python c_mitsubishi.py 3          # chi 3 dong dau (xem thu)
    python c_mitsubishi.py --force    # lam lai tu dau

Can chay truoc: python mitsu_map.py   (ten dong SP -> URL trang san pham)

mitsubishielectric.com/fa KHONG chan curl -> tai bang HTTP thuong.

Khac 3 hang kia: trang dong SP khong co bang ma SP san. Ma model nam rai trong
bai duoi dang link
    <a href="/fa/products/faspec/point.page?kisyu=/plcr&formNm=RD78G4" class="model">RD78G4</a>
va chinh trang faspec do moi la bang thong so cua model:
    Model name | Overview | Conformance standard | Sales status

Theo 'Template nội dung Mitsubishi Electric.docx':
    Mo ta ngan  = doan Features dau trang (dich sang tieng Viet)
    Mo ta dai   = full thong so gom tu bang faspec cua cac model
    Bang ma SP  = bang specifications (Model + cac cot cua faspec)

Ket qua: out/07 Mitsubishi Electric/content.json
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

import out_dir                                         # noqa: E402

from shared.fetcher import try_fetch                 # noqa: E402

BRAND = "Mitsubishi Electric"
SITE = "https://www.mitsubishielectric.com"
OUT = os.path.join(HERE, "out", "07 Mitsubishi Electric")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung Mitsubishi Electric")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_MODELS = 120        # so model doc bang faspec cho mot dong SP
MAX_HL = 8


def txt(s):
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()


# ---------------------------------------------------------------------------
# 1. Trang dong SP: doan Features + danh sach model
# ---------------------------------------------------------------------------
H1 = re.compile(r"<h1[^>]*>(.*?)</h1>", re.S)
H2 = re.compile(r"<h2[^>]*>(.*?)</h2>", re.S)
PARA = re.compile(r'<p class="c-text"[^>]*>(.*?)</p>', re.S)
MODEL = re.compile(r'<a[^>]+href="(/fa/products/faspec/[^"]+)"[^>]*class="[^"]*\bmodel\b[^"]*"[^>]*>(.*?)</a>'
                   r'(?:\s*<span class="storage[^"]*">(.*?)</span>)?', re.S)
ALLIMG = re.compile(r'<img[^>]+src="([^"]+)"')
# Anh dung chung cua ca site (logo, icon, anh bia danh muc) - khong phai anh SP
IMG_SKIP = ("/fa/shared/", "/assets/img/", "/common/img/", "logo", "icon_",
            "banner", "/img/mv_", "img_mv-")
FASPEC_ROW = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S)
FASPEC_CELL = re.compile(r"<t[hd][^>]*>(.*?)</t[hd]>", re.S)


TABLE = re.compile(r"<table(.*?)</table>", re.S)
TROW = re.compile(r"<tr[^>]*>(.*?)</tr>", re.S)
TCELL = re.compile(r"<t[hd][^>]*>(.*?)</t[hd]>", re.S)
BAD_TABLE = re.compile(r"we apologize for the incon", re.I)


def page_models(html, base):
    """[(url faspec, ma model, ghi chu ngan)] - giu thu tu, bo trung."""
    out, seen = [], set()
    for href, code, note in MODEL.findall(html):
        code = txt(code)
        if not code or code in seen:
            continue
        seen.add(code)
        u = href if href.startswith("http") else SITE + H.unescape(href)
        out.append((u, code, txt(note)))
    return out


def page_specs(html, max_rows=60):
    """Bang 'Specifications' tren trang dong SP -> [(ten thong so, gia tri)].

    Chi 22/332 dong cua Mitsubishi co danh sach model (link faspec); phan lon
    chi co bang thong so dang 'Function | Specifications' hoac 'Item |
    Description' - dung chinh bang do lam bang thong so cua bai, giong cach da
    lam voi nhung dong SICK khong co bang ma.
    """
    out, seen = [], set()
    for m in TABLE.finditer(html):
        blk = m.group(1)
        if BAD_TABLE.search(blk):
            continue
        rows = TROW.findall(blk)
        if len(rows) < 2:
            continue
        for r in rows:
            c = [txt(x) for x in TCELL.findall(r)]
            c = [x for x in c if x]
            if len(c) < 2:
                continue
            k, v = c[0], " | ".join(c[1:])
            if not k or not v or len(k) > 80 or (k, v) in seen:
                continue
            seen.add((k, v))
            out.append([k, v])
            if len(out) >= max_rows:
                return out
    return out


def lead_paras(html):
    """Doan gioi thieu dau trang (phan 'Features' trong template)."""
    out = []
    for p in PARA.findall(html):
        t = txt(p)
        if len(t) > 40:
            out.append(t)
        if len(out) >= 3:
            break
    return out


def faspec(url):
    """Trang faspec cua mot model -> [(ten thong so, gia tri)]."""
    html = try_fetch(url, retries=2, timeout=45)
    if not html:
        return []
    m = re.search(r"<table(.*?)</table>", html, re.S)
    if not m:
        return []
    rows = []
    for r in FASPEC_ROW.findall(m.group(1)):
        c = [txt(x) for x in FASPEC_CELL.findall(r)]
        if len(c) >= 2 and c[0] and c[1]:
            rows.append([c[0], c[1]])
    return rows


def fetch_line(url):
    html = try_fetch(url, retries=3, timeout=60)
    if not html:
        return [], "", [], "", []
    m = H1.search(html)
    h1 = txt(m.group(1)) if m else ""
    lead = lead_paras(html)
    img = ""
    src = ""
    for cand in ALLIMG.findall(html):
        c = H.unescape(cand)
        low = c.lower()
        if any(k in low for k in IMG_SKIP) or low.endswith(".svg"):
            continue
        src = c
        break
    if src:
        if src.startswith("http"):
            img = src
        elif src.startswith("/"):
            img = SITE + src
        else:                                  # duong dan tuong doi voi trang
            img = url.rsplit("/", 1)[0] + "/" + src.lstrip("./")
            while "/../" in img:
                img = re.sub(r"/[^/]+/\.\./", "/", img)

    items = []
    for u, code, note in page_models(html, url)[:MAX_MODELS]:
        rows = faspec(u)
        d = {k: v for k, v in rows}
        items.append({"art": code, "name": d.get("Model name", code),
                      "desc": note or d.get("Overview", ""),
                      "attrs": [kv for kv in rows if kv[0] != "Model name"],
                      "url": u})
    return items, h1, lead, img, page_specs(html)


# ---------------------------------------------------------------------------
# 2. Mo ta dai / cot bang ma SP
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
    """'MELSEC iQ-R Series', 'RD78G4' -> giu; 'CPU module' -> dich."""
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
            req = urllib.request.Request(url, headers={"User-Agent":
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/131.0.0.0"})
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
        json.dump({"brand": "07 Mitsubishi Electric", "articles": arts},
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

    # ---- Pha 1: cao het trang (HTTP co cache tren dia), gom chuoi roi dich
    # MOT LUOT. Xem prewarm.py.
    import prewarm
    raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
    for i, ln in enumerate(todo, 1):
        items, h1, lead, imgurl, specs = fetch_line(ln["url"])
        title = ln["name"] or h1
        img = grab_image(imgurl, title, slugify, used)
        raw[ln["url"]] = (items, title, lead, img, specs)
        cols = attr_columns(items)
        long_en = long_desc(items, title) or ["%s: %s" % (k, v) for k, v in specs]
        bag["key"].update(cols)
        bag["key"].add(title)
        bag["key"].update(it["desc"] for it in items if it.get("desc"))
        bag["key"].update(k for k, _v in specs)
        bag["value"].update(long_en)
        bag["value"].update(v for it in items for _k, v in it.get("attrs", []))
        bag["value"].update(v for _k, v in specs)
        bag["prose"].update(lead or long_en[:MAX_HL])
        print("  cao %3d/%d %-34s ma=%-4d ts=%-3d %s"
              % (i, len(todo), title[:34], len(items), len(specs),
                 "img" if img else "---"))
        sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        items, title, lead, img, specs = raw[ln["url"]]
        cols = attr_columns(items)
        long_en = long_desc(items, title) or ["%s: %s" % (k, v) for k, v in specs]
        short_en = lead or long_en[:MAX_HL]

        vals = sorted({v for it in items for _k, v in it.get("attrs", [])}
                      | {v for _k, v in specs})
        d_en = sorted({it.get("desc", "") for it in items if it.get("desc")})
        sk_en = sorted({k for k, _v in specs})
        tr_long = translate_many(long_en, "value") if long_en else []
        tr_short = translate_many(short_en, "prose") if short_en else []
        tr_cols = translate_many(cols, "key") if cols else []
        tr_vals = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        tr_d = dict(zip(d_en, translate_many(d_en, "key"))) if d_en else {}
        tr_sk = dict(zip(sk_en, translate_many(sk_en, "key"))) if sk_en else {}
        specs_vi = [[tr_sk.get(k, k), tr_vals.get(v, v)] for k, v in specs]
        for it in items:
            it["desc_vi"] = tr_d.get(it.get("desc", ""), "")
            it["attrs_vi"] = [[k, tr_vals.get(v, v)] for k, v in it.get("attrs", [])]

        done[ln["url"]] = dict(ln, url_src=ln["url"], items=items, img_file=img,
                               cols_en=cols, cols_vi=tr_cols,
                               long_en=long_en, long_vi=tr_long,
                               short_en=short_en, short_vi=tr_short,
                               specs_en=specs, specs_vi=specs_vi,
                               title_en=title,
                               name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-34s ma=%-4d cot=%-2d ts=%-2d %s"
              % (i, len(todo), title[:34], len(items), len(cols), len(long_en),
                 "img" if img else "---"))
        if i % 20 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
