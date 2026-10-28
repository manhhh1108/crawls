# -*- coding: utf-8 -*-
"""IFM (hang 04) - crawl noi dung bai dang cho tung dong san pham.

    python c_ifm.py            # chay tiep nhung dong chua co
    python c_ifm.py 3          # chi 3 dong dau (xem thu)
    python c_ifm.py --force    # lam lai tu dau

Can chay truoc: python ifm_map.py   (ten dong SP -> ma danh muc)

ifm.com chan moi request khong phai trinh duyet (403), nhung ben trong trang co
REST API tra san moi thu duoi dang JSON, khong phai boc HTML:

    /category/<id>/content                  headline, bulletPoints, text
    /category/<id>/productsAndAttributes    viewAttributes + productResults

-> moi dong SP chi can 2 lenh goi, khong phan trang. Van phai chay qua Chrome
that vi goi tu ngoai bi 403 (xem ifm_map.Rest).

Theo 'Template nội dung IFM.docx':
    Mo ta ngan  = bulletPoints o dau trang danh muc
    Mo ta dai   = full thong so gom tu bang ma SP
    Bang ma SP  = Ma SP | Ten SP | cac cot trong viewAttributes

Ket qua: out/04 IFM/content.json  -> build_ifm_excel.py dung tiep.
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

BRAND = "IFM"
SITE = "https://www.ifm.com"
OUT = os.path.join(HERE, "out", "04 IFM")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung IFM")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_COLS = 8
MAX_HL = 8


def txt(s):
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()


def paras(htm):
    """<p>...</p> cua truong 'text' -> danh sach doan van."""
    out = []
    for p in re.split(r"</p>|<br\s*/?>", htm or ""):
        t = txt(p)
        if len(t) > 20:
            out.append(t)
    return out


# ---------------------------------------------------------------------------
# 1. Doc mot danh muc
# ---------------------------------------------------------------------------
def fetch_line(rest, code):
    c = rest.get("/category/%s/content" % code) or {}
    pa = rest.get("/category/%s/productsAndAttributes" % code) or {}

    # Nhieu thuoc tinh KHAC ID nhung CUNG TEN ("Measuring range" x8, moi cai mot
    # don vi: ml/min, l/min, m3/h...). In moi cai mot cot thi bang co 8 cot
    # trung ten va gan nhu rong het -> gop lai thanh MOT cot theo ten.
    groups, order = {}, []
    for a in pa.get("viewAttributes") or []:
        aid, lab = a.get("id"), (a.get("label") or "").strip()
        if not aid or not lab:
            continue
        if lab not in groups:
            groups[lab] = []
            order.append(lab)
        groups[lab].append((aid, a.get("uom") or ""))

    def one(at, aid, uom):
        v = at.get(aid, "")
        if isinstance(v, list):
            v = ", ".join(str(x) for x in v)
        elif isinstance(v, float) and v == int(v):
            v = str(int(v))
        elif v is None:
            v = ""
        v = str(v).strip()
        return "%s %s" % (v, uom) if (v and uom) else v

    items = []
    for p in pa.get("productResults") or []:
        at = p.get("attributes") or {}
        pid = at.get("productId") or at.get("pimId") or ""
        if not pid:
            continue
        vals = []
        for lab in order:
            got = [one(at, aid, uom) for aid, uom in groups[lab]]
            vals.append(" / ".join(x for x in got if x))
        items.append({"art": pid, "name": str(at.get("at_produktbezeichnung") or ""),
                      "vals": vals})
    return c, order, items


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
# 2. Ten tieng Viet + anh
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


DAM = "https://media.ifm.com/dam/%s/original,f_jpeg,w_600,h_600,bc_white/%s.jpeg"


def grab_image(rest, code, line_en, slugify, used):
    """Anh dai dien cua danh muc: /assets/category/<ma> tra ve id anh trong DAM."""
    ids = rest.get("/assets/category/%s" % code)
    if isinstance(ids, dict):
        ids = ids.get("photoAssetIds") or ids.get("assetIds") or []
    aid = ""
    if isinstance(ids, list) and ids:
        a = ids[0]
        aid = a.get("assetId") if isinstance(a, dict) else str(a)
    if not aid:
        return ""
    url = DAM % (aid, code)
    base = slugify(line_en)[:60].strip("-") or "sp"
    if base in used:
        n = 2
        while "%s-%d" % (base, n) in used:
            n += 1
        base = "%s-%d" % (base, n)
    used.add(base)
    src = os.path.join(SRC_IMG, base + ".jpg")
    try:
        if not (os.path.exists(src) and os.path.getsize(src) > 500):
            data = rest.fetch_bytes(url)
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
        json.dump({"brand": "04 IFM", "articles": arts}, f, ensure_ascii=False)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    force = "--force" in sys.argv
    limit = next((int(a) for a in sys.argv[1:] if a.isdigit()), 0)

    from c_pepperl_try import slugify
    from vi_engine import translate_many
    from translate import _flush
    from ifm_map import Rest

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

    # ---- Pha 1: goi het API + tai anh, gom chuoi roi dich MOT LUOT ----
    import prewarm
    raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        rest = Rest(pw)
        try:
            for i, ln in enumerate(todo, 1):
                try:
                    c, cols, items = fetch_line(rest, ln["code"])
                except Exception as e:
                    print("  %3d/%d %-34s LOI %s"
                          % (i, len(todo), ln["name"][:34], str(e)[:40]))
                    continue
                title = (c.get("headline") or "").strip() or ln["name"]
                cols = cols[:MAX_COLS]
                img = grab_image(rest, ln["code"], title, slugify, used)
                bullets = [b for b in (c.get("bulletPoints") or []) if b]
                prose = paras(c.get("text"))
                raw[ln["url"]] = (cols, items, title, bullets, prose, img)
                long_en = long_desc(cols, items, title)
                bag["key"].update(cols)
                bag["key"].add(title)
                bag["key"].update(it["name"] for it in items if it.get("name"))
                bag["value"].update(long_en)
                bag["value"].update(v for it in items for v in it["vals"] if v)
                bag["prose"].update(bullets or prose[:2] or long_en[:MAX_HL])
                bag["prose"].update(prose)
                if i % 20 == 0 or i == len(todo):
                    print("  goi API %d/%d danh muc" % (i, len(todo)))
                    sys.stdout.flush()
        finally:
            rest.close()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    if True:
        if True:
            for i, ln in enumerate(todo, 1):
                if ln["url"] not in raw:
                    continue
                cols, items, title, bullets, prose, img = raw[ln["url"]]
                long_en = long_desc(cols, items, title)
                short_en = bullets or prose[:2] or long_en[:MAX_HL]

                vals = sorted({v for it in items for v in it["vals"] if v})
                n_en = sorted({it["name"] for it in items if it.get("name")})
                tr_long = translate_many(long_en, "value") if long_en else []
                tr_short = translate_many(short_en, "prose") if short_en else []
                tr_cols = translate_many(cols, "key") if cols else []
                tr_prose = translate_many(prose, "prose") if prose else []
                tr_vals = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
                tr_n = dict(zip(n_en, translate_many(n_en, "key"))) if n_en else {}
                for it in items:
                    it["name_vi"] = tr_n.get(it.get("name", ""), "")
                    it["vals_vi"] = [tr_vals.get(v, v) for v in it["vals"]]

                done[ln["url"]] = dict(ln, url_src=ln["url"], items=items,
                                       img_file=img, cols_en=cols, cols_vi=tr_cols,
                                       long_en=long_en, long_vi=tr_long,
                                       short_en=short_en, short_vi=tr_short,
                                       prose_en=prose, prose_vi=tr_prose,
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
