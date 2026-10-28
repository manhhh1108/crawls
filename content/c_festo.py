# -*- coding: utf-8 -*-
"""Festo (hang 10) - crawl noi dung bai dang cho tung dong san pham.

    python c_festo.py            # chay tiep nhung dong chua co
    python c_festo.py 3          # chi 3 dong dau (xem thu)
    python c_festo.py --force    # lam lai tu dau
    python c_festo.py --pages 5  # chi lay 5 trang ma SP moi dong (xem thu)

Can chay truoc: python cat_map.py festo

festo.com dung Akamai, tra 403 cho curl. Nhung shared.browser.Browser mo mot
Chrome that roi goi fetch() TU BEN TRONG trang -> chi mat ~0.9s/trang va co
cache tren dia, nhanh hon nhieu so voi dieu hinh (goto) tung trang.

Trang mot dong SP la Remix, nhung du lieu nam san trong `_ROUTER_DATA = {...}`:

    productSeriesData.name/shortCode        ten dong + ma series
    productSeriesData.advantageText[]       gach dau dong overview
    productSeriesData.festoGalleryImages    anh 1056x1024
    ssrArticleListData.facets               cac thong so ky thuat + tuy chon
    ssrArticleListData.articles             orderCode / code / name
    ssrArticleListData.pagination           10 ma SP mot trang, ?page=N

Theo 'Template nội dung Festo.docx':
    Mo ta ngan  = cac gach dau dong trong overview (advantageText)
    Mo ta dai   = full thong so (facets), bo cai trung ten dong
    Bang ma SP  = Model | Ma dat hang (Part No.) | Mo ta

Ket qua: out/10 Festo/content.json -> build_festo_excel.py dung tiep.
"""
import io
import json
import os
import re
import sys
import time
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

BRAND = "Festo"
SITE = "https://www.festo.com"
BASE = SITE + "/us/en"
OUT = os.path.join(HERE, "out", "10 Festo")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = os.path.join(REPO, "data", "Noi dung Festo")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_COLS = 10
MAX_HL = 8
MAX_PAGES = 60         # tran an toan: 60 trang = 600 ma SP cho mot dong
RESTART_EVERY = 120   # mo lai Chrome sau moi 120 dong (tranh ro ri bo nho)

MARK = "_ROUTER_DATA = "
# Hai "facet" dau khong phai thong so ky thuat ma la bo loc tiep thi, xuat hien
# y nguyen o moi dong SP -> bo di.
SKIP_FACET = re.compile(r"^(fastDelivery|Specific catalogues)$", re.I)


def txt(s):
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()


# ---------------------------------------------------------------------------
# 1. Doc mot dong SP
# ---------------------------------------------------------------------------
def ok_page(t):
    """Trang dong SP hop le phai co _ROUTER_DATA.

    festo tra ve HTTP 200 kem mot vo micro-frontend 4.5KB ("isMfBootstrapError",
    khong co _ROUTER_DATA) khi bi chan hoac khi trinh duyet da yeu. Lan chay dau
    27 dong SP an phai vo nay va NO DA VAO CACHE -> moi lan chay lai deu doc lai
    tu cache va 27 bai do rong vinh vien. Kiem o day de Browser.api khong nhan.
    """
    return bool(t) and MARK in t


def router_data(html):
    i = html.find(MARK)
    if i < 0:
        return None
    try:
        d, _ = json.JSONDecoder().raw_decode(html[i + len(MARK):])
    except ValueError:
        return None
    pg = [v for k, v in (d.get("loaderData") or {}).items()
          if k.endswith("/page")]
    return pg[0] if pg else None


def page_url(url, page):
    u = url if url.startswith("http") else BASE + url
    sep = "&" if "?" in u else "?"
    return u if page == 0 else "%s%spage=%d" % (u, sep, page)


def fetch_line(br, url, max_pages=MAX_PAGES):
    pg = router_data(br.api(page_url(url, 0), accept=ok_page))
    if pg is None:
        return None
    ps = pg.get("productSeriesData") or {}
    al = pg.get("ssrArticleListData") or {}

    cols = []
    for f in al.get("facets") or []:
        name = txt(f.get("name"))
        if not name or SKIP_FACET.match(name):
            continue
        vals = [txt(v.get("name")) for v in (f.get("values") or [])]
        vals = [v for v in vals if v]
        if vals:
            cols.append((name, vals))
    cols = cols[:MAX_COLS]

    def arts(al_):
        out = []
        for a in al_.get("articles") or []:
            oc = (a.get("orderCode") or "").strip()
            cd = (a.get("code") or "").strip()
            if not (oc or cd):
                continue
            out.append({"art": cd, "name": oc, "desc": txt(a.get("name")),
                        "img": a.get("imageUrl") or ""})
        return out

    items = arts(al)
    pag = al.get("pagination") or {}
    total = pag.get("totalPages") or 1
    for p in range(1, min(total, max_pages)):
        d2 = router_data(br.api(page_url(url, p), accept=ok_page))
        if d2 is None:
            break
        got = arts(d2.get("ssrArticleListData") or {})
        if not got:
            break
        items += got

    gi = (ps.get("festoGalleryImages") or {}).get("imageData") or []
    img = gi[0].get("src") if gi else ""
    if not img:
        for im in ps.get("images") or []:
            if im.get("url", "").lower().endswith((".jpg", ".png", ".jpeg")):
                img = im["url"]
                break

    bullets = [txt(x.get("longText")) for x in (ps.get("advantageText") or [])]
    bullets += [txt(x.get("longText")) for x in (ps.get("specialText") or [])]
    short = (ps.get("name") or "").strip()
    code = (ps.get("shortCode") or "").strip()
    title = ("%s %s" % (short, code)).strip() if code and code not in short \
        else short
    return {"title": title, "series": code, "cols": cols, "items": items,
            "bullets": [b for b in bullets if b],
            "prose": [t for t in [txt(ps.get("summary")),
                                  txt(ps.get("description"))] if len(t) > 25],
            "img": img,
            "total": pag.get("totalCount") or len(items)}


# ---------------------------------------------------------------------------
# 2. Mo ta dai / ngan
# ---------------------------------------------------------------------------
def fmt(k, vals, max_vals=8):
    v = vals[:max_vals]
    more = len(vals) - len(v)
    return "%s: %s%s" % (k, "; ".join(v),
                         " (+%d tuy chon khac)" % more if more > 0 else "")


def long_desc(cols, line_en=""):
    """facets -> 'Piston diameter [mm]: 32; 40; 50...'. Bo facet chi co mot gia
    tri ma gia tri do da nam trong ten dong SP."""
    low = (line_en or "").lower()
    out = []
    for k, vals in cols:
        if len(vals) == 1 and vals[0].lower() in low:
            continue
        out.append(fmt(k, vals))
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


def name_vi_of(name, series, translate_many):
    """'Standards-based cylinder DNC' -> dich phan ten, giu nguyen ma series."""
    base = name
    if series and base.endswith(series):
        base = base[:-len(series)].strip()
    if not base or is_code_name(base):
        return name
    from vi_engine import lookup
    vi = lookup(base) or translate_many([base], "key")[0]
    return ("%s %s" % (vi, series)).strip() if series else vi


def fetch_bytes(br, url):
    """Tai anh. festo.com/media/catalog nam sau Akamai, tra 403 cho urllib ->
    mo mot tab phu dieu hinh sang thang URL anh roi lay body cua response."""
    pg = br.ctx.new_page()
    try:
        r = pg.goto(url, wait_until="commit", timeout=60000)
        return r.body() if r and r.ok else b""
    except Exception:
        return b""
    finally:
        pg.close()


def grab_image(br, url, line_en, slugify, used):
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
            data = fetch_bytes(br, url)
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
        json.dump({"brand": "10 Festo", "articles": arts}, f, ensure_ascii=False)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    force = "--force" in sys.argv
    limit = next((int(a) for a in sys.argv[1:] if a.isdigit()), 0)
    max_pages = MAX_PAGES
    if "--pages" in sys.argv:
        max_pages = int(sys.argv[sys.argv.index("--pages") + 1])

    from c_pepperl_try import slugify
    from vi_engine import translate_many
    from translate import _flush
    from shared.browser import Browser

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
    # Chay dau tien Chrome chet o dong 170/1206 ("Target page, context or
    # browser has been closed") va 1036 dong con lai bao loi het luot ma khong
    # cao duoc gi. Nen giu trinh duyet trong mot o de mo lai duoc giua duong,
    # va mo lai chu dong theo tung lo de khong tich luy ro ri bo nho.
    bw = {"br": Browser(BASE), "n": 0}

    def restart(why):
        print("  [mo lai trinh duyet: %s]" % why)
        sys.stdout.flush()
        try:
            bw["br"].close()
        except Exception:
            pass
        time.sleep(3)
        bw["br"] = Browser(BASE)
        bw["n"] = 0

    def crawl(ln):
        """Cao mot dong; neu trinh duyet chet thi mo lai roi cao lai dong do."""
        for k in range(2):
            try:
                return fetch_line(bw["br"], ln["url"], max_pages), ""
            except Exception as e:
                msg = repr(e)
                if k == 0 and re.search(r"closed|crash|disconnect|Target",
                                        msg, re.I):
                    restart(msg[-44:])
                    continue
                return None, msg[-50:]

    try:
        for i, ln in enumerate(todo, 1):
            if bw["n"] >= RESTART_EVERY:
                restart("da cao %d dong" % bw["n"])
            bw["n"] += 1
            d, err = crawl(ln)
            if err:
                print("  cao %3d/%d %-36s LOI %s"
                      % (i, len(todo), ln["name"][:36], err))
                sys.stdout.flush()
                continue
            if d is None:
                print("  cao %3d/%d %-36s KHONG CO DU LIEU"
                      % (i, len(todo), ln["name"][:36]))
                continue
            title = d["title"] or ln["name"]
            d["title"] = title
            d["img_file"] = grab_image(bw["br"], d["img"], title, slugify, used)
            raw[ln["url"]] = d

            long_en = long_desc(d["cols"], title)
            bag["key"].add(title)
            bag["key"].update(k for k, _v in d["cols"])
            bag["key"].update(it["desc"] for it in d["items"] if it["desc"])
            bag["value"].update(long_en)
            bag["value"].update(v for _k, vs in d["cols"] for v in vs)
            bag["prose"].update(d["bullets"] or d["prose"] or long_en[:MAX_HL])
            if i % 10 == 0 or i <= 20:
                print("  cao %3d/%d %-36s ma=%-4d cot=%-2d bullet=%-2d %s"
                      % (i, len(todo), title[:36], len(d["items"]),
                         len(d["cols"]), len(d["bullets"]),
                         "img" if d["img_file"] else "---"))
                sys.stdout.flush()
    finally:
        try:
            bw["br"].close()
        except Exception:
            pass
    # Pha 1 khong dich gi, chi tai trang va anh (da vao cache dia) -> chay rieng
    # duoc trong luc mot hang khac dang dich, khong tranh cache dich.
    if "--crawl-only" in sys.argv:
        print("\n--crawl-only: da cao %d/%d dong vao cache, dung o day."
              % (len(raw), len(todo)))
        return
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, cols, items = d["title"], d["cols"], d["items"]
        long_en = long_desc(cols, title)
        short_en = d["bullets"] or d["prose"] or long_en[:MAX_HL]

        cnames = [k for k, _v in cols]
        d_en = sorted({it["desc"] for it in items if it["desc"]})
        tr_d = dict(zip(d_en, translate_many(d_en, "key"))) if d_en else {}
        for it in items:
            it["desc_vi"] = tr_d.get(it["desc"], "")

        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title, series=d["series"],
            items=items, img_file=d["img_file"],
            cols_en=cnames,
            cols_vi=translate_many(cnames, "key") if cnames else [],
            long_en=long_en,
            long_vi=translate_many(long_en, "value") if long_en else [],
            short_en=short_en,
            short_vi=translate_many(short_en, "prose") if short_en else [],
            total=d["total"],
            name_vi=name_vi_of(title, d["series"], translate_many))
        if i % 10 == 0 or i <= 20:
            print("  %3d/%d %-36s ma=%-4d cot=%-2d ts=%-2d %s"
                  % (i, len(todo), title[:36], len(items), len(cols),
                     len(long_en), "img" if d["img_file"] else "---"))
        if i % 50 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
