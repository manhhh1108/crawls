# -*- coding: utf-8 -*-
"""wenglor (hang 06) - crawl noi dung bai dang cho tung dong san pham.

    python c_wenglor.py            # chay tiep nhung dong chua co
    python c_wenglor.py 3          # chi 3 dong dau (xem thu)
    python c_wenglor.py --force    # lam lai tu dau

Can chay truoc: python wenglor_map.py   (nhanh danh muc -> URL tu sitemap)

wenglor.com KHONG chan curl -> tai bang HTTP thuong, khong can Chrome. Trang
danh muc render san ca the san pham, phan trang ?q=:relevance&page=N (12 ma mot
trang, pageSize khong doi duoc).

Theo 'Template nội dung wenglor.docx':
    Mo ta ngan  = doan gioi thieu duoi <h1> cua trang danh muc
    Mo ta dai   = full thong so gom tu featureList cua cac ma SP
    Bang ma SP  = Ma SP | Ten SP | cac cot thong so (khac nhau theo tung dong)

Ket qua: out/06 wenglor/content.json  -> build_wenglor_excel.py dung tiep.
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

from shared.fetcher import try_fetch                 # noqa: E402

BRAND = "wenglor"
SITE = "https://www.wenglor.com"
OUT = os.path.join(HERE, "out", "06 wenglor")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = os.path.join(REPO, "data", "Noi dung wenglor")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_PAGE = 60                  # 12 ma mot trang -> toi da 720 ma mot dong
MIN_HL, MAX_HL = 5, 8          # Mo ta ngan: 5-8 thong so (khi khong co overview)


def txt(s):
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()


# ---------------------------------------------------------------------------
# 1. Doc trang danh muc
# ---------------------------------------------------------------------------
# Cat the san pham bang split chu khong bang regex co lookahead: trang chi co
# DUNG MOT ma SP (vd Evaluation Technology) thi khong co dau ket thuc nao de
# lookahead bam vao -> truoc day bi tra ve 0 ma.
ITEM_SPLIT = '<div class="product__list--item">'
ITEM_END = ('<div class="product__listing--footer', '</section>', '<footer')
HREF = re.compile(r'href="(/[^"]*?/p/[^"]+)"')
CODE = re.compile(r'<span class="code">(.*?)</span>(.*?)</h2>', re.S)
IMG = re.compile(r'<img src="([^"]+)"')
FEAT = re.compile(r'<span class="feature-name">(.*?)</span>\s*'
                  r'<span class="feature-value">(.*?)</span>', re.S)
H1 = re.compile(r'<h1[^>]*>(.*?)</h1>', re.S)
DESC = re.compile(r'<div class="description[^"]*">(.*?)</div>', re.S)
NEXT = re.compile(r'<li class="pagination-next">\s*<a href="([^"]+)"')


def blocks(html):
    for blk in html.split(ITEM_SPLIT)[1:]:
        for end in ITEM_END:
            blk = blk.split(end)[0]
        yield blk


def parse_items(html):
    out = []
    for blk in blocks(html):
        c = CODE.search(blk)
        if not c:
            continue
        a = HREF.search(blk)
        im = IMG.search(blk)
        src = im.group(1) if im else ""
        feats = [[txt(k), txt(v)] for k, v in FEAT.findall(blk)]
        out.append({"art": txt(c.group(1)), "name": txt(c.group(2)),
                    "url": SITE + a.group(1) if a else "",
                    "img": SITE + src if src.startswith("/") else src,
                    "attrs": [kv for kv in feats if kv[0] and kv[1]]})
    return out


def fetch_line(url):
    """Tai het cac trang cua mot dong SP (bam 'next' cho den khi het)."""
    items, seen, h1, desc = [], set(), "", ""
    u = url
    for _ in range(MAX_PAGE):
        html = try_fetch(u, retries=3, timeout=60)
        if not html:
            break
        if not h1:
            m = H1.search(html)
            h1 = txt(m.group(1)) if m else ""
            d = DESC.search(html)
            desc = txt(d.group(1)) if d else ""
        got = parse_items(html)
        new = [x for x in got if x["art"] not in seen]
        for x in new:
            seen.add(x["art"])
        items += new
        n = NEXT.search(html)
        if not new or not n:
            break
        nxt = H.unescape(n.group(1))
        u = SITE + nxt if nxt.startswith("/") else nxt
    return items, h1, desc


# ---------------------------------------------------------------------------
# 2. Mo ta dai / ngan tu featureList
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
    """Full thong so, bo nhom chi lap lai ten dong SP (template: 'loai cac
    thong so trung ten dong')."""
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
    """Ten dong SP la ma/ten rieng (MLSL, OPT, M2SL) hay cum tu thuong?

    Khong dung is_proper_name() cua c_sick.py: ham do coi moi cum 2 tu la ten
    rieng, ma ben wenglor rat nhieu dong SP dat ten 2 tu ("Reflex Sensors",
    "Fork Sensors") - deu phai dich.
    """
    s = (s or "").strip()
    if not s:
        return False
    if " " not in s:                       # MLSL, OPT, MLSLxxxS40
        return bool(re.search(r"[A-Z]{2,}", s)) or bool(re.search(r"\d", s))
    return not re.search(r"[a-z]", s)      # cum nhieu tu toan HOA -> giu nguyen


def name_vi_of(name, translate_many):
    if is_code_name(name):
        return name
    from vi_engine import lookup
    return lookup(name) or translate_many([name], "key")[0]


def grab_image(items, line_en, slugify, used):
    import urllib.request
    url = next((it.get("img") for it in items if it.get("img")), "")
    if not url:
        return ""
    base = slugify(line_en)[:60].strip("-") or "sp"
    if base in used:                       # ten dong SP trung - khong de len nhau
        n = 2
        while "%s-%d" % (base, n) in used:
            n += 1
        base = "%s-%d" % (base, n)
    used.add(base)
    ext = os.path.splitext(url.split("?")[0])[1][:5] or ".png"
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
        json.dump({"brand": "06 wenglor", "articles": arts}, f, ensure_ascii=False)


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

    # ---- Pha 1: cao HET trang (HTTP co cache tren dia nen chay lai rat nhanh),
    # gom toan bo chuoi roi dich MOT LUOT. Xem prewarm.py.
    import prewarm
    raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
    for i, ln in enumerate(todo, 1):
        items, h1, desc = fetch_line(ln["url"])
        title = h1 or ln["name"]
        raw[ln["url"]] = (items, title, desc)
        cols = attr_columns(items)
        long_en = long_desc(items, title)
        bag["key"].update(cols)
        bag["key"].update(it["name"] for it in items if it.get("name"))
        bag["key"].add(title)
        bag["value"].update(long_en)
        bag["value"].update(v for it in items for _k, v in it.get("attrs", []))
        bag["prose"].update([desc] if desc else long_en[:MAX_HL])
        if i % 20 == 0 or i == len(todo):
            print("  cao %d/%d trang danh muc" % (i, len(todo)))
            sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        items, title, desc = raw[ln["url"]]
        cols = attr_columns(items)
        long_en = long_desc(items, title)
        # Mo ta ngan = doan overview cua trang; khong co thi lay 5-8 thong so dau
        short_en = [desc] if desc else long_en[:MAX_HL]
        img = grab_image(items, title, slugify, used)

        vals = sorted({v for it in items for _k, v in it.get("attrs", [])})
        n_en = sorted({it["name"] for it in items if it.get("name")})
        tr_long = translate_many(long_en, "value") if long_en else []
        tr_short = translate_many(short_en, "prose") if short_en else []
        tr_cols = translate_many(cols, "key") if cols else []
        tr_vals = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        tr_n = dict(zip(n_en, translate_many(n_en, "key"))) if n_en else {}
        for it in items:
            it["name_vi"] = tr_n.get(it.get("name", ""), "")
            it["attrs_vi"] = [[k, tr_vals.get(v, v)] for k, v in it.get("attrs", [])]

        done[ln["url"]] = dict(ln, url_src=ln["url"], items=items, img_file=img,
                               cols_en=cols, cols_vi=tr_cols,
                               long_en=long_en, long_vi=tr_long,
                               short_en=short_en, short_vi=tr_short,
                               title_en=title,
                               name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-34s ma=%-4d cot=%-2d ts=%-2d %s"
              % (i, len(todo), title[:34], len(items), len(cols), len(long_en),
                 "img" if img else "---"))
        if i % 5 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
