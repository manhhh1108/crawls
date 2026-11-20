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

import out_dir                                         # noqa: E402

BRAND = "OTENNLUX"
SITE = "https://en.otennlux.com"
OUT = os.path.join(HERE, "out", "16 OTENNLUX Lighting")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung OTENNLUX Lighting")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_HL = 8

# Bay dong trong danh muc Ntrend tro vao TRANG DANH MUC cua hang (duoi _cNN)
# chu khong phai trang san pham (_pNN). Da mo tung trang kiem: chi co thanh
# danh muc ben trai, KHONG co khoi danh sach san pham nao - hang chua dang gi
# vao 7 muc nay. Khoi 8 san pham trong trang la "hot products" dung chung cho
# moi trang, khong phai san pham cua danh muc, dung tuong nham ma di cao.
# De lai thi ra 7 bai rong (khong mo ta, khong bang, phai lay logo lam anh).
BO_TRANG_DANH_MUC = {
    "https://en.otennlux.com/flood-light_c39",
    "https://en.otennlux.com/o51-indicator-light_c50",
    "https://en.otennlux.com/o52-indicator-light_c51",
    "https://en.otennlux.com/embedded-led-work-light_c26",
    "https://en.otennlux.com/tube-led-work-light_c24",
    "https://en.otennlux.com/can-type-sound-signal-light_c62",
    "https://en.otennlux.com/rs485-type-sound-signal-light_c64",
}

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

# Nhan khong co noi dung. Giu lai "Key Features:" / "Application:" / "Details:"
# vi chung chia bo cuc cho mo ta dai, nhung "More information:" (41 bai) va
# "Other Item:" chi la chu dan sang bang/anh phia duoi, dua vao bai thi treo
# lung. Truoc day danh sach nay co ca "details?" nen cat mat chu "Details:",
# doan gach dau dong phia sau dinh lien vao doan tren, doc khong ra y.
LABEL_ONLY = re.compile(r"^(more information|other items?|advantage|download"
                        r"|product (details?|images?))\s*[:：]?$", re.I)

# Dong tieu de chia bo cuc: "Details:", "Application :", "Key Features:".
HEADING = re.compile(r"^[A-Za-z][A-Za-z /&]{2,30}\s*[:：]\s*$")
BULLET = re.compile(r"^[•\-\*·]\s*")


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


def nhac_lai_ten(t, title):
    """Dong nay chi la nhac lai ten san pham?

    Trang hang mo dau phan mo ta bang chinh ten san pham, co khi them bot vai
    chu: "IO-LINK IP69 24V AC DC Intelligent signal WARMING light FOR AUTOMATIC
    SYSTEM" so voi ten bai "... Intelligent signal light". So theo tap tu moi
    bat duoc nhung ban them bot nhu vay.
    """
    tu = lambda s: set(re.findall(r"[a-z0-9]+", (s or "").lower()))
    a, b = tu(t), tu(title)
    return bool(b) and len(a & b) >= 0.85 * len(b)


def prose_of(desc_html, title=""):
    """Van ban PHIA TREN bang dau tien, da cat doan quang cao.

    Template OTENNLUX: "Lay full phan thong tin phia tren bang thong so ky
    thuat" -> cat o the <table> dau tien; trang nao khong co bang thi lay het
    roi loc theo BOILER.

    Giu nguyen bo cuc cua trang hang: dong tieu de ("Details:") va dau `*` dau
    moi gach dau dong. Truoc day cat het hai thu do nen ca doan dinh lien vao
    nhau. Dong dau neu chi nhac lai ten san pham thi bo - no da nam o ten bai.
    """
    head = TABLE.split(desc_html)[0] if TABLE.search(desc_html) else desc_html
    head = INLINE.sub("", head)
    head = re.sub(r"(?i)<br\s*/?>", "</p>", head)
    out, than = [], 0
    for p in re.split(r"(?i)</p>|</div>|</li>|</h[1-6]>", head):
        t = txt(p)
        if not t or BOILER.search(t) or LABEL_ONLY.match(t):
            continue
        if HEADING.match(t):
            t = re.sub(r"\s*[:：]\s*$", ":", t)
        else:
            dau = "* " if BULLET.match(t) else ""
            t = BULLET.sub("", t)
            if len(t) < 12:
                continue
            if not than and nhac_lai_ten(t, title):
                continue
            than += 1
            t = dau + t
        if t not in out:
            out.append(t)
    return out[:16]


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
    models = []
    for tbl in TABLE.findall(desc):
        rows = grid(tbl)
        if len(rows) >= 2 and len(rows[0]) >= 2:
            models.append({"head": rows[0], "rows": rows[1:]})
    return {"h1": title,
            "overview": txt(ov.group(1)) if ov else "",
            "attrs": attrs_of(s),
            "prose": prose_of(desc, title),
            "models": models,
            "img": img}


# ---------------------------------------------------------------------------
# 2. Mo ta dai / ngan
# ---------------------------------------------------------------------------
CODE_LBL = re.compile(r"item\s*no|model|part\s*no|order\s*code|type", re.I)


def code_col(t):
    """Chi so cot chua ma SP, -1 neu bang khong co cot nao nhu vay."""
    for i, c in enumerate(t["head"]):
        if CODE_LBL.search(c):
            return i
    return -1


def kv_table(t):
    """Bang DOC: cot trai la ten thong so, cot phai la gia tri.

    114/125 bang cua hang nay la kieu nay (Item No. | OBS1-USB-01 / Brand |
    Otennlux / Materail | Aluminum ...). Truoc day code_col khong khop gi thi
    tra ve 0, thanh ra cot NHAN bi coi la cot ma SP, keo theo hai loi:
      * cot nhan khong duoc dich -> bang hien ra nua Anh nua Viet
      * cac nhan bi gom vao danh sach ma SP -> the tag thanh
        "ItemNo., Brand, Materail, Voltage, Color, ..., OTENNLUX"
    """
    return len(t["head"]) == 2 and code_col(t) < 0


def model_codes(models):
    out, seen = [], set()

    def them(c):
        c = re.sub(r"\s+", "", c or "")
        if c and c not in seen and len(c) <= 40:
            seen.add(c)
            out.append(c)

    for t in models:
        i = code_col(t)
        if i < 0:
            # Bang doc: ma SP la GIA TRI cua dong co nhan "Item No." / "Model".
            # Bo qua ca bang thi mat ma that (vd OBS1-USB-01), ma lay ca cot
            # trai thi vo phai toan nhan.
            for r in t["rows"]:
                if len(r) > 1 and CODE_LBL.search(r[0] or ""):
                    them(r[1])
            continue
        for r in t["rows"]:
            if i < len(r):
                them(r[i])
    return out


def model_cells(models):
    """(ten thong so, gia tri) trong cac bang - de dich dung kieu tung loai.

    Ten thong so dich kieu "key", gia tri dich kieu "value"; dua nham kieu thi
    vi_engine cho ra ban dich khac han.
    """
    keys, vals = set(), set()
    for t in models:
        ci = code_col(t)
        kv = kv_table(t)
        if kv:
            keys.add(t["head"][0])
            vals.add(t["head"][1])
        else:
            keys.update(t["head"])
        for r in t["rows"]:
            for i, c in enumerate(r):
                if not c or i == ci:
                    continue
                (keys if (kv and i == 0) else vals).add(c)
    return keys, vals


def dich_prose(prose_en, translate_many):
    """Dich mo ta dai, giu lai dau `*` va dau hai cham cua dong tieu de.

    Dich ban da go dau `*` roi gan lai, va ep dong tieu de van ket thuc bang
    dau hai cham (may dich hay nuot no, vd "Application :" -> "Ứng dụng").
    """
    if not prose_en:
        return []
    tron = [BULLET.sub("", t) for t in prose_en]
    out = []
    for en, vi in zip(prose_en, translate_many(tron, "prose")):
        if HEADING.match(en) and not vi.rstrip().endswith((":", u"：")):
            vi = vi.rstrip() + ":"
        out.append(("* " + vi) if BULLET.match(en) else vi)
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
          if l.get("url") and l["url"] not in BO_TRANG_DANH_MUC]
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
        mk, mv = model_cells(d["models"])
        bag["key"].add(title)
        bag["key"].update(k for k, _v in d["attrs"])
        bag["key"].update(mk)
        bag["value"].update(long_en)
        bag["value"].update(v for _k, v in d["attrs"])
        bag["value"].update(mv)
        bag["prose"].update([d["overview"]] if d["overview"] else [])
        # Dich ban KHONG co dau `*`, dich xong moi gan lai - de nguyen thi may
        # dich nuot hoac nhan dau sao thanh chu.
        bag["prose"].update(BULLET.sub("", t) for t in d["prose"])
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

        mk, mv = model_cells(models)
        keys = sorted({k for k, _v in attrs} | mk)
        vals = sorted({v for _k, v in attrs} | mv)
        tr_k = dict(zip(keys, translate_many(keys, "key"))) if keys else {}
        tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        attrs_vi = [[tr_k.get(k, k), tr_v.get(v, v)] for k, v in attrs]
        models_vi = []
        for t in models:
            ci = code_col(t)
            if kv_table(t):
                hv = [tr_k.get(t["head"][0], t["head"][0]),
                      tr_v.get(t["head"][1], t["head"][1])]
                rows = [[tr_k.get(r[0], r[0])]
                        + [tr_v.get(c, c) for c in r[1:]] for r in t["rows"]]
            else:
                hv = [tr_k.get(c, c) for c in t["head"]]
                rows = [[c if i == ci else tr_v.get(c, c)
                         for i, c in enumerate(r)] for r in t["rows"]]
            models_vi.append({"head": hv, "rows": rows})
        prose_vi = dich_prose(d["prose"], translate_many)

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
            prose_vi=prose_vi,
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
