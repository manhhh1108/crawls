# -*- coding: utf-8 -*-
"""ESPE Technology (hang 15) - crawl noi dung bai dang cho tung dong san pham.

    python c_espe.py            # chay tiep nhung dong chua co
    python c_espe.py 3          # chi 3 dong dau (xem thu)
    python c_espe.py --force    # lam lai tu dau

Can chay truoc: python cat_map.py espe

espetech.com khong chan request thuong -> dung shared.fetcher (curl + cache).
Luu y URL: nhieu trang cua ho co DAU CACH trong ten file
("accessory/Proximity Sensor.html") nen phai quote truoc khi goi.

Moi trang la MOT dong SP (vd "Safety Relay Module", model "ESR Series"), noi
dung chia thanh cac khoi xep gap `div.pr22lie`:

    p.pr22biao     tieu de khoi ("Technical Parameters", "Option for Brackets")
    div.chanxiang  than khoi (bang hoac anh)

Theo 'Template nội dung ESPE Technology.docx':
    Mo ta ngan  = Product Feature, giu nguyen danh sach 1. 2. 3.
    Mo ta dai   = bang chon model ("Sau day la thong tin de chon SP phu hop")
    Bang thong so = bang Technical Parameters

Ket qua: out/15 ESPE Technology/content.json -> build_espe_excel.py dung tiep.
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

BRAND = "ESPE"
SITE = "https://www.espetech.com"
OUT = os.path.join(HERE, "out", "15 ESPE Technology")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung ESPE Technology")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_HL = 8


# Trang ESPE do nguoi Trung Quoc soan nen lan dau cau toan rong: '（mm）',
# 'DC24V±10%，', '……'. De nguyen thi bang tieng Viet nhin lech han chu.
FULLWIDTH = {u"（": "(", u"）": ")", u"，": ",", u"：": ":",
             u"；": ";", u"～": "~", u"／": "/", u"－": "-",
             u"……": "...", u"？": "?", u"！": "!",
             u"、": ", ", u"。": ". ", u"％": "%",
             u"＜": "<", u"＞": ">", u"＝": "=", u"＋": "+"}

# The nam GIUA chu: trinh soan thao cua ESPE cat vun chu thanh nhieu <span>
# ("<span>Lampshad</span><span>e</span>") nen neu doi the thanh dau cach thi
# chu bi tach ra ("Lampshad e"). Bo han cac the nay truoc khi doi.
INLINE = re.compile(r"(?i)</?(span|strong|b|i|u|em|font|a|sub|sup|small)\b[^>]*>")


def plain(s):
    for a, b in FULLWIDTH.items():
        s = s.replace(a, b)
    return s


def txt(s):
    s = INLINE.sub("", s or "")
    return plain(H.unescape(
        re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s)))).strip()


def cell_txt(s):
    """O bang: giu ngat dong (<br>, </p>) thanh '; ' vi nhieu o chua 2-3 y."""
    s = INLINE.sub("", s or "")
    s = re.sub(r"(?i)<br\s*/?>|</p>|</div>", "\n", s)
    s = plain(H.unescape(re.sub(r"<[^>]+>", " ", s)))
    parts = [re.sub(r"[ \t\xa0]+", " ", p).strip(" ;,") for p in s.split("\n")]
    return "; ".join(p for p in parts if p)


# ---------------------------------------------------------------------------
# 1. Boc mot trang dong SP
# ---------------------------------------------------------------------------
H1 = re.compile(r"<h1[^>]*>(.*?)</h1>", re.S)
INTRO = re.compile(r'<div class="pr21b">(.*?)INQUIRE NOW', re.S)
MODEL_OF = re.compile(r"Model\s*[:：]\s*([^:]{1,40}?Series|[A-Z0-9][\w./-]{1,24})",
                      re.I)
MAIN_IMG = re.compile(r'class="chanpin1">\s*<img src="([^"]+)"')
SEC_SPLIT = '<div class="pr22lie">'
SEC_TITLE = re.compile(r'class="pr22biao[^"]*">(.*?)(?:<em|<)')
TABLE = re.compile(r"<table.*?</table>", re.S)
ROW = re.compile(r"<tr.*?</tr>", re.S)
CELL = re.compile(r"<(t[dh])([^>]*)>(.*?)</\1>", re.S)
SPAN = re.compile(r'(rowspan|colspan)\s*=\s*"?(\d+)', re.I)

# khoi la bang thong so ky thuat
RX_SPEC = re.compile(r"technical\s*parameter|specification|product\s*detail"
                     r"|product\s*parameter", re.I)
# o 'Download' cua cot file 3D - khong phai thong so
RX_DROP_COL = re.compile(r"^(3d\s*file|download|drawing|file)$", re.I)


def sections(html):
    """[(tieu de khoi, html than khoi)] theo thu tu tren trang."""
    out = []
    for blk in html.split(SEC_SPLIT)[1:]:
        blk = blk.split('<div class="pr22n3')[0]
        m = SEC_TITLE.search(blk)
        out.append((txt(m.group(1)) if m else "", blk))
    return out


def grid(tbl):
    """<table> -> luoi chu nhat, da tra lai cac o bi gop va bo cot file tai ve.

    Phai giai rowspan/colspan: bang chon model cua ESPE gop o rat nhieu, khong
    tra lai thi cac dong duoi bi thieu o va lech cot - vd dong
    'Unique coding | SLR11-001N' mat sach 3 cot dau (Locking mode /
    Configuration / Output) nen khong doc ra duoc ma SP thuoc loai nao.
    """
    out, hang = [], {}        # hang: {cot: [so dong con treo, gia tri]}
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
            v = cell_txt(body)
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
    drop = {i for i, c in enumerate(out[0]) if RX_DROP_COL.match(c)}
    if drop:
        out = [[c for i, c in enumerate(r) if i not in drop] for r in out]
    return [r for r in out if any(r)]


def _paired(rows):
    """Bang thong so xep 2 cap mot dong: 'Diameter|50mm|Life-span|50000 hours'.

    Phan biet voi bang ma tran 4 cot (Door Status|Lock Status|OSSD|AUX) bang
    cach xem cot 0 va cot 2: o bang xep cap thi moi o la mot TEN thong so khac
    nhau; o bang ma tran thi cot 0 lap lai ('Close', 'Close', 'Open').
    """
    full = [r for r in rows if len(r) >= 4 and r[0] and r[1] and r[2] and r[3]]
    if len(full) < 2 or len(full) < len(rows) - 1:
        return False
    c0 = [r[0] for r in full]
    c2 = [r[2] for r in full]
    return len(set(c0)) == len(c0) and len(set(c2)) == len(c2)


def spec_pairs(rows):
    """Bang thong so -> [[ten thong so, gia tri]].

    Dong CHI CO MOT o la tieu de nhom ("Electric parameter", "Environment"),
    khong phai thong so -> danh dau '#' o cot dau de bai dang in thanh dong
    phan nhom.
    """
    def key(s):
        """Ten thong so khong bao gio gom nhieu gia tri, nen dau ';' ma cell_txt
        chen vao cho <br> o day chi la cho xuong dong ("Lampshade; Material")
        -> doi lai thanh dau cach."""
        return re.sub(r"\s*;\s*", " ", s).strip()

    pair = _paired(rows)
    out = []
    for r in rows:
        # O trai ra nhieu cot (colspan) bi grid() nhan ban thanh nhieu o giong
        # nhau -> gop lai, neu khong thi ten thong so bi in hai lan
        # ("Standard | Standard | EN ISO 13849-1").
        flat = []
        for c in r:
            if not flat or c != flat[-1]:
                flat.append(c)
        cells = [c for c in flat if c]
        if len(cells) == 1:
            out.append(["#", key(cells[0])])
        elif pair and len(r) >= 4 and r[0] and r[2]:
            out.append([key(r[0]), r[1]])
            out.append([key(r[2]), " | ".join(x for x in r[3:] if x)])
        elif len(cells) == 3:
            # ten thong so hai cap ("Response time" + "Lock-unlock")
            out.append(["%s - %s" % (key(cells[0]), key(cells[1])), cells[2]])
        elif len(cells) >= 2:
            out.append([key(cells[0]), " | ".join(cells[1:])])
    return out


def model_tables(secs):
    """Cac bang CO cot 'Model' -> bang ma SP (giu nguyen dang luoi).

    Bang nay nam rai rac: 'Option for Brackets', 'Model Selection',
    'Combination & Model Selection' -> quet het cac khoi chu khong loc theo ten.
    """
    out = []
    for name, blk in secs:
        for tbl in TABLE.findall(blk):
            rows = grid(tbl)
            if len(rows) < 2 or len(rows[0]) < 2:
                continue
            if not any(re.fullmatch(r"model(s)?", c, re.I) for c in rows[0]):
                continue
            out.append({"sec": name, "head": rows[0], "rows": rows[1:]})
    return out


def spec_table(secs):
    """Bang thong so: uu tien khoi co tieu de 'Technical Parameters'/..., khong
    co thi lay bang 2 cot dai nhat tren trang."""
    def ok(rows):
        if len(rows) < 3:
            return False
        w = max(len(r) for r in rows)
        if any(re.fullmatch(r"model(s)?", c, re.I) for c in rows[0]):
            return False
        return w <= 3 or (w == 4 and _paired(rows))

    for name, blk in secs:
        if not RX_SPEC.search(name):
            continue
        for tbl in TABLE.findall(blk):
            rows = grid(tbl)
            if ok(rows):
                return spec_pairs(rows)
    best = []
    for _name, blk in secs:
        for tbl in TABLE.findall(blk):
            rows = grid(tbl)
            if ok(rows) and len(rows) > len(best):
                best = rows
    return spec_pairs(best) if best else []


RX_NUMBERED = re.compile(r"^\s*(\d{1,2})\s*[.、)）]\s*(.+)$")


def features(html):
    """Product Feature: giu nguyen cach danh so 1. 2. 3. cua hang."""
    m = INTRO.search(html)
    if not m:
        return []
    body = re.sub(r"(?i)<br\s*/?>", "</p>", m.group(1))
    lines = [txt(p) for p in re.split(r"</p>|</li>|</div>", body)]
    out, started = [], False
    for t in lines:
        if not t:
            continue
        if re.match(r"(?i)^(product\s*)?features?\s*[:：]?$", t):
            started = True
            continue
        if RX_NUMBERED.match(t):
            started = True
            out.append(t)
        elif started and len(t) > 25 and not re.match(r"(?i)^(brand|model)\s*[:：]", t):
            out.append(t)
    return out[:8]


def series_of(html):
    """'Model: ESV Series'. Nhan va gia tri nam o HAI the khac nhau nen phai
    xoa the truoc roi moi do, khong do thang tren HTML."""
    m = INTRO.search(html)
    if not m:
        return ""
    mm = MODEL_OF.search(txt(m.group(1)))
    return mm.group(1).strip() if mm else ""


def fetch_line(url):
    from shared.fetcher import try_fetch
    s = try_fetch(urllib.parse.quote(url, safe=":/?&=#%"), retries=2)
    if not s:
        return None
    if isinstance(s, bytes):
        s = s.decode("utf-8", "replace")
    secs = sections(s)
    m = H1.search(s)
    im = MAIN_IMG.search(s)
    return {"h1": txt(m.group(1)) if m else "",
            "series": series_of(s),
            "feats": features(s),
            "specs": spec_table(secs),
            "models": model_tables(secs),
            "img": (SITE + "/" + im.group(1).lstrip("/")) if im else ""}


# ---------------------------------------------------------------------------
# 2. Mo ta dai / ngan
# ---------------------------------------------------------------------------
def code_col(t):
    """Chi so cot 'Model' trong mot bang chon model."""
    for i, c in enumerate(t["head"]):
        if re.fullmatch(r"model(s)?", c, re.I):
            return i
    return -1


def model_cells(models):
    """Cac o CAN DICH trong bang chon model: bo cot ma SP (dich se pha ma)."""
    out = set()
    for t in models:
        ci = code_col(t)
        for r in t["rows"]:
            for i, c in enumerate(r):
                if c and i != ci:
                    out.add(c)
    return out


def model_codes(models):
    """Cac ma SP trong cot 'Model' - dung cho the Tag."""
    out = []
    for t in models:
        i = code_col(t)
        if i < 0:
            continue
        for r in t["rows"]:
            if i < len(r) and r[i]:
                # web ho go lech ("E SV 2010") -> bo khoang trang trong ma
                out.append(re.sub(r"\s+", "", r[i]))
    seen, uniq = set(), []
    for c in out:
        if c not in seen:
            seen.add(c)
            uniq.append(c)
    return uniq


def long_desc(specs, line_en=""):
    """Mo ta dai: thong so cua dong SP, bo cai trung ten dong."""
    low = (line_en or "").lower()
    out = []
    for k, v in specs:
        if k == "#" or not v:
            continue
        if v.lower() in low:
            continue
        out.append("%s: %s" % (k, v))
    return out


# ---------------------------------------------------------------------------
# 3. Ten tieng Viet + anh
# ---------------------------------------------------------------------------
def is_code_name(s):
    """Ten dong SP la ma/ten rieng (EB13, WLC) hay cum tu thuong?"""
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
        json.dump({"brand": "15 ESPE Technology", "articles": arts},
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
            print("  cao %3d/%d %-40s KHONG TAI DUOC" % (i, len(todo), ln["name"][:40]))
            continue
        title = d["h1"] or ln["name"]
        d["img_file"] = grab_image(d["img"], title, slugify, used)
        d["title"] = title
        raw[ln["url"]] = d

        long_en = long_desc(d["specs"], title)
        bag["key"].add(title)
        bag["key"].update(k for k, _v in d["specs"] if k and k != "#")
        bag["key"].update(v for k, v in d["specs"] if k == "#")
        bag["key"].update(c for t in d["models"] for c in t["head"])
        bag["value"].update(long_en)
        bag["value"].update(v for k, v in d["specs"] if k != "#" and v)
        bag["value"].update(model_cells(d["models"]))
        bag["prose"].update(d["feats"] or long_en[:MAX_HL])
        print("  cao %3d/%d %-40s ts=%-3d bang-model=%-2d feat=%-2d %s"
              % (i, len(todo), title[:40], len(d["specs"]), len(d["models"]),
                 len(d["feats"]), "img" if d["img_file"] else "---"))
        sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, specs, models = d["title"], d["specs"], d["models"]
        long_en = long_desc(specs, title)
        short_en = d["feats"] or long_en[:MAX_HL]

        keys = sorted({k for k, _v in specs if k and k != "#"}
                      | {v for k, v in specs if k == "#"}
                      | {c for t in models for c in t["head"]})
        vals = sorted({v for k, v in specs if k != "#" and v}
                      | model_cells(models))
        tr_k = dict(zip(keys, translate_many(keys, "key"))) if keys else {}
        tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        # giu nguyen quy uoc "#" o cot dau = dong tieu de nhom
        specs_vi = [["#", tr_k.get(v, v)] if k == "#"
                    else [tr_k.get(k, k), tr_v.get(v, v)] for k, v in specs]
        models_vi = []
        for t in models:
            ci = code_col(t)
            models_vi.append({
                "sec": t["sec"],
                "head": [tr_k.get(c, c) for c in t["head"]],
                # cot ma SP de nguyen, cac cot con lai dich
                "rows": [[c if i == ci else tr_v.get(c, c)
                          for i, c in enumerate(r)] for r in t["rows"]]})

        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title, series=d["series"],
            img_file=d["img_file"], codes=model_codes(models),
            specs_en=specs, specs_vi=specs_vi,
            models_en=models, models_vi=models_vi,
            long_en=long_en, long_vi=translate_many(long_en, "value") if long_en else [],
            short_en=short_en,
            short_vi=translate_many(short_en, "prose") if short_en else [],
            name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-40s ts=%-3d ma=%-4d %s"
              % (i, len(todo), title[:40], len(specs),
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
