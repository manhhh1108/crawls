# -*- coding: utf-8 -*-
"""Supmea Automation (hang 12) - crawl noi dung bai dang cho tung san pham.

    python c_supmea.py            # chay tiep nhung dong chua co
    python c_supmea.py 3          # chi 3 dong dau (xem thu)
    python c_supmea.py --force    # lam lai tu dau

Can chay truoc: python cat_map.py supmea

en.supmea.com khong chan request thuong -> dung shared.fetcher (curl + cache).
Moi dong trong danh muc Ntrend la MOT san pham cu the, bo cuc trang rat deu:

    h1.title            ten san pham
    div.summary         doan gioi thieu ngan canh anh
    div.features        "Characteristics" + cac <p> gach dau dong
    div#product-content doan Introduction, roi bang Specification (2 cot doc)
    meta og:image       anh dai dien

Theo 'Template nội dung Supmea.docx':
    Mo ta dai   = phan Introduction
    Mo ta ngan  = phan Characteristics
    Thong so    = bang specifications

Trang hang KHONG co bang ma SP (template cung khong co o do) - ma san pham
nam ngay trong bang thong so o dong "Model" / "Type", lay ra lam the tag.

Ket qua: out/12 Supmea/content.json -> build_supmea_excel.py dung tiep.
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

BRAND = "Supmea"
SITE = "https://en.supmea.com"
OUT = os.path.join(HERE, "out", "12 Supmea")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung Supmea")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_HL = 8

# Ba dong trong danh muc Ntrend tro vao TRANG DANH MUC cua hang chu khong phai
# trang san pham. Da mo tung trang kiem: khong co the <div class="item-col
# product-col"> nao - tuc hang chua dang san pham nao vao ba muc nay (trang
# danh muc that, vd /magnetic-flowmeter, co 7 the nhu vay). De lai thi ra ba
# bai rong: khong mo ta, khong thong so, phai lay logo lam anh.
BO_TRANG_DANH_MUC = {
    "https://en.supmea.com/online-water-quality-analyzer",
    "https://en.supmea.com/tss-ss-sensor-and-meter",
    "https://en.supmea.com/digital-pressure-gauge",
}

# Trang do nguoi Trung Quoc soan, dau cau hay o dang rong.
FULLWIDTH = {u"（": "(", u"）": ")", u"，": ",", u"：": ":",
             u"；": ";", u"～": "~", u"／": "/", u"－": "-",
             u"＜": "<", u"＞": ">", u"％": "%", u"、": ", ",
             u"。": ". ", u"＝": "=", u"＋": "+", u"？": "?"}
INLINE = re.compile(r"(?i)</?(span|strong|b|i|u|em|font|a|sub|sup|small)\b[^>]*>")


def plain(s):
    for a, b in FULLWIDTH.items():
        s = s.replace(a, b)
    # O bang cua ho dong nao cung mo dau bang &nbsp; -> thanh \xa0, de nguyen
    # thi ten thong so nhin nhu bi thut vao.
    return re.sub(r"\s+", " ", s.replace(u"\xa0", " "))


def txt(s):
    s = INLINE.sub("", s or "")
    return plain(H.unescape(re.sub(r"<[^>]+>", " ", s))).strip()


def cell_txt(s):
    """O bang: giu ngat dong (<br>) thanh '; ' vi nhieu o chua 2-3 y."""
    s = INLINE.sub("", s or "")
    s = re.sub(r"(?i)<br\s*/?>|</p>|</div>", "\n", s)
    s = plain(H.unescape(re.sub(r"<[^>]+>", " ", s)))
    parts = [p.strip(" ;,") for p in s.split("\n")]
    return "; ".join(p for p in parts if p)


# ---------------------------------------------------------------------------
# 1. Boc mot trang san pham
# ---------------------------------------------------------------------------
H1 = re.compile(r"<h1[^>]*>(.*?)</h1>", re.S)
SUMMARY = re.compile(r'<div class="summary">(.*?)</div>', re.S)
FEATURES = re.compile(r'<div class="features">(.*?)</div>', re.S)
# Khoi noi dung chinh: chi lay den tab ke tiep ("Related Products"), neu khong
# thi vo luon phan san pham lien quan o duoi.
PCONTENT = re.compile(r'id="product-content"[^>]*>(.*?)<div class="tab-pane',
                      re.S)
OG_IMG = re.compile(r'<meta property="og:image" content="([^"]+)"')
TABLE = re.compile(r"<table.*?</table>", re.S)
ROW = re.compile(r"<tr.*?</tr>", re.S)
CELL = re.compile(r"<(t[dh])([^>]*)>(.*?)</\1>", re.S)
SPAN = re.compile(r'(rowspan|colspan)\s*=\s*"?(\d+)', re.I)

# Dong chi la nhan khoi tren trang, khong phai cau van. Phai khop CA dong thi
# moi coi la nhan: trang cua ho boc danh sach dac diem trong cung the <ul> voi
# nhan khoi, nhan nham thi mat han doan do.
LABEL_ONLY = re.compile(r"^(main\s+|key\s+|product\s+)?(introduction"
                        r"|specifications?|parameters?|features?"
                        r"|characteristics|applications?|details?"
                        r"|advantages?|technical\s+(data|parameters?))"
                        r"\s*[:：]?$", re.I)
# Khoi danh dau tieu de: co trang dung <h3>, co trang dung <ul><li><p>
LABEL_BLOCK = re.compile(r"(?is)<(ul|ol|h[1-6]).*?</>")
# O chua ma san pham trong bang thong so
CODE_LBL = re.compile(r"^\s*(product\s*)?(model|type)s?\s*$", re.I)


def grid(tbl):
    """<table> -> luoi chu nhat, da tra lai cac o bi gop (rowspan/colspan).

    Bang thong so cua ho gop o rat nhieu: "Nominal pressure" trai dai 2 dong
    (DN15-DN250 va DN300-DN1000). Khong tra lai thi dong thu hai mat ten thong
    so, doc ra chi con mot gia tri tro troi.
    """
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
    keep = [i for i in range(w) if any(r[i] for r in out)]
    out = [[r[i] for i in keep] for r in out]
    return [r for r in out if any(r)]


# Tu ngu chung o dong tieu de bang ("Item | Index | Value") - khong phai ten
# thong so nao ca.
HDR_WORD = re.compile(r"^(item|index|value|parameter|name|content|unit|remark"
                      r"|no\.?|specifications?|descriptions?|project)s?$", re.I)


def bo_dong_tieu_de(rows):
    """Bang 3 cot co khi mo dau bang mot dong tieu de chung, bo di cho khoi
    lot vao bang thong so thanh mot dong vo nghia."""
    if len(rows) > 1 and all(HDR_WORD.match(c) for c in flat_row(rows[0])):
        return rows[1:]
    return rows


def flat_row(r):
    """Bo cac o lap ngay canh nhau - do colspan sinh ra khi tra lai luoi."""
    out = []
    for c in r:
        if not out or c != out[-1]:
            out.append(c)
    return [c for c in out if c]


def spec_pairs(rows):
    """Bang thong so doc -> [[ten thong so, gia tri]], "#" la dong tieu de nhom.

    Hai dong lien nhau cung TEN thong so la mot o bi gop bang rowspan -> gop
    gia tri lai, neu khong thi bai dang in ten thong so hai lan.
    """
    out, nhom = [], None

    def them(k, v):
        if out and out[-1][0] == k and k != "#":
            out[-1][1] += "; " + v
        else:
            out.append([k, v])

    for r in rows:
        cells = flat_row(r)
        if len(cells) == 1:
            out.append(["#", cells[0]])
            nhom = cells[0]
        elif len(cells) == 2:
            them(cells[0], cells[1])
        elif len(cells) >= 3:
            # Cot trai la TEN NHOM trai dai nhieu dong (System / Measurement) -
            # dua len thanh dong tieu de nhom cho khoi in ra lap lai.
            if cells[0] != nhom:
                out.append(["#", cells[0]])
                nhom = cells[0]
            them(cells[1], " | ".join(cells[2:]))
    return out


def doc_table(rows):
    """Bang nay la bang thong so chu khong phai bang tra cuu nhieu cot.

    Phan lon la bang hai cot "ten | gia tri". Nhung 14 trang de them mot cot
    NHOM ben trai ("System | working power | (220+-22)VAC") - van la bang thong
    so, chi khac la ten thong so gom hai cap. Phan biet voi bang tra cuu 3 cot
    that (Degree no | Signal type | Range) bang dau vet cua o bi gop: bang
    thong so co cot nhom lap lai hoac co dong bi colspan keo ngang.
    """
    if len(rows) < 3:
        return False
    w = max(len(r) for r in rows)
    if w > 3:
        return False
    if w == 2:
        return True
    col0 = [r[0] for r in rows]
    return len(set(col0)) < len(col0) or any(len(flat_row(r)) < 3 for r in rows)


def sections(body):
    """[(nhan khoi da viet thuong, than khoi)] theo dung thu tu tren trang.

    Khong cat theo "van truoc bang, bang sau" duoc: nua so trang de bang
    Specification LEN TRUOC roi moi den Introduction (vd Coriolis mass
    flowmeter), co trang phan Introduction chi toan anh chup.
    """
    out, last, name = [], 0, ""
    for m in LABEL_BLOCK.finditer(body):
        t = txt(m.group(0))
        if not LABEL_ONLY.match(t):
            continue
        out.append((name, body[last:m.start()]))
        name, last = t.rstrip(u":：").lower(), m.end()
    out.append((name, body[last:]))
    return out


def prose_of(body):
    """Doan Introduction (mo ta dai theo template Supmea)."""
    secs = sections(body)
    phan = [b for n, b in secs if "introduction" in n]
    if not phan:
        # Trang khong ghi nhan "Introduction" -> lay van o cac khoi khong phai
        # bang thong so.
        phan = [b for n, b in secs
                if not re.match(r"(?i)^(specification|parameter|technical)", n)]
    out = []
    for b in phan:
        b = TABLE.sub(" ", b)          # bo bang, chi lay van
        b = re.sub(r"(?i)<br\s*/?>", "</p>", b)
        for p in re.split(r"(?i)</p>|</li>|</div>|</h[1-6]>", b):
            if "<img" in p:            # doan chi chua anh minh hoa
                continue
            t = txt(p)
            if not t or len(t) < 12 or LABEL_ONLY.match(t) or t in out:
                continue
            out.append(t)
    return out[:12]


def feats_of(html):
    """Characteristics: moi <p> la mot gach dau dong."""
    m = FEATURES.search(html)
    if not m:
        return []
    out = []
    body = re.sub(r"(?i)<br\s*/?>", "</p>", m.group(1))
    for p in re.split(r"(?i)</p>|</h[1-6]>", body):
        t = txt(p)
        if t and not LABEL_ONLY.match(t) and t not in out:
            out.append(t)
    return out[:MAX_HL]


def fetch_line(url):
    from shared.fetcher import try_fetch
    s = try_fetch(urllib.parse.quote(url, safe=":/?&=#%"), retries=2)
    if not s:
        return None
    if isinstance(s, bytes):
        s = s.decode("utf-8", "replace")
    m = PCONTENT.search(s)
    body = m.group(1) if m else ""

    # Mot trang co the co NHIEU bang thong so doc noi nhau (SUP-TRC 400: bang
    # chung + bang cua dau do) -> gop het vao mot bang thong so. Bang nhieu cot
    # (bang tra ma tin hieu) de rieng xuong khoi bang phu.
    specs, extra = [], []
    for tbl in TABLE.findall(body):
        rows = grid(tbl)
        if not rows:
            continue
        if doc_table(rows):
            specs.extend(spec_pairs(bo_dong_tieu_de(rows)))
        elif len(rows) >= 2 and len(rows[0]) >= 2:
            extra.append({"head": rows[0], "rows": rows[1:]})

    h1 = H1.search(s)
    sm = SUMMARY.search(s)
    im = OG_IMG.search(s)
    prose = prose_of(body)
    if not prose and sm:
        # Trang khong co khoi Introduction -> dung doan gioi thieu canh anh.
        prose = [t for t in (txt(sm.group(1)),) if len(t) >= 12]
    return {"h1": txt(h1.group(1)) if h1 else "",
            "summary": txt(sm.group(1)) if sm else "",
            "feats": feats_of(s),
            "prose": prose,
            "specs": specs,
            "extra": extra,
            "img": im.group(1) if im else ""}


# ---------------------------------------------------------------------------
# 2. Ma SP, mo ta dai
# ---------------------------------------------------------------------------
def spec_codes(specs):
    """Ma SP nam trong bang thong so: dong "Model" / "Type" - dung lam the tag."""
    out, seen = [], set()
    for k, v in specs:
        if k == "#" or not CODE_LBL.match(k):
            continue
        for c in re.split(r"[;,/]| or ", v):
            c = c.strip()
            # "Model | FMC240" thi lay, "Type | Electromagnetic flowmeter" la
            # ten loai chu khong phai ma -> chi nhan chuoi co chu HOA + so.
            if (c and len(c) <= 30 and re.search(r"\d", c)
                    and re.match(r"^[A-Z0-9][\w.\-/ ]*$", c) and c not in seen):
                seen.add(c)
                out.append(c)
    return out


def long_desc(specs, line_en=""):
    """Mo ta dai du phong: thong so cua san pham, bo cai trung ten bai."""
    low = (line_en or "").lower()
    out = []
    for k, v in specs:
        if k == "#" or not v or v.lower() in low:
            continue
        out.append("%s: %s" % (k, v))
    return out


def extra_cells(extra):
    """Cac o can dich trong bang phu."""
    out = set()
    for t in extra:
        out.update(c for c in t["head"] if c)
        for r in t["rows"]:
            out.update(c for c in r if c)
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
    # Anh cua ho la .webp tren CDN rieng (c70.hongcdn.com) - PIL doc duoc,
    # brandimg se xuat ra .jpg nhu cac hang khac.
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
        json.dump({"brand": "12 Supmea", "articles": arts}, f,
                  ensure_ascii=False)


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
        d["img_file"] = grab_image(d["img"], title, slugify, used)
        d["title"] = title
        raw[ln["url"]] = d

        bag["key"].add(title)
        bag["key"].update(k for k, _v in d["specs"] if k and k != "#")
        bag["key"].update(v for k, v in d["specs"] if k == "#")
        bag["value"].update(v for k, v in d["specs"] if k != "#" and v)
        bag["value"].update(extra_cells(d["extra"]))
        bag["prose"].update(d["prose"])
        bag["prose"].update(d["feats"])
        print("  cao %3d/%d %-40s ts=%-3d bang-phu=%-2d feat=%-2d van=%-2d %s"
              % (i, len(todo), title[:40], len(d["specs"]), len(d["extra"]),
                 len(d["feats"]), len(d["prose"]),
                 "img" if d["img_file"] else "---"))
        sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai (moi ban dich deu da nam trong cache) ----
    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, specs, extra = d["title"], d["specs"], d["extra"]
        keys = sorted({k for k, _v in specs if k and k != "#"}
                      | {v for k, v in specs if k == "#"})
        vals = sorted({v for k, v in specs if k != "#" and v}
                      | extra_cells(extra))
        tr_k = dict(zip(keys, translate_many(keys, "key"))) if keys else {}
        tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        specs_vi = [["#", tr_k.get(v, v)] if k == "#"
                    else [tr_k.get(k, k), tr_v.get(v, v)] for k, v in specs]
        extra_vi = [{"head": [tr_v.get(c, c) for c in t["head"]],
                     "rows": [[tr_v.get(c, c) for c in r] for r in t["rows"]]}
                    for t in extra]

        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title,
            series=(spec_codes(specs) or [""])[0],
            img_file=d["img_file"], codes=spec_codes(specs),
            summary_en=d["summary"],
            specs_en=specs, specs_vi=specs_vi,
            extra_en=extra, extra_vi=extra_vi,
            long_en=d["prose"],
            long_vi=translate_many(d["prose"], "prose") if d["prose"] else [],
            short_en=d["feats"],
            short_vi=translate_many(d["feats"], "prose") if d["feats"] else [],
            name_vi=name_vi_of(title, translate_many))
        print("  %3d/%d %-40s ts=%-3d ma=%-3d %s"
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
