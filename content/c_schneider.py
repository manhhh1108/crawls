# -*- coding: utf-8 -*-
"""Schneider Electric (hang 11) - crawl noi dung bai dang cho tung dong SP.

    python c_schneider.py            # chay tiep nhung dong chua co
    python c_schneider.py 5          # chi 5 dong dau (xem thu)
    python c_schneider.py --force    # lam lai tu dau

Can chay truoc: python cat_map.py schneider

se.com KHONG chan curl (co Akamai nhung khong bat request thuong), chi co dieu
trang dong SP la mot ung dung Javascript: danh sach ma SP KHONG nam trong HTML
ma duoc goi them sau. Mo tab Products bang Chrome that va ghi lai request da
tim ra hai API JSON cong khai, khong can khoa:

    /ranges/<id>/products?brand=se&country-code=ca&language-code=en&Nrpp=1000
        -> {"totalNumRecs":132, "productIds":[...]}
    /products-card/main?brand=se&country-code=ca&language-code=en&ids=A,B,C
        -> {"products":[{commercialReference, description, pictureUrl, ...}]}

Phan trang kieu Endeca: No = bo qua bao nhieu ma, Nrpp = so ma moi lan. Cac ten
quen thuoc (page, offset, pageSize...) deu bi bo qua - tung thu het. Nrpp=1000
lay duoc ca danh sach trong mot lan goi.

/products-card/main chi nhan duoc khoang 80 ma mot lan (URL qua dai thi tra ve
rong) -> goi theo lo 40.

Phan Overview thi nguoc lai: nam san trong HTML trang, nhung trong mot khoi
Javascript cua builder.io chu khong phai trong the HTML:

    main:{id:"61769",name:"...",description:"...",longDescription:"...",
          image:{url:"https://download.se.com/files?p_Doc_Ref=..."}}
    presentation:{presentation1:"<p>Remote Mount Controllers..."}

Theo 'Template nội dung Schneider.docx':
    Mo ta dai  = "Lấy thông tin ở phần Feature trong Overview"  -> presentation1
                 (khong co thi longDescription, roi den description)
    Mo ta ngan = mau DESCDESC (template ghi san cau do)
    Bang ma SP = Model | Ma dat hang (Part No.)

LUU Y DUNG LUONG: moi trang dong SP nang ~1,5 MB, 759 trang -> cache HTTP phong
them vai tram MB. Dung xoa cache cua hang nay roi chay lai neu khong can.

Ket qua: out/11 Schneider/content.json -> build_schneider_excel.py dung tiep.
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

BRAND = "Schneider Electric"
SITE = "https://www.se.com"
Q = "brand=se&country-code=ca&language-code=en"
OUT = os.path.join(HERE, "out", "11 Schneider")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung Schneider")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

BATCH = 40              # so ma moi lan goi /products-card/main
MAX_ITEMS = 400
# Tran so dong mo ta dai. De 12 thi khoi Features cua TeSys LR9 (25 dong, ba
# dong SP LR9D01-110S / LR9D / LR9F) bi cat mat hon mot nua - dung loi "lay
# chua het thong tin" khach bao (review 04/11).
MAX_HL = 60

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

# Chuoi Javascript: "..." co the chua \" va <
JS_STR = r'((?:[^"\\]|\\.)*)'


def js_val(blob, key):
    """Doc mot truong chuoi trong khoi Javascript cua trang."""
    m = re.search(key + r':"' + JS_STR + '"', blob or "")
    if not m:
        return ""
    try:
        return json.loads('"' + m.group(1) + '"')
    except Exception:
        return m.group(1)


def txt(s):
    return re.sub(r"\s+", " ",
                  H.unescape(re.sub(r"<[^>]+>", " ", s or ""))).strip()


def plain(s):
    """Bo the HTML KHONG chen khoang trang - dung cho ten dong SP.

    Ho boc ten bang the to mau: `Power<span class="accent">PacT</span>` va
    `<span class="highlight-naming">...</span>`. Thay the bang khoang trang
    (nhu txt) thi ra "Power PacT" - sai ten. Phai noi lien.
    """
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s or ""))).strip()


# It nhat hai ky tu chu/so thi moi coi la mot dong noi dung
CHU = re.compile(r"[0-9A-Za-z\u00c0-\u024f]")


def prose_of(html_frag):
    """HTML cua phan Overview -> tung dong mot.

    Phai cat o CA THE MO: phan Overview cua ho viet kieu cu, the <LI> khong
    dong (`<UL><LI>a <LI>b <LI>c</LI></UL>`) - chi cat o </li> thi ba y dinh
    lien thanh mot dong dai.
    """
    b = re.sub(r"(?i)<iframe\b.*?</iframe>", " ", html_frag or "", flags=re.S)
    b = re.sub(r"(?i)<br\s*/?>", "</p>", b)
    out = []
    for p in re.split(r"(?i)</?(?:p|li|div|h[1-6]|tr)\b[^>]*>", b):
        t = txt(p)
        # Nguong cu la 12 KY TU, va bo trung tren CA khoi. Hai luat do cung
        # cat nham noi dung that: tieu de dong SP ("LR9D", "LR9F") chi 4 ky
        # tu, con dong "Three-phase and single-phase applications" lap lai o
        # ca ba dong SP la co that chu khong phai markup trung. Nay chi doi
        # dong co it nhat hai ky tu chu/so, va chi bo dong trung LIEN TIEP.
        if len(CHU.findall(t)) < 2:
            continue
        if out and t == out[-1]:
            continue
        out.append(t)
    return out[:MAX_HL]


TAG_RE = re.compile(r'tag:"([a-z0-9][a-z0-9-]{3,60})"')
HEAD_TAG = "heading-section-tittle-h1-h2-h3"
TITLE_RE = re.compile(r'title:"((?:[^"\\]|\\.)*)"')
# Hai truong chua van ban cua khoi Features:
#   content-text-with-left-title   -> rightColumn
#   content-accordion-container    -> itemDescription (moi panel mot cai)
VAN_RE = re.compile(r'(?:rightColumn|itemDescription):"((?:[^"\\]|\\.)*)"')


def _giai(v):
    """Chuoi Javascript -> chu that."""
    try:
        return json.loads('"' + v + '"')
    except Exception:
        return v


def khoi_features(s):
    """Van ban khoi Features, doc theo CAY COMPONENT cua trang.

    Trang se.com la mot chuoi component `{name:...,tag:"...",options:{...}}`.
    Tieu de moi muc la component `heading-section-tittle-h1-h2-h3`, nen khoi
    Features = tat ca component nam giua tieu de "Features" va tieu de KE
    TIEP. Cach nay thay cho `overview_of()` - doc theo doan van roi doan,
    hong hai kieu (review 04/11):

      * `content-accordion-container` co NHIEU panel, moi panel mot
        `itemDescription`. TeSys island co 6 panel ma cach cu chi lay panel
        dau -> mat 5/6 y khach dem duoc tren trang.
      * doan dai nam NGAY SAU khoi Features (van ban SEO cuoi trang) de bi vo
        nham vao, vi theo doan van thi khong biet khoi ket thuc o dau.

    Trang kieu cu khong co component nao nhu vay -> tra ve rong, nguoi goi
    quay lai chuoi uu tien cu.
    """
    tags = [(m.start(), m.group(1)) for m in TAG_RE.finditer(s)]
    heads = []
    for p, t in tags:
        if t != HEAD_TAG:
            continue
        m = TITLE_RE.search(s, p, p + 1200)
        heads.append((p, txt(_giai(m.group(1))) if m else ""))
    k = next((i for i, (_p, h) in enumerate(heads)
              if h.lower().startswith("features")), None)
    if k is None:
        return []
    dau = heads[k][0]
    cuoi = heads[k + 1][0] if k + 1 < len(heads) else len(s)
    out, thay = [], set()
    for m in VAN_RE.finditer(s, dau, cuoi):
        v = _giai(m.group(1))
        if "<p>" not in v and "<li" not in v.lower():
            continue
        ls = prose_of(v)
        # Trang chua hai ban cua cung cay component; bo ban trung theo noi dung.
        key = " ".join(ls)[:160]
        if not ls or key in thay:
            continue
        thay.add(key)
        out += ls
    return out[:MAX_HL]


# Moc mot doan van trong trang; <p> co the con nguyen hoac da escape \u003C
MOC_P = re.compile(r"\\u003Cp>|<p>")
# Tieu de mo dau vung Overview - lay tu day tro di, de khong vo phai doan
# "dang ky nhan tin" cua khung trang (nam truoc, cung la mot the <p> dai).
MO_DAU = re.compile(r"^(discover\b|features\s*$|benefits of\b|facts about\b)",
                    re.I)
TIEU_DE = re.compile(r"^(features|benefits\b|discover\b|facts about\b|"
                     r"delivering value\b|additional resources\b|"
                     r"what'?s new\b)", re.I)
# Nut bam lan trong khoi van, khong phai noi dung
NUT = re.compile(r"^(view|download|learn more|read more|contact|explore|"
                 r"find out|see )", re.I)


def _dau_chuoi(s, i):
    """Lui tu i ve dau nhay MO chuoi, bo qua nhay bi escape."""
    j = i
    while True:
        j = s.rfind('"', 0, j)
        if j < 0:
            return -1
        k, n = j - 1, 0
        while k >= 0 and s[k] == "\\":
            n += 1
            k -= 1
        if n % 2 == 0:
            return j


def _cuoi_chuoi(s, j):
    e, n = j + 1, len(s)
    while e < n:
        c = s[e]
        if c == "\\":
            e += 2
            continue
        if c == '"':
            return e
        e += 1
    return -1


def doan_p(s):
    """[(vi tri, html)] moi chuoi Javascript co chua mot the <p>.

    Trang chua noi dung Overview hai lan: mot ban trong cay builder.io va mot
    ban trong mang du lieu cua SvelteKit. Ban builder.io bi rut gon - chuoi
    nao dung hai lan bi thay bang bien (`description:v`) nen doc mot minh no
    la mat van. Ban SvelteKit thi du chu nhung xep theo VI TRI, khong co ten
    truong. Cach chac an nhat: quet chuoi Javascript o CA HAI ban roi bo trung
    theo noi dung.

    Phai doc dung kieu chuoi Javascript (\" va \u003C), dung thay the tho
    `s.replace('\\"', '"')` - lam vay thi dau nhay BEN TRONG chuoi thanh dau
    nhay that, khong con biet chuoi het o dau.
    """
    out, thay, qua = [], set(), -1
    for m in MOC_P.finditer(s):
        if m.start() < qua:
            continue
        j = _dau_chuoi(s, m.start())
        if j < 0:
            continue
        e = _cuoi_chuoi(s, j)
        if e < 0 or e - j > 20000:
            continue
        qua = e
        try:
            v = json.loads(s[j:e + 1])
        except Exception:
            continue
        if "<p>" not in v:
            continue
        k = txt(v)[:120]
        if not k or k in thay:
            continue
        thay.add(k)
        out.append((j, v))
    return out


def overview_of(s):
    """Van ban Overview theo dung thu tu khach yeu cau trong review:

    khoi Features > phan Overview khac > (nguoi goi tu lui ve blurb canh anh).
    """
    ds = doan_p(s)
    i = next((n for n, (_p, h) in enumerate(ds) if MO_DAU.match(txt(h))), None)
    if i is None:
        return []
    ds = ds[i:]
    for dau in ("features", "discover"):
        # Than khoi nam ngay sau tieu de cua no. "Discover <ten dong SP>" la
        # doan gioi thieu chinh cua tab Overview; cac khoi con lai
        # ("Delivering value...", "Benefits of...") dung o cuoi cung vi co
        # trang xep chung TRUOC doan gioi thieu.
        for n, (_p, h) in enumerate(ds):
            if txt(h).lower().startswith(dau) and n + 1 < len(ds):
                t = ds[n + 1][1]
                if len(txt(t)) >= 80:
                    return prose_of(t)
    for _p, h in ds:
        t = txt(h)
        if len(t) >= 80 and not TIEU_DE.match(t) and not NUT.match(t):
            return prose_of(h)
    return []


def main_blob(s, rid):
    """Khoi main:{...} cua dong SP - cat ra de khong doc nham truong trung ten.

    Trang co hang chuc cho dung key "description"; chi khoi nay la cua san pham.
    """
    i = s.find('main:{id:"%s"' % rid)
    if i < 0:
        i = s.find("main:{id:")
    return s[i:i + 4000] if i >= 0 else ""


OG_IMG = re.compile(r'<meta property="og:image" content="([^"]+)"', re.I)


def fetch_line(url, rid):
    from shared.fetcher import try_fetch
    s = try_fetch(url, retries=2)
    if not s:
        return None
    blob = main_blob(s, rid)
    name = plain(js_val(blob, "name"))
    desc = js_val(blob, "description")
    ldesc = js_val(blob, "longDescription")
    pres = js_val(s, "presentation1")
    img = ""
    m = re.search(r'image:\{url:"' + JS_STR + '"', blob)
    if m:
        try:
            img = json.loads('"' + m.group(1) + '"')
        except Exception:
            img = m.group(1)

    # Review khach: uu tien Features trong Overview > phan Overview khac >
    # mo ta ngan canh anh san pham. presentation1 la kieu trang cu, van giu.
    # khoi_features() doc theo cay component (chinh xac, lay du ca accordion);
    # overview_of() la cach cu, van giu cho trang khong co cay component do.
    prose = (khoi_features(s) or overview_of(s) or prose_of(pres)
             or prose_of(ldesc) or prose_of(desc))

    items, total = [], 0
    t = try_fetch("%s/ranges/%s/products?%s&Nrpp=1000" % (SITE, rid, Q),
                  retries=2)
    if t:
        try:
            d = json.loads(t)
            total = d.get("totalNumRecs") or 0
            ids = (d.get("productIds") or [])[:MAX_ITEMS]
        except Exception:
            ids = []
        for k in range(0, len(ids), BATCH):
            lo = ids[k:k + BATCH]
            r = try_fetch("%s/products-card/main?%s&ids=%s"
                          % (SITE, Q, ",".join(lo)), retries=2)
            if not r:
                continue
            try:
                for p in json.loads(r).get("products") or []:
                    items.append({"art": p.get("commercialReference") or "",
                                  "name": plain(p.get("description") or ""),
                                  "img": p.get("pictureUrl") or "",
                                  "url": p.get("pdpPublicUrl") or ""})
            except Exception:
                continue
    if not img:
        img = next((i["img"] for i in items if i["img"]), "")
    if not img:
        # Cay builder.io bi rut gon: nhieu trang ghi `image:{url:a,...}` - `a`
        # la BIEN, khong phai chuoi, nen regex o tren khong thay gi. The
        # og:image o dau trang luon ghi day du URL anh do (82/84 bai truoc day
        # phai dung logo thuc ra co anh - review 04/11).
        m = OG_IMG.search(s)
        if m:
            img = H.unescape(m.group(1))
    return {"h1": name, "summary": plain(desc), "prose": prose,
            "items": items, "total": total or len(items), "img": img}


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


def anh_to(url):
    """URL anh, ban to nhat truoc.

    So trong `rendition_369_jpg` chinh la so pixel. Hang chi dang san vai co;
    369 va 113 la hai co hay gap trong du lieu san pham, nhung **1500 cung co
    that** va net hon han - anh 113 px phong len khung 800 thi vo nhoe, dung
    la loi "anh mo va nho" khach bao. Co nao khong co thi may chu tra ve anh
    mac dinh ~1,3 KB, nguoi goi loc theo dung luong.
    """
    if "rendition_" not in (url or ""):
        return [url] if url else []
    return [re.sub(r"rendition_\d+_\w+", "rendition_1500_jpg", url),
            re.sub(r"rendition_\d+_\w+", "rendition_369_jpg", url),
            url]


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
    src = os.path.join(SRC_IMG, base + ".img")
    try:
        if not (os.path.exists(src) and os.path.getsize(src) > 1500):
            data = b""
            for v in anh_to(url):
                try:
                    req = urllib.request.Request(
                        v, headers={"User-Agent": UA,
                                    "Referer": SITE + "/ca/en/"})
                    b = urllib.request.urlopen(req, timeout=90).read()
                except Exception:
                    continue
                if len(b) > len(data):
                    data = b
                if len(data) > 1500:
                    break
            # Anh mac dinh cua ho (DefaultProductImage.png) chi vai tram byte -
            # dung lam anh bai thi ra o trang tron.
            if len(data) < 1500:
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
        json.dump({"brand": "11 Schneider", "articles": arts}, f,
                  ensure_ascii=False)


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

    # Cao va dung bai theo tung lo 50: 759 dong SP, gom het roi moi dich thi
    # mot lan dut la mat ca buoi cao.
    LOT = 50
    for lo0 in range(0, len(todo), LOT):
        lot = todo[lo0:lo0 + LOT]
        import prewarm
        raw, bag = {}, {"key": set(), "value": set(), "prose": set()}
        for i, ln in enumerate(lot, lo0 + 1):
            try:
                d = fetch_line(ln["url"], ln["code"])
            except Exception as e:
                d = None
                print("     LOI", repr(e)[:70])
            if not d:
                print("  cao %3d/%d %-38s KHONG TAI DUOC"
                      % (i, len(todo), ln["name"][:38]))
                continue
            title = d["h1"] or ln["name"]
            d["img_file"] = grab_image(d["img"], title, slugify, used)
            d["title"] = title
            raw[ln["url"]] = d

            bag["key"].add(title)
            bag["value"].update(it["name"] for it in d["items"] if it["name"])
            bag["prose"].update(d["prose"])
            print("  cao %3d/%d %-38s ma=%-4d/%-4d van=%-2d %s"
                  % (i, len(todo), title[:38], len(d["items"]), d["total"],
                     len(d["prose"]), "img" if d["img_file"] else "---"))
            sys.stdout.flush()
        prewarm.run(bag)

        for i, ln in enumerate(lot, lo0 + 1):
            d = raw.get(ln["url"])
            if not d:
                continue
            title, items = d["title"], d["items"]
            vals = sorted({it["name"] for it in items if it["name"]})
            tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
            for it in items:
                it["desc_vi"] = tr_v.get(it["name"], it["name"])

            done[ln["url"]] = dict(
                ln, url_src=ln["url"], title_en=title,
                series="", codes=[it["art"] for it in items if it["art"]],
                img_file=d["img_file"], total=d["total"], items=items,
                summary_en=d["summary"],
                long_en=d["prose"],
                long_vi=translate_many(d["prose"], "prose") if d["prose"] else [],
                short_en=[], short_vi=[],
                name_vi=name_vi_of(title, translate_many))
            print("  %3d/%d %-38s ma=%-4d dai=%-2d %s"
                  % (i, len(todo), title[:38], len(items), len(d["prose"]),
                     "img" if d["img_file"] else "---"))
        _flush(True)
        save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
