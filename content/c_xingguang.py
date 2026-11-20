# -*- coding: utf-8 -*-
"""Guangdong Xingguang / Starshine Drive (hang 14) - crawl noi dung bai dang.

    python c_xingguang.py          # chay tiep nhung dong chua co
    python c_xingguang.py 3        # chi 3 dong dau (xem thu)
    python c_xingguang.py --force  # lam lai tu dau

Can chay truoc: python cat_map.py xingguang

en.xgcd.cn khong chan request thuong -> dung shared.fetcher (curl + cache).
Trang do mot nen tang website Trung Quoc sinh ra, lop the deu nhu nhau:

    <title>Ten san pham-Guangdong Xingguang Transmission Co., Ltd</title>
    <div class="small-img" data-url="portal-saas/.../cms/image/<uuid>.jpg">
    <div class='reset_style ...'>   khoi 1: anh so do tren dau trang
    <div class='reset_style ...'>   khoi 2: PHAN "Product Description"

Khoi Product Description khong co bang nao - ca trang chi la mot doan van
ngan cach bang <br>. Ba kieu dong trong do:

    "Output speed: 7.15rpm-155.56rpm"   thong so (ten: gia tri)
    "Characteristic:"                   nhan khoi
    "1. ..." / "● ..."                  gach dau dong dac diem

Theo 'Template nội dung Xingguang-Starshine.docx':
    Mo ta dai  = "Lấy full thông tin của mục Feature" - GIU DUNG THU TU tren
                 trang, con danh dau dong nao la gach dau dong (review 02/11)
    Mo ta ngan = "Lấy 5 gạch đầu dòng Feature nhỏ"
    O "Bang ma SP" trong template de TRONG - va dung la trang hang khong dang
    ma SP nao, nen bai khong co bang ma. Nhung dong "ten: gia tri" o dau phan
    mo ta la thong so that cua dong SP -> dua vao khoi "Thong so ky thuat".

Ket qua: out/14 Xingguang-Starshine/content.json -> build_xingguang_excel.py.
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

BRAND = "Starshine Drive"
SITE = "https://en.xgcd.cn"
IMG_HOST = "https://omo-oss-image.thefastimg.com/"
OUT = os.path.join(HERE, "out", "14 Xingguang-Starshine")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
XLSX_DIR = out_dir.brand_dir("Noi dung Xingguang-Starshine")
IMG_DIR = os.path.join(XLSX_DIR, "images")
for d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(d, exist_ok=True)

MAX_HL = 8

TITLE = re.compile(r"<title>(.*?)</title>", re.S)
RICH = re.compile(r"(?s)<div class='reset_style js-reset_style js-adapMobile '>"
                  r"(.*?)</div>")
MAIN_IMG = re.compile(r'class="small-img" data-url="([^"]+)"')
OG_IMG = re.compile(r'<meta property="og:image" content="([^"]+)"')
RICH_IMG = re.compile(r'<img[^>]+src="([^"]+)"', re.I)

# Dong chi la nhan khoi ("Feature:", "Key Benefits:", "SNFEG3000 Features:")
LABEL_ONLY = re.compile(r"^[\w\- ]{0,22}(feature|characteristic|key\s+benefit"
                        r"|advantage|specification|parameter|application)s?\s*"
                        r"[:：]?$", re.I)
# Gach dau dong: "1. ..." hoac "● ..." hoac "- ..."
BULLET = re.compile(r"^\s*(?:[0-9]{1,2}\s*[.、)]|[●•▲*\-–])\s*")
# "Output speed: 7.15rpm-155.56rpm" - ten thong so ngan, co dau hai cham
SPEC_LINE = re.compile(r"^([A-Z][A-Za-z0-9 /()\-]{1,26})\s*[:：]\s*(.*)$")


def plain(s):
    return re.sub(r"\s+", " ", (s or "").replace(u"\xa0", " ")).strip()


def txt(s):
    return plain(H.unescape(re.sub(r"<[^>]+>", " ", s or "")))


def lines_of(block):
    """Khoi Product Description -> danh sach dong theo dung thu tu tren trang."""
    b = re.sub(r"(?i)<br\s*/?>", "\n", block or "")
    out = []
    for ln in re.split(r"(?i)\n|</p>|</li>|</h[1-6]>|</div>", b):
        if "<img" in ln:                 # dong chi chua anh so do
            continue
        t = txt(ln)
        if t and t not in (".", "-"):
            out.append(t)
    return out


# Dong rac o CUOI khoi mo ta: vai trang bi dan lai danh sach tu khoa
# ("Line series reducer; reducer") - khong phai cau (khong co dau cham), cac
# ve ngan cach bang dau cham phay. Chi loai o dong CUOI de khong cham vao
# tieu de phu giua khoi ("Leak proof - sealed" cua trang SKC87).
RAC_CUOI = re.compile(r"^[^.!?]{1,80}$")


def _bo_rac(body):
    while body and not body[-1][0]:
        t = body[-1][1]
        if ";" in t and RAC_CUOI.match(t):
            body.pop()
            continue
        break
    return body


def doc_desc(block):
    """Khoi mo ta -> (thong so, dac diem, than bai).

    `than bai` la TOAN BO phan mo ta theo DUNG THU TU tren trang, moi phan tu
    la [co_phai_gach_dau_dong, chu]. Truoc day ham tra ve hai danh sach roi
    (van xuoi, dac diem) va nguoi goi noi `prose + feats`, nen doan ket
    "Therefore, it can be used as..." bi day len TRUOC cac gach dau dong -
    dung loi "noi dung bi nguoc" khach bao (review 02/11).

    Trang nao cung mo dau bang vai dong "ten: gia tri" (Output speed, Input
    power, Input) roi moi den nhan khoi va danh sach dac diem. Rieng "Input:"
    de gia tri TRONG roi liet ke tung kieu dau vao o cac dong sau - nen phai
    gom cac dong khong phai gach dau dong ngay sau no lai lam gia tri.
    """
    specs, feats, body = [], [], []
    cho = None                            # ten thong so dang cho gia tri
    trong_feat = False
    for t in lines_of(block):
        if LABEL_ONLY.match(t):
            trong_feat, cho = True, None
            continue
        if BULLET.match(t):
            trong_feat, cho = True, None
            v = BULLET.sub("", t).strip()
            if v and v not in feats:
                feats.append(v)
                body.append([True, v])
            continue
        m = SPEC_LINE.match(t)
        if m and not trong_feat:
            k, v = m.group(1).strip(), m.group(2).strip()
            if v:
                specs.append([k, v])
                cho = None
            else:
                specs.append([k, ""])
                cho = specs[-1]
            continue
        if cho is not None:
            cho[1] = (cho[1] + "; " + t).strip("; ")
            continue
        if len(t) >= 12 and all(t != x[1] for x in body):
            body.append([False, t])
    body = _bo_rac(body)
    if body and not any(b for b, _t in body):
        # Trang khong danh so "1. 2." thi cu doan thanh mot gach dau dong cho
        # de doc (review 02/11). Ca 16 trang hien tai deu co danh so nen nhanh
        # nay chua chay, nhung trang hang dang them sau se khong phai sua tay.
        body = [[True, t] for _b, t in body]
    return [s for s in specs if s[1]], feats, body


def fetch_line(url):
    from shared.fetcher import try_fetch
    s = try_fetch(url, retries=2)
    if not s:
        return None
    if isinstance(s, bytes):
        s = s.decode("utf-8", "replace")
    ti = TITLE.search(s)
    title = txt(ti.group(1)) if ti else ""
    title = re.sub(r"\s*-\s*Guangdong Xingguang.*$", "", title).strip()

    rich = RICH.findall(s)
    # Khoi mo ta la khoi NHIEU CHU nhat; khoi con lai chi chua anh.
    bi = max(range(len(rich)), key=lambda i: len(txt(rich[i]))) if rich else -1
    block = rich[bi] if bi >= 0 else ""
    specs, feats, body = doc_desc(block)

    # BANG THONG SO cua hang nay chi co o DANG ANH, nam o khoi khong co chu
    # (thuong la khoi dau trang). Khach xin tai ve kem bai (review 02/11).
    ts = ""
    for i, b in enumerate(rich):
        if i == bi:
            continue
        g = RICH_IMG.search(b)
        if g:
            ts = g.group(1)
            break

    im = MAIN_IMG.search(s)
    img = IMG_HOST + im.group(1) if im else ""
    if not img:
        og = OG_IMG.search(s)
        img = og.group(1) if og else ""
    return {"h1": title, "specs": specs, "feats": feats, "body": body,
            "img": img, "ts_img": ts}


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


# Ma dong SP nam ngay trong ten bai ("SVF Series Worm Gear Reducer" -> SVF).
SERIES = re.compile(r"\b([A-Z][A-Z0-9]{1,9}(?:\d{2,4})?)\b")
BO_TU = {"SERIES", "GEAR", "MOTOR", "UNIT", "UNITS", "SMART", "BEVEL",
         "HELICAL", "GEARED", "WORM", "SPEED", "DRIVE", "BOX", "FOR", "AND"}


def codes_of(title):
    out = []
    for c in SERIES.findall(title or ""):
        if c not in BO_TU and c not in out:
            out.append(c)
    return out


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


def grab_spec_img(url, base):
    """Tai ANH BANG THONG SO, GIU NGUYEN BAN - khong ghep khung NTREND.

    Hang chi dang bang nay o dang anh ~500 px (khach da xac nhan trong review).
    Ghep vao khung 800 thi bang bi thu nho lai, chu trong o kho doc; de nguyen
    ban thi dan thang vao than bai tren web duoc.
    """
    import urllib.request
    if not url:
        return ""
    if url.startswith("//"):
        url = "https:" + url
    elif url.startswith("/"):
        url = SITE + url
    ext = os.path.splitext(url.split("?")[0])[1][:5] or ".jpg"
    name = base + "-thong-so" + ext
    dst = os.path.join(IMG_DIR, name)
    try:
        if os.path.exists(dst) and os.path.getsize(dst) > 2000:
            return name
        req = urllib.request.Request(
            urllib.parse.quote(url, safe=":/?&=#%"),
            headers={"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64;"
                                   " x64) Chrome/131.0.0.0",
                     "Referer": SITE + "/"})
        data = urllib.request.urlopen(req, timeout=60).read()
        if len(data) < 2000:
            return ""
        with open(dst, "wb") as f:
            f.write(data)
        return name
    except Exception as e:
        print("     [anh bang loi]", repr(e)[:60])
        return ""


def save(ls, done):
    arts, seen = [], set()
    for l in ls:
        u = l.get("url")
        if u in done and u not in seen:
            seen.add(u)
            arts.append(done[u])
    with io.open(CONTENT, "w", encoding="utf-8") as f:
        json.dump({"brand": "14 Xingguang-Starshine", "articles": arts}, f,
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
    # Hai trang nam o hai nhanh danh muc khac nhau -> chi lam mot bai.
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
            print("  cao %2d/%d %-40s KHONG TAI DUOC"
                  % (i, len(todo), ln["name"][:40]))
            continue
        title = d["h1"] or ln["name"]
        d["img_file"] = grab_image(d["img"], title, slugify, used)
        d["ts_file"] = grab_spec_img(
            d["ts_img"], os.path.splitext(d["img_file"])[0]
            or (slugify(title)[:60].strip("-") or "sp"))
        d["title"] = title
        raw[ln["url"]] = d

        bag["key"].add(title)
        bag["key"].update(k for k, _v in d["specs"])
        bag["value"].update(v for _k, v in d["specs"] if v)
        bag["prose"].update(t for _b, t in d["body"])
        print("  cao %2d/%d %-40s ts=%-2d than-bai=%-2d %s %s"
              % (i, len(todo), title[:40], len(d["specs"]), len(d["body"]),
                 "img" if d["img_file"] else "---",
                 "bang" if d["ts_file"] else "----"))
        sys.stdout.flush()
    prewarm.run(bag)

    # ---- Pha 2: dung bai ----
    for i, ln in enumerate(todo, 1):
        d = raw.get(ln["url"])
        if not d:
            continue
        title, specs = d["title"], d["specs"]
        keys = sorted({k for k, _v in specs})
        vals = sorted({v for _k, v in specs if v})
        tr_k = dict(zip(keys, translate_many(keys, "key"))) if keys else {}
        tr_v = dict(zip(vals, translate_many(vals, "value"))) if vals else {}
        specs_vi = [[tr_k.get(k, k), tr_v.get(v, v)] for k, v in specs]
        # Mo ta dai = ca phan Feature theo DUNG THU TU tren trang; long_bul
        # ghi dong nao la gach dau dong de bai in ra "- ...".
        long_en = [t for _b, t in d["body"]]
        long_bul = [bool(b) for b, _t in d["body"]]
        done[ln["url"]] = dict(
            ln, url_src=ln["url"], title_en=title,
            series=(codes_of(title) or [""])[0], codes=codes_of(title),
            img_file=d["img_file"], spec_img_file=d["ts_file"],
            spec_img_url=d["ts_img"],
            specs_en=specs, specs_vi=specs_vi,
            long_en=long_en, long_bul=long_bul,
            long_vi=translate_many(long_en, "prose") if long_en else [],
            short_en=d["feats"][:5],
            short_vi=(translate_many(d["feats"][:5], "prose")
                      if d["feats"] else []),
            name_vi=name_vi_of(title, translate_many))
        print("  %2d/%d %-42s ts=%-2d dai=%-2d ngan=%-2d %s %s"
              % (i, len(todo), title[:42], len(specs), len(long_en),
                 len(d["feats"][:5]), "img" if d["img_file"] else "---",
                 "bang" if d["ts_file"] else "----"))
        if i % 10 == 0:
            _flush(True)
            save(uniq, done)
    _flush(True)
    save(uniq, done)
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
