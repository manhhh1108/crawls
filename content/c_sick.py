# -*- coding: utf-8 -*-
"""SICK (hang 01) - PHA 2: tu raw.json -> content.json (da dich + da ghep anh).

Chay sau c_sick_fetch.py:
    python c_sick.py            # tat ca
    python c_sick.py 20         # chi 20 dong dau (de xem thu)

Dich bang vi_engine (tu dien thuat ngu truoc, may dich sau, co mask ma/don vi).
"""
import io
import json
import os
import re
import subprocess
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, REPO)
sys.path.insert(0, HERE)

import brandimg                                   # noqa: E402
import ntrend_cat                                 # noqa: E402
import sick_tdo                                   # noqa: E402
from translate import _flush                      # noqa: E402
from vi_engine import (lookup, suspicious, translate_many,  # noqa: E402
                       add_keep_names)

# tu tieng Anh pho thong -> KHONG coi la ten rieng (de con dich binh thuong)
_COMMON = set("""safety sensors sensor cables cable systems system controllers controller
relays relay switches switch scanners scanner monitor monitors modules module light
curtains curtain barriers barrier encoders encoder cameras camera lasers laser radar
connection connections accessories accessory products product series mats mat bumpers
speed distance level flow pressure temperature vision network networks gateway gateways
interface interfaces software hardware kits kit sets set tools tool""".split())


def is_proper_name(s):
    """Ten rieng cua SICK (giu nguyen khi dich) hay cum tieng Anh thuong?"""
    s = (s or "").strip()
    if not s or len(s) > 40:
        return False
    if re.search(r"\d", s):                       # W16, S300, IME2S, safeRS3
        return True
    if re.search(r"[a-z][A-Z]", s):               # microScan3, deTec, flexLock
        return True
    words = re.findall(r"[A-Za-z][A-Za-z/\-]*", s)
    if not words:
        return False
    if all(w.lower() in _COMMON for w in words):  # "Safety light curtains" -> dich
        return False
    return len(words) <= 2                        # Lector, ReLy, Flexi Loop

IMG_EXT = ".jpg"        # duoi anh thanh pham trong images/
LOGO_FALLBACK = os.path.join(HERE, "assets", "sick-logo.jpg")
OUT_DIR = os.path.join(HERE, "out", "01 SICK")
SRC_IMG = os.path.join(OUT_DIR, "src_img")
BRAND_IMG = os.path.join(OUT_DIR, "images")
for p in (OUT_DIR, SRC_IMG, BRAND_IMG):
    os.makedirs(p, exist_ok=True)

TDO = sick_tdo.load()   # {gcode: [[ten thong so, gia tri], ...]} - xem sick_tdo.py

MAX_OVERVIEW = 4        # so doan mo ta dai
MAX_BULLETS = 10        # so gach dau dong (Features + Benefits)
MAX_VALS = 6            # so gia tri hien thi cho 1 thong so
BRAND = "SICK"


def slugify(name):
    s = re.sub(r"[^a-z0-9]+", "-", (name or "").lower()).strip("-")
    return re.sub(r"-{2,}", "-", s) or "sp"


# Facet o tab Variants co ca metadata catalog - dang web vo nghia, phai bo.
FACET_SKIP = re.compile(
    r"^\s*(new\s*products?|products?|product\s*(sub)?segments?|product\s*family|"
    r"product\s*categor(y|ies)|part\s*(no|number)s?|art(icle)?\.?\s*(no|number)?s?|"
    r"order(ing)?\s*(no|number|code)s?|type\s*codes?|eclass|etim|unspsc|gtin|ean|"
    # 'Description' bi bo vi no la mo ta cua MOT variant bat ky trong dong
    # (thu tu API khong on dinh), de vao bang thong so se gay hieu nham.
    r"customs\s*tariff.*|country\s*of\s*origin|brand|manufacturer|description|"
    r"bestseller|web\s*sellable|phase\s*status|visibility.*|seo.*|"
    r"show\s*in\s*product\s*listing|release\s*date|signature|timestamp)"
    r"\.?\s*$", re.I)


def facet_specs(rec, max_rows=12):
    """Facet tab Variants -> [[ten, gia tri]] - dung khi PDF khong co bang TDO.

    Moi facet gom gia tri cua MOI variant trong dong, nen chi lay MAX_VALS gia
    tri dau roi ghi '(+n tuy chon khac)' - dung quy uoc cua bang thong so.
    """
    out = []
    for name, unit, vals in rec.get("attrs", []):
        if FACET_SKIP.match(str(name)):
            continue
        vals = [str(v).strip() for v in vals if str(v).strip()]
        if not vals:
            continue
        shown = vals[:MAX_VALS]
        sfx = (" (+%d tuỳ chọn khác)" % (len(vals) - len(shown))) if len(vals) > len(shown) else ""
        out.append([name + (" [%s]" % unit if unit else ""), "; ".join(shown) + sfx])
        if len(out) >= max_rows:
            break
    return out


def type_of(rec):
    """Loai san pham: uu tien nhom ho SP cua SICK, sau do cap 2 trong breadcrumb."""
    en = (rec.get("family_group") or rec.get("name_addon") or "").strip()
    if not en:
        path = rec.get("path_en") or []
        en = path[2] if len(path) > 2 else (path[-1] if path else "Sensor")
    en = en.rstrip(".")
    vn = lookup(en) or lookup(en.rstrip("s")) or translate_many([en], "key")[0]
    # bo dang so nhieu tieng Anh trong ten loai
    return en, vn


def download(url, dst):
    if os.path.exists(dst) and os.path.getsize(dst) > 500:
        return True
    try:
        p = subprocess.run(["curl", "-sS", "-L", "--max-time", "60", "-o", dst, url],
                           capture_output=True, timeout=90)
    except Exception:
        return False
    return p.returncode == 0 and os.path.exists(dst) and os.path.getsize(dst) > 500


def build_one(rec, review):
    name = rec["name"]
    slug = slugify(name)
    type_en, type_vi = type_of(rec)
    cat_vi = ntrend_cat.cat_vi(name, rec.get("cat_vi", ""))

    overview_en = rec.get("overview_en", [])[:MAX_OVERVIEW]
    bullets_en = (rec.get("features_en", []) + rec.get("benefits_en", []))[:MAX_BULLETS]
    bullets_en = [b.strip() for b in bullets_en if b.strip()]

    # Thong so ky thuat = bang TECHNICAL DATA OVERVIEW trong tai lieu Product
    # overview cua hang (review 13/10). Dong nao tai lieu khong co muc do thi
    # de trong - build_sick_excel se in BANG VARIANT thay the.
    tdo = TDO.get(rec.get("code")) or []
    if not tdo and bullets_en:
        # Tai lieu Product overview cua dong nay khong co muc TECHNICAL DATA
        # OVERVIEW (49/558 dong). Bai van co Features/Benefits nen de trong bang
        # thong so thi thieu -> bu bang facet o tab Variants, DA LOC rac.
        tdo = facet_specs(rec)
    a_keys = [k for k, _v in tdo]
    a_vals = [v for _k, v in tdo]
    # mo ta tung variant (cot Description tren web) - chi co o cac dong se in
    # bang Variant
    v_desc = [x for x in dict.fromkeys(
        (v.get("desc") or "").strip() for v in rec.get("variants", [])) if x]
    # Slogan (hang) + Linh vuc ung dung da bo khoi bai dang -> khong dich nua
    prose_src = overview_en + bullets_en

    keys_vi = translate_many(a_keys, "key")
    vals_vi = translate_many(a_vals, "value")
    prose_vi = translate_many(prose_src, "prose")
    desc_vi = dict(zip(v_desc, translate_many(v_desc, "value"))) if v_desc else {}
    slogan_vi = ""

    n_ov, n_bl = len(overview_en), len(bullets_en)
    ov_vi = prose_vi[:n_ov]
    bl_vi = prose_vi[n_ov:n_ov + n_bl]
    app_vi = []

    specs_en = [[k, v] for k, v in tdo]
    specs_vi = [[kv, vv] for kv, vv in zip(keys_vi, vals_vi)]

    # soat dich
    for field, pairs in (("Mô tả dài", zip(overview_en, ov_vi)),
                         ("Mô tả ngắn", zip(bullets_en, bl_vi)),
                         ("Tên thông số", ((a[0], b[0]) for a, b in zip(specs_en, specs_vi))),
                         ("Giá trị thông số", ((a[1], b[1]) for a, b in zip(specs_en, specs_vi)))):
        for en, vn in pairs:
            why = suspicious(en, vn)
            if why:
                review.append([name, field, why, en, vn])

    img_file = None
    url = rec.get("main_image") or ""
    if url:
        ext = os.path.splitext(url.split("?")[0])[1] or ".png"
        src = os.path.join(SRC_IMG, slug + ext)
        if download(url, src):
            try:
                img_file = brandimg.compose(src, os.path.join(BRAND_IMG, slug + IMG_EXT))
            except Exception as e:
                print("  [img fail]", name, repr(e))
    if not img_file and os.path.exists(LOGO_FALLBACK):
        # dong SP khong co anh tren web SICK -> dung logo SICK lam anh thay the
        try:
            img_file = brandimg.compose(LOGO_FALLBACK,
                                        os.path.join(BRAND_IMG, slug + IMG_EXT))
        except Exception as e:
            print("  [logo fail]", name, repr(e))

    return {
        "name": name, "slug": slug, "url_src": rec.get("url", ""),
        "path": rec.get("path_en") or ["Products", name],
        "brand": BRAND, "brand_field": "SICK Sensor",
        "code": rec.get("code", ""), "stt_src": rec.get("stt", ""),
        "type_en": type_en, "type_vi": type_vi,
        "cat_vi": cat_vi, "cat_en": ntrend_cat.cat_en(cat_vi, rec.get("cat_en", "")),
        "slogan_en": rec.get("slogan", ""),
        "slogan_vi": slogan_vi,
        "overview_en": overview_en, "overview_vi": ov_vi,
        "features_en": bullets_en, "features_vi": bl_vi,
        "applications_en": rec.get("applications_en", []), "applications_vi": app_vi,
        "specs_en": specs_en, "specs_vi": specs_vi,
        "items": [{"name": v["name"], "material": "", "art": v["part"],
                   "desc": (v.get("desc") or "").strip(),
                   "desc_vi": desc_vi.get((v.get("desc") or "").strip(), "")}
                  for v in rec.get("variants", [])],
        "has_tdo": bool(tdo),
        "variant_total": rec.get("variant_total", 0),
        "n_products": rec.get("n_products", 0),
        "img_src_url": url, "img_file": img_file,
        "not_on_web": bool(rec.get("not_on_web")),
    }


def prewarm(lines, chunk=400):
    """Gom TOAN BO chuoi cua ca 558 bai roi dich 1 luot, thay vi dich tung bai.

    Dich tung bai thi moi lo chi co ~5-10 chuoi -> moi request API cho duoc rat
    it, dot sach hang muc request/ngay (free tier Gemini = 500 req/ngay/model).
    Gom lai thi 1 request cho duoc hang tram chuoi, giam so request ~50 lan.
    Dich xong nam trong cache, vong lap chinh ben duoi chi viec lay ra.
    """
    bag = {"key": set(), "value": set(), "prose": set()}
    for r in lines:
        for k, v in TDO.get(r.get("code")) or []:
            bag["key"].add(k)
            bag["value"].add(v)
        for v in r.get("variants", []):
            if (v.get("desc") or "").strip():
                bag["value"].add(v["desc"].strip())
        bag["prose"].update(r.get("overview_en") or [])
        bag["prose"].update(r.get("features_en") or [])
        bag["prose"].update(r.get("benefits_en") or [])
        for n in (r.get("family_group"), r.get("name_addon")):
            if n:
                bag["key"].add(n.rstrip("."))
    tot = sum(len(v) for v in bag.values())
    print("prewarm: gom %d chuoi (key=%d value=%d prose=%d), dich theo lo %d"
          % (tot, len(bag["key"]), len(bag["value"]), len(bag["prose"]), chunk))
    done = 0
    for kind in ("key", "value", "prose"):
        items = sorted(x for x in bag[kind] if x and str(x).strip())
        for i in range(0, len(items), chunk):
            part = items[i:i + chunk]
            translate_many(part, kind)
            done += len(part)
            print("  prewarm %-5s %5d/%-5d (tong %d/%d)"
                  % (kind, min(i + chunk, len(items)), len(items), done, tot))
            _flush(True)
    print("prewarm xong.")


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    limit = int(sys.argv[1]) if len(sys.argv) > 1 else 0
    with io.open(os.path.join(OUT_DIR, "raw.json"), encoding="utf-8") as f:
        lines = json.load(f)["lines"]

    # vai dong SICK da go khoi web sg/en -> van tao bai (dung mau van), danh dau ro
    from c_sick_fetch import rows_from_xlsx
    have = {l["code"] for l in lines}
    order = {}
    missing = []
    for n, r in enumerate(rows_from_xlsx()):
        order[r["code"]] = n
        if r["code"] not in have:
            r.update({"url": "", "attrs": [], "variants": [], "variant_total": 0,
                      "overview_en": [], "features_en": [], "benefits_en": [],
                      "applications_en": [], "main_image": "", "path_en": [],
                      "family_group": "", "slogan": "", "not_on_web": True})
            missing.append(r)
    if missing:
        print("khong con tren web SICK (van tao bai theo mau):",
              ", ".join(m["name"] for m in missing))
        lines += missing
    lines.sort(key=lambda l: order.get(l["code"], 10 ** 6))
    # ten rieng cua SICK (dong SP + model) -> khong dich
    # chi can nap ten KHONG co chu so; ten co chu so (W16, FLX3-CPUN100...) da
    # duoc RX_CODE trong vi_engine bat san.
    # Ten da co ban dich tieng Viet (xem sick_name_translate.py) thi KHONG phai
    # ten rieng -> de may dich dich binh thuong, neu khong cau van se ra kieu
    # "GPS antennas tu SICK cung cap..." (review 13/10 muc 4).
    from sick_names_vi import name_vi
    names = set()
    for r in lines:
        for n in (r.get("name"), r.get("api_name")):
            n = (n or "").strip()
            if n and name_vi(n) != n:
                continue
            if n and is_proper_name(n) and not re.search(r"\d", n):
                names.add(n)
    names |= {"Safety Designer", "Flexi Loop", "Safe Motion", "Smart Task",
              "Safety Configuration Deployment", "EFI", "EFI-pro", "ReLy",
              "BluePilot", "AppSpace", "SOPAS", "SOPAS ET"}
    print("ten rieng giu nguyen:", add_keep_names(names))

    if limit:
        lines = lines[:limit]
    print("so dong:", len(lines))

    prewarm(lines)

    arts, review = [], []
    for i, rec in enumerate(lines, 1):
        arts.append(build_one(rec, review))
        a = arts[-1]
        print("  %3d/%d %-32s ov=%d bl=%d spec=%2d var=%4d %s"
              % (i, len(lines), a["name"][:32], len(a["overview_vi"]),
                 len(a["features_vi"]), len(a["specs_vi"]), len(a["items"]),
                 "img" if a["img_file"] else "---"))
        if i % 20 == 0:
            _flush(True)
            with io.open(os.path.join(OUT_DIR, "content.json"), "w", encoding="utf-8") as f:
                json.dump({"brand": "01 SICK", "articles": arts}, f, ensure_ascii=False)
    _flush(True)
    with io.open(os.path.join(OUT_DIR, "content.json"), "w", encoding="utf-8") as f:
        json.dump({"brand": "01 SICK", "articles": arts}, f, ensure_ascii=False)
    with io.open(os.path.join(OUT_DIR, "review_dich.json"), "w", encoding="utf-8") as f:
        json.dump(review, f, ensure_ascii=False, indent=1)
    print("\nxong %d bai | can soat dich: %d dong -> review_dich.json"
          % (len(arts), len(review)))


if __name__ == "__main__":
    main()
