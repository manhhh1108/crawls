# -*- coding: utf-8 -*-
"""SICK (hang 01) - PHA 1: keo du lieu tho tung dong san pham tu API sick.com.

Dau vao : data/SICK_Web_Products_Statistics.xlsx, sheet "Thong ke SP SICK Web SP",
          cac dong co danh dau 'x' (= dong san pham can viet bai).
Dau ra  : content/out/01 SICK/raw.json  (chua dich, chua ghep anh)

API dung:
  records/detail/<gcode>            -> mo ta, slogan, features, benefits, anh, thuoc tinh
  search ...ProductFamilyCode:<g>   -> danh sach id bien the (idsOnly, rat nhe)
  occ/v2/sg/sick-products?codes=... -> ten model + part number (gon ~0.9KB/SP)
"""
import io, json, math, os, re, sys, html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, REPO)
from shared.fetcher import try_fetch  # noqa: E402

XLSX = os.path.join(REPO, "data", "SICK_Web_Products_Statistics.xlsx")
SHEET = "Thong ke SP SICK Web SP"
OUT_DIR = os.path.join(HERE, "out", "01 SICK")
os.makedirs(OUT_DIR, exist_ok=True)

DETAIL = ("https://www.sick.com/api/fact-finder/records/detail/enSG-sick/%s"
          "?usePersonalization=false&withCampaigns=false&withRecommendations=false"
          "&withSimilarProducts=false")
IDS = ("https://www.sick.com/api/fact-finder/search/enSG-sick?query=%%2A"
       "&filter=DefType%%3AProductVariant&filter=ProductFamilyCode%%3A%s"
       "&page=%d&hitsPerPage=200&showPermutedSearchParams=false&idsOnly=true")
PRODS = "https://api.sick.com/occ/v2/sg/sick-products?productCodes=%s&lang=en&curr=SGD"
REF = "https://www.sick.com/sg/en/"
SITE = "https://www.sick.com"


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", " ", s or ""))).strip()


def paras(htmltext):
    """<p>..</p> -> list doan van; neu khong co the <p> thi tach theo xuong dong."""
    if not htmltext:
        return []
    out = [txt(p) for p in re.findall(r"<p[^>]*>(.*?)</p>", htmltext, re.S)]
    out = [p for p in out if len(p) > 30]
    if not out:
        s = txt(htmltext)
        out = [s] if len(s) > 30 else []
    return out


def attr_rows(mv):
    """Cac khoa dang 'Ten<@>id[~~unit]' -> [(ten, don vi, [gia tri])] giu thu tu."""
    rows = []
    for k, v in mv.items():
        if "<@>" not in k:
            continue
        name = k.split("<@>", 1)[0].strip()
        tail = k.split("<@>", 1)[1]
        unit = ""
        m = re.search(r"~~([^~<]+)$", tail)
        if m:
            unit = m.group(1).strip()
        vals = v if isinstance(v, list) else [v]
        vals = [txt(str(x)) for x in vals if str(x).strip() not in ("", "None")]
        vals = list(dict.fromkeys(vals))
        if name and vals:
            rows.append([name, unit, vals])
    rows.sort(key=lambda r: r[0].lower())
    return rows


def jget(url, **kw):
    kw.setdefault("retries", 3)
    kw.setdefault("timeout", 90)
    kw.setdefault("referer", REF)
    t = try_fetch(url, json_mode=True, **kw)
    if not t:
        return None
    try:
        return json.loads(t)
    except Exception:
        return None


def variants(code, total_hint=0):
    ids, page, pages = [], 1, 1
    while page <= pages and page <= 40:
        d = jget(IDS % (code, page))
        if not d:
            break
        hits = d.get("hits") or []
        ids += [h.get("id") for h in hits if h.get("id")]
        tot = d.get("totalHits") or len(ids)
        pages = max(1, math.ceil(tot / 200))
        if not hits:
            break
        page += 1
    out = []
    for i in range(0, len(ids), 50):
        chunk = ids[i:i + 50]
        d = jget(PRODS % ",".join(chunk))
        for p in (d or {}).get("products") or []:
            out.append({"name": p.get("name") or "",
                        "part": p.get("partNumber") or "",
                        "url": SITE + (p.get("url") or "") if p.get("url") else ""})
    return out, len(ids)


def rows_from_xlsx():
    from openpyxl import load_workbook
    ws = load_workbook(XLSX, data_only=True)[SHEET]
    out = []
    for r in ws.iter_rows(values_only=True):
        if not r or not r[0] or not re.match(r"^[\d.]+$", str(r[0]).strip()):
            continue
        if len(r) < 7 or str(r[6]).strip().lower() != "x":
            continue
        out.append({"stt": str(r[0]).strip(), "name": (r[1] or "").strip(),
                    "code": (r[2] or "").strip(),
                    "n_products": int(float(r[3] or 0)),
                    "cat_vi": (r[4] or "").strip(), "cat_en": (r[5] or "").strip()})
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    lines = rows_from_xlsx()
    print("dong san pham can lam:", len(lines))
    done, fails = [], []
    for i, ln in enumerate(lines, 1):
        d = jget(DETAIL % ln["code"])
        mv = ((d or {}).get("record") or {}).get("masterValues")
        if not mv:
            fails.append(ln["code"])
            print("  [FAIL detail] %s %s" % (ln["code"], ln["name"]))
            continue
        gallery = []
        try:
            g = mv.get("Gallery")
            g = json.loads(g) if isinstance(g, str) else (g or [])
            gallery = [x.get("URL") for x in g if x.get("URL")]
        except Exception:
            pass
        main_img = (mv.get("DefPictureUrl") or [None])[0] if isinstance(
            mv.get("DefPictureUrl"), list) else mv.get("DefPictureUrl")
        if not main_img and gallery:
            main_img = gallery[0]

        vs, ntot = variants(ln["code"], ln["n_products"])
        rec = dict(ln)
        rec.update({
            "api_name": mv.get("DefName") or ln["name"],
            "name_addon": mv.get("DefNameAddon") or "",
            "family_group": mv.get("ProductFamilyGroup") or "",
            "slogan": txt(mv.get("Slogan") or ""),
            "overview_en": paras(mv.get("Description") or mv.get("DefDescription") or ""),
            "features_en": [txt(x) for x in (mv.get("Features") or []) if txt(x)],
            "benefits_en": [txt(x) for x in (mv.get("Benefits") or []) if txt(x)],
            "applications_en": [txt(x) for x in (mv.get("FieldsOfApplication") or []) if txt(x)],
            "meta_title": mv.get("MetaTitle") or "",
            "meta_desc": txt(mv.get("MetaDescription") or ""),
            "path_en": (mv.get("DefCategoryPathEn") or [])[:],
            "url": SITE + "/sg/en" + (mv.get("DefUrl") or ""),
            "def_type": mv.get("DefType") or "",
            "phase": mv.get("PhaseStatus") or "",
            "attrs": attr_rows(mv),
            "main_image": (SITE + main_img) if main_img else "",
            "gallery": [SITE + u for u in gallery[:6]],
            "variants": vs, "variant_total": ntot,
        })
        done.append(rec)
        print("  %3d/%d %-34s ov=%d ft=%d bn=%d attr=%2d var=%4d %s"
              % (i, len(lines), ln["name"][:34], len(rec["overview_en"]),
                 len(rec["features_en"]), len(rec["benefits_en"]),
                 len(rec["attrs"]), ntot, "img" if main_img else "---"))
        if i % 25 == 0:
            with io.open(os.path.join(OUT_DIR, "raw.json"), "w", encoding="utf-8") as f:
                json.dump({"brand": "01 SICK", "lines": done}, f, ensure_ascii=False)
    with io.open(os.path.join(OUT_DIR, "raw.json"), "w", encoding="utf-8") as f:
        json.dump({"brand": "01 SICK", "lines": done}, f, ensure_ascii=False)
    print("\nxong %d dong, loi %d" % (len(done), len(fails)))
    if fails:
        print("  loi:", fails)


if __name__ == "__main__":
    main()
