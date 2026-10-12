# -*- coding: utf-8 -*-
"""SICK (hang 01) - cay danh muc tu www.sick.com/sg/en.

Nguon: API fact-finder cua chinh web SG/EN
  search ...DefType:Category       -> cac cap danh muc (272 ban ghi, ca Products/Applications/Industries)
  search ...DefType:MasterProduct  -> dong san pham (gXXXXXX), la cap ngay tren ma SP
  search ...ProductFamilyCode:<g>  -> so variant cua tung dong SP; dong nao <= MAX_VAR
                                      thi lay luon ten variant (VD Lector65x System Pro/Core/Prime)

Chi giu nhanh "Products" va dung PrimaryCategoryPath (moi dong SP xuat hien dung 1 lan,
du web SICK xep no vao nhieu nhanh).
"""
import os, sys, json, math

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, REPO)
sys.path.insert(0, HERE)
from shared.fetcher import try_fetch          # noqa: E402
from common import save                       # noqa: E402

SITE = "https://www.sick.com"
BASE = SITE + "/sg/en"
REF = BASE + "/"
SEARCH = (SITE + "/api/fact-finder/search/enSG-sick?query=%%2A&filter=DefType%%3A%s"
          "&page=%d&hitsPerPage=%d&showPermutedSearchParams=false")
VARS = (SITE + "/api/fact-finder/search/enSG-sick?query=%%2A&filter=DefType%%3AProductVariant"
        "&filter=ProductFamilyCode%%3A%s&page=1&hitsPerPage=%d&showPermutedSearchParams=false%s")

MAX_VAR = 5       # dong SP co toi da bay nhieu variant thi liet ke variant ra thanh dong rieng
ROOT = "Products"


def jget(url):
    t = try_fetch(url, json_mode=True, referer=REF, retries=4, timeout=90)
    if not t:
        return None
    try:
        return json.loads(t)
    except Exception:
        return None


def all_of(deftype, per=200):
    """Lay toan bo ban ghi mot DefType."""
    out, page, pages = [], 1, 1
    while page <= pages:
        d = jget(SEARCH % (deftype, page, per))
        if not d:
            break
        hits = d.get("hits") or []
        out += [h.get("masterValues") or {} for h in hits]
        pages = max(1, math.ceil((d.get("totalHits") or len(out)) / per))
        if not hits:
            break
        page += 1
    print("  %-14s %d ban ghi" % (deftype, len(out)))
    return out


def ppath(mv):
    """Duong dan chinh, da bo cap goc 'Products'."""
    p = mv.get("PrimaryCategoryPath") or mv.get("DefCategoryPath") or []
    if p and isinstance(p[0], list):
        p = p[0]
    if not p:
        p = mv.get("DefCategoryPathEn") or []
    p = [str(x).strip() for x in p if str(x).strip()]
    if not p or p[0] != ROOT:
        return None
    return p[1:]


def url_of(mv):
    u = mv.get("DefUrl") or ""
    return (BASE + u) if u else ""


def variants(code):
    """(so variant, [ban ghi variant neu it]) cho mot dong SP."""
    d = jget(VARS % (code, 1, "&idsOnly=true"))
    n = (d or {}).get("totalHits") or 0
    if not n or n > MAX_VAR:
        return n, []
    d = jget(VARS % (code, MAX_VAR, ""))
    return n, [h.get("masterValues") or {} for h in (d or {}).get("hits") or []]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    nodes, seen = [], set()

    for mv in all_of("Category"):
        p = ppath(mv)
        if not p:
            continue
        k = tuple(p)
        if k in seen:
            continue
        seen.add(k)
        nodes.append({"path": p, "type": "category", "url": url_of(mv),
                      "code": mv.get("DefArticleNo") or ""})

    fams = []
    for mv in all_of("MasterProduct"):
        p = ppath(mv)
        if not p:
            continue
        k = tuple(p)
        if k in seen:
            continue
        seen.add(k)
        code = mv.get("DefArticleNo") or ""
        nodes.append({"path": p, "type": "range", "url": url_of(mv), "code": code})
        fams.append((p, code))

    print("  -> %d danh muc, %d dong SP (nhanh Products)" % (len(nodes) - len(fams), len(fams)))

    nvar_tot, nsmall, nrow = 0, 0, 0
    fams.sort(key=lambda x: [s.lower() for s in x[0]])
    for i, (p, code) in enumerate(fams, 1):
        if not code:
            continue
        n, vs = variants(code)
        nvar_tot += n
        if vs:
            nsmall += 1
        for v in vs:
            nm = (v.get("DefName") or "").strip()
            if not nm:
                continue
            vp = p + [nm]
            if tuple(vp) in seen:
                continue
            seen.add(tuple(vp))
            nodes.append({"path": vp, "type": "product", "url": url_of(v),
                          "code": (v.get("PartNumber") or "").strip() or v.get("DefArticleNo") or ""})
            nrow += 1
        if i % 50 == 0 or i == len(fams):
            print("    variant %4d/%d  (%d dong SP nho, %d dong variant)"
                  % (i, len(fams), nsmall, nrow))

    note = ("Crawl tu API fact-finder cua www.sick.com/sg/en (%s). Lay nhanh Products theo "
            "PrimaryCategoryPath. Cap danh dau 'x' = dong san pham (ma gXXXXXX) - cap ngay tren ma SP. "
            "Nhung dong SP chi co <=%d variant thi liet ke luon variant (chu nghieng xam, ma pXXXXXX "
            "hoac part number) vi day thuc chat la cac dong nho rieng, VD Lector65x System "
            "Pro/Core/Prime. Tong %d san pham cu the tren web."
            % (__import__("datetime").date.today().isoformat(), MAX_VAR, nvar_tot))
    save("01 SICK", BASE, nodes, note)


if __name__ == "__main__":
    main()
