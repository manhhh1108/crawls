# -*- coding: utf-8 -*-
"""Balluff - www.balluff.com (en-sg).
Sitemap gives areas > groups > product families; names come from the JSON-LD
BreadcrumbList. A handful of single-product families redirect straight to the
article page - for those the name is taken from <title> and the parent names from
the area/group maps built out of the successful pages."""
import re, html as H
from concurrent.futures import ThreadPoolExecutor
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save, ld_breadcrumb

SM = "https://www.balluff.com/assets/sitemap/sitemap-en-sg.xml"
TITLE = re.compile(r"<title>(.*?)</title>", re.S)


def title_name(t):
    m = TITLE.search(t or "")
    if not m:
        return None
    s = H.unescape(re.sub(r"\s+", " ", m.group(1))).strip()
    s = re.split(r"\s+-\s+BALLUFF", s)[0]
    if ")" in s:
        s = s.split(")", 1)[1]
    return s.strip(" -") or None


def ids(u):
    a = re.search(r"/areas/([^/]+)", u)
    g = re.search(r"/groups/([^/]+)", u)
    f = re.search(r"/products/(F[^/]+)$", u)
    return (a.group(1) if a else None, g.group(1) if g else None, f.group(1) if f else None)


def crawl():
    locs = sorted(set(l for l in re.findall(r"<loc>([^<]+)</loc>", fetch(SM))
                      if "/products/areas/" in l))
    print("product URLs:", len(locs))

    def one(u):
        t = try_fetch(u, retries=4, sleep=0.8)
        return u, t, (ld_breadcrumb(t) if t else None)

    got = []
    with ThreadPoolExecutor(max_workers=3) as ex:
        got = list(ex.map(one, locs))

    area, group, nodes = {}, {}, []
    for u, t, path in got:
        if not path:
            continue
        a, g, f = ids(u)
        if a and len(path) >= 1:
            area[a] = path[0]
        if g and len(path) >= 2:
            group[g] = path[1]
        for i in range(1, len(path) + 1):
            nodes.append({"path": path[:i], "type": "category" if i < len(path) else "range",
                          "url": u if i == len(path) else None,
                          "code": (f or g or a) if i == len(path) else None})

    miss = 0
    for u, t, path in got:
        if path or not t:
            continue
        a, g, f = ids(u)
        nm = title_name(t)
        if not nm or a not in area:
            print("  [skip]", u)
            continue
        p = [area[a]] + ([group[g]] if g in group else []) + [nm]
        nodes.append({"path": p, "type": "range", "url": u, "code": f or g})
        miss += 1
    print("  recovered from <title>:", miss)
    return nodes


if __name__ == "__main__":
    save("17 Balluff", "https://www.balluff.com/en-sg", crawl(),
         "Tu sitemap en-sg + breadcrumb JSON-LD. Cap cuoi = ho san pham (F....) ngay tren ma SP (BES/BFT...).")
