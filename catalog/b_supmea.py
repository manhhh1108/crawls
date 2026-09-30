# -*- coding: utf-8 -*-
"""Supmea Automation - en.supmea.com (EN).
The nav "Products" dropdown carries the full 2-level tree: each level-1
category is a <li class="dropend"> whose first link is the category and the
nested <ul> holds the level-2 links. Every level-2 page then lists its
concrete products (links of the form /<level2-slug>/<product>), whose names
start with the model code (FMC240, FMX400, SUP-...)."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

URL = "https://en.supmea.com/products"
BASE = "https://en.supmea.com"
DROPEND = re.compile(r'<li class="dropend"[^>]*>(.*?)</ul>', re.S)
LINK = re.compile(r'<a href="(https://en\.supmea\.com/([^"]+))"[^>]*class="dropdown-item[^"]*"[^>]*>(.*?)</a>', re.S)
SKIP_L1 = {"products", "download", "downloads"}
MODEL = re.compile(r"^([A-Z][A-Za-z0-9]*\d[A-Za-z0-9-]*|SUP-[A-Za-z0-9-]+)\b")


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def products(slug, url):
    """Product links on a level-2 page: /<slug>/<tail>."""
    t = try_fetch(url)
    if not t:
        return []
    pat = re.compile(r'<a[^>]+href="(%s/%s/([^"/]+))"[^>]*>(.*?)</a>'
                     % (re.escape(BASE), re.escape(slug)), re.S)
    best = {}
    for u, tail, label in pat.findall(t):
        n = txt(label)
        if not n or n.lower() in ("view more", "more", "read more"):
            continue
        if len(n) > len(best.get(u, ("", ""))[0]):
            best[u] = (n, tail)
    out = []
    for u, (n, tail) in best.items():
        m = MODEL.match(n)
        out.append({"name": n, "url": u, "code": m.group(1) if m else tail})
    return out


def crawl():
    t = fetch(URL)
    nodes, seen = [], set()
    for blk in DROPEND.findall(t):
        links = LINK.findall(blk)
        if not links:
            continue
        u1, slug1, l1 = links[0]
        p1 = txt(l1)
        if not p1 or slug1.strip("/") in SKIP_L1:
            continue
        if p1 not in seen:
            seen.add(p1)
            nodes.append({"path": [p1], "type": "category", "url": u1, "code": slug1.strip("/")})
        for u2, slug2, l2 in links[1:]:
            n = txt(l2)
            k = (p1, n)
            if not n or k in seen:
                continue
            seen.add(k)
            s2 = slug2.strip("/")
            nodes.append({"path": [p1, n], "type": "range", "url": u2, "code": s2})
            for pr in products(s2, u2):
                nodes.append({"path": [p1, n, pr["name"]], "type": "product",
                              "url": pr["url"], "code": pr["code"]})
    return nodes


if __name__ == "__main__":
    ns = crawl()
    n_prod = sum(1 for n in ns if n["type"] == "product")
    print("  products:", n_prod)
    save("12 Supmea", "https://en.supmea.com/", ns,
         "Cây 2 cấp từ menu Products bản EN + mã SP (level 3, type=product) từ trang từng dòng SP.")
