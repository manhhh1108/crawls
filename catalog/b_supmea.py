# -*- coding: utf-8 -*-
"""Supmea Automation - en.supmea.com (EN).
The nav "Products" dropdown carries the full 2-level tree: each level-1
category is a <li class="dropend"> whose first link is the category and the
nested <ul> holds the level-2 links (the level directly above the models)."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch
from common import save

URL = "https://en.supmea.com/products"
BASE = "https://en.supmea.com"
DROPEND = re.compile(r'<li class="dropend"[^>]*>(.*?)</ul>', re.S)
LINK = re.compile(r'<a href="(https://en\.supmea\.com/([^"]+))"[^>]*class="dropdown-item[^"]*"[^>]*>(.*?)</a>', re.S)
SKIP_L1 = {"products", "system-products", "system-productss", "download", "downloads"}


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


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
            nodes.append({"path": [p1, n], "type": "range", "url": u2, "code": slug2.strip("/")})
    return nodes


if __name__ == "__main__":
    ns = crawl()
    save("12 Supmea", "https://en.supmea.com/", ns,
         "Cay 2 cap tu menu Products cua ban EN; cap 2 la cap ngay tren ma SP cu the.")
