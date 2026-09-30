# -*- coding: utf-8 -*-
"""Guangdong Xingguang / Starshine Drive - en.xgcd.cn (EN).
Level 1 = the category pages in the nav (/product/<id>.html), level 2 = the
series cards on each category page (/products_details/<id>.html), which sit
directly above the concrete model codes."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

BASE = "https://en.xgcd.cn"
ROOT = BASE + "/product/1356565743905951744.html"  # "Products" root page
ROOT_ID = "1356565743905951744"
NAV = re.compile(r'<a[^>]+href="(/product/(\d+)\.html)"[^>]*>(.*?)</a>', re.S)
CARD = re.compile(r'<a href="(/products_details/(\d+)\.html)"[^>]*>(.*?)</a>', re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def crawl():
    t = try_fetch(ROOT, retries=4)
    if not t:
        return []
    cats, seen = [], set()
    for href, cid, label in NAV.findall(t):
        n = txt(label)
        if not n or cid == ROOT_ID or n.lower() in ("products", "more products"):
            continue
        if cid in seen:
            continue
        seen.add(cid)
        cats.append((n, href, cid))
    print("  categories:", len(cats))

    nodes, have = [], set()
    for n, href, cid in cats:
        nodes.append({"path": [n], "type": "category", "url": BASE + href, "code": cid})
        ct = try_fetch(BASE + href, retries=4)
        if not ct:
            continue
        for h2, pid, l2 in CARD.findall(ct):
            n2 = txt(l2)
            k = (n, n2)
            if not n2 or k in have:
                continue
            have.add(k)
            nodes.append({"path": [n, n2], "type": "range", "url": BASE + h2, "code": pid})
    return nodes


if __name__ == "__main__":
    ns = crawl()
    save("14 Xingguang-Starshine", "https://en.xgcd.cn/", ns,
         "Cây 2 cấp từ bản EN: nav /product/<id> (cấp 1) + thẻ series card trên trang (cấp 2, ngay trên mã SP).")
