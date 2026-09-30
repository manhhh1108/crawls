# -*- coding: utf-8 -*-
"""Leuze - www.leuze.com (en-int, Shopware).
Full category tree from the sitemap; exact names from the microdata breadcrumb.
The deepest category is the level directly above the article number."""
import re
from concurrent.futures import ThreadPoolExecutor
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save, micro_breadcrumb

IDX = "https://www.leuze.com/en-int/sitemap.xml"


def crawl():
    sms = re.findall(r"<loc>([^<]+)</loc>", fetch(IDX))
    locs = []
    for s in sms:
        locs += re.findall(r"<loc>([^<]+)</loc>", fetch(s))
    cats = sorted(set(l for l in locs if "/en-int/products/" in l))
    print("category pages:", len(cats))

    def one(u):
        t = try_fetch(u, retries=2)
        return u, (micro_breadcrumb(t) if t else None)

    nodes = []
    with ThreadPoolExecutor(max_workers=5) as ex:
        for u, path in ex.map(one, cats):
            if not path:
                print("  [skip]", u)
                continue
            for i in range(1, len(path) + 1):
                nodes.append({"path": path[:i], "type": "category",
                              "url": u if i == len(path) else None, "code": None})
    return nodes


if __name__ == "__main__":
    save("08 Leuze", "https://www.leuze.com/en-int", crawl(),
         "Cây danh mục đầy đủ từ sitemap en-int. Cấp sâu cùng = cấp ngay trên mã SP.")
