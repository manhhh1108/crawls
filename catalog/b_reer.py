# -*- coding: utf-8 -*-
"""ReeR - www.reersafety.com (EN, WooCommerce).
Category tree from the product_cat sitemaps (EN locale); names from the Yoast
JSON-LD BreadcrumbList. The deepest category is the level above the model code."""
import re
from concurrent.futures import ThreadPoolExecutor
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import try_fetch
from common import save, ld_breadcrumb

SMS = ["https://www.reersafety.com/product_cat-sitemap%s.xml" % i for i in ("", "2", "3", "4", "5")]


def crawl():
    locs = []
    for s in SMS:
        t = try_fetch(s, retries=3, timeout=120)
        if t:
            locs += re.findall(r"<loc>([^<]+)</loc>", t)
    cats = sorted(set(l for l in locs if l.startswith("https://www.reersafety.com/en/product-category/")))
    print("EN category pages:", len(cats))

    def one(u):
        t = try_fetch(u, retries=3, timeout=120, sleep=0.5)
        return u, (ld_breadcrumb(t) if t else None)

    nodes = []
    with ThreadPoolExecutor(max_workers=4) as ex:
        for u, path in ex.map(one, cats):
            if not path:
                print("  [skip]", u)
                continue
            for i in range(1, len(path) + 1):
                nodes.append({"path": path[:i], "type": "category",
                              "url": u if i == len(path) else None, "code": None})
    return nodes


if __name__ == "__main__":
    save("03 Reer", "https://www.reersafety.com/en/", crawl(),
         "Cay danh muc EN tu product_cat sitemap + breadcrumb JSON-LD.")
