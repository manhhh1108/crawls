# -*- coding: utf-8 -*-
"""perma-tec - www.perma-tec.com (EN).
Small catalogue: the sitemap already carries the whole product tree; names come
from the JSON-LD breadcrumb."""
import re
from concurrent.futures import ThreadPoolExecutor
from fetcher import fetch, try_fetch
from common import save, ld_breadcrumb

SM = "https://www.perma-tec.com/en/sitemap.xml"
KEEP = ("/en/lubrication-systems", "/en/accessories-lubricants")
DROPTAIL = ("request-form",)


def crawl():
    locs = sorted(set(l for l in re.findall(r"<loc>([^<]+)</loc>", fetch(SM))
                      if any(k in l for k in KEEP)
                      and not any(d in l for d in DROPTAIL)))
    print("category pages:", len(locs))

    def one(u):
        t = try_fetch(u, retries=3, sleep=0.4)
        return u, (ld_breadcrumb(t, drop=("perma", "home")) if t else None)

    nodes = []
    with ThreadPoolExecutor(max_workers=4) as ex:
        for u, path in ex.map(one, locs):
            if not path:
                print("  [skip]", u)
                continue
            for i in range(1, len(path) + 1):
                nodes.append({"path": path[:i], "type": "category",
                              "url": u if i == len(path) else None, "code": None})
    return nodes


if __name__ == "__main__":
    save("05 Perma", "https://www.perma-tec.com/en", crawl(),
         "Cay danh muc tu sitemap EN + breadcrumb JSON-LD.")
