# -*- coding: utf-8 -*-
"""wenglor - www.wenglor.com (EN).
Category tree from the sitemap (Category + CategoryLanding); exact display names
taken from each category page's breadcrumb. Concrete article codes (P1PC011...)
hang directly under the deepest categories, so the deepest category IS the level
immediately above the product code."""
import re, urllib.parse, html as H
from concurrent.futures import ThreadPoolExecutor
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

SITEMAP = "https://www.wenglor.com/en/sitemap.xml"
BC = re.compile(r'<ol class="breadcrumb">(.*?)</ol>', re.S)
LI = re.compile(r'<li[^>]*>\s*(?:<a[^>]*>)?(.*?)(?:</a>)?\s*</li>', re.S)


def cat_urls():
    idx = fetch(SITEMAP)
    locs = re.findall(r"<loc>([^<]+)</loc>", idx)
    urls = []
    for name in ("CategoryLanding-en", "/Category-en"):
        sm = [l for l in locs if name in l]
        if sm:
            urls += re.findall(r"<loc>([^<]+)</loc>", fetch(sm[0]))
    return sorted(set(urls))


def breadcrumb(url):
    t = try_fetch(url, retries=2)
    if not t:
        return None
    m = BC.search(t)
    if not m:
        return None
    out = []
    for li in LI.findall(m.group(1)):
        txt = H.unescape(re.sub(r"<[^>]+>", "", li)).strip()
        if txt:
            out.append(txt)
    if out and out[0].lower() == "products":
        out = out[1:]
    return out or None


def crawl():
    urls = cat_urls()
    print("category pages:", len(urls))
    nodes = []
    with ThreadPoolExecutor(max_workers=6) as ex:
        res = list(ex.map(breadcrumb, urls))
    for u, path in zip(urls, res):
        if not path:
            print("  [skip]", u)
            continue
        for i in range(1, len(path) + 1):
            nodes.append({"path": path[:i], "type": "category",
                          "url": u if i == len(path) else None,
                          "code": u.rsplit("/", 1)[-1] if i == len(path) else None})
    return nodes


if __name__ == "__main__":
    save("06 wenglor", "https://www.wenglor.com/en/", crawl(),
         "Cay danh muc day du tu sitemap; cap sau cung la cap ngay tren ma SP (vd P1PC011).")
