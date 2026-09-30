# -*- coding: utf-8 -*-
"""Schneider Electric - www.se.com (ca/en: the most complete English catalogue).
Category / sub-category / range sitemaps + the JSON-LD breadcrumb on each page.
"Range" (e.g. Harmony XB4) is the level directly above the commercial reference."""
import re, html as H
from concurrent.futures import ThreadPoolExecutor
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save, ld_breadcrumb

LOC = "ca/en"
CC = "CA"
BASE = "https://www.se.com/%s" % LOC
SMI = BASE + "/product-%s/google-%s-sitemapindex-%s-en.xml"
RANGE_CARD = re.compile(
    r'href="(/[^"]*?/product-range/([^/"?]+)/[^"]*)"[\s\S]{0,4000}?<!--RANGE-NAME_BEGIN-->([\s\S]*?)<!--RANGE-NAME_END-->')
DROP = ("all products", "home", "products")


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def sitemap(kind):
    idx = fetch(SMI % (kind, kind, CC))
    out = []
    for s in re.findall(r"<loc>([^<]+)</loc>", idx):
        out += re.findall(r"<loc>([^<]+)</loc>", fetch(s))
    return sorted(set(out))


def crawl():
    cats = sitemap("category")
    subs = sitemap("subcategory")
    print("categories:", len(cats), "subcategories:", len(subs))

    def bc(u):
        t = try_fetch(u, retries=3, sleep=0.4)
        if not t:
            return u, None, None
        path = ld_breadcrumb(t, drop=DROP) or []
        ranges = []
        seen = set()
        for href, slug, name in RANGE_CARD.findall(t):
            n = txt(name)
            if n and slug not in seen:
                seen.add(slug)
                ranges.append((n, "https://www.se.com" + href.split("?")[0], slug))
        return u, path, ranges

    catnames = set()
    nodes = []
    with ThreadPoolExecutor(max_workers=4) as ex:
        catres = list(ex.map(bc, cats))
    for u, path, ranges in catres:
        if not path:
            print("  [skip cat]", u)
            continue
        catnames.add(path[-1])
    print("  category names:", len(catnames))

    def trim(path):
        # breadcrumb can start with a market segment above the real category
        while len(path) > 1 and path[0] not in catnames:
            path = path[1:]
        return path

    with ThreadPoolExecutor(max_workers=4) as ex:
        subres = list(ex.map(bc, subs))

    # a category page also lists every range of its sub-categories - keep those
    # only where the range has no sub-category of its own
    placed = set()
    for u, path, ranges in subres:
        for n, ru, slug in ranges or []:
            placed.add(slug)

    for u, path, ranges in catres:
        if not path:
            continue
        p = trim(path)
        nodes.append({"path": p, "type": "category", "url": u,
                      "code": u.rstrip("/").rsplit("/", 1)[-1].split("-")[0]})
        for n, ru, slug in ranges or []:
            if slug in placed:
                continue
            nodes.append({"path": p + [n], "type": "range", "url": ru, "code": slug.split("-")[0]})

    for u, path, ranges in subres:
        if not path:
            print("  [skip sub]", u)
            continue
        p = trim(path)
        for i in range(1, len(p) + 1):
            nodes.append({"path": p[:i], "type": "category",
                          "url": u if i == len(p) else None, "code": None})
        for n, ru, slug in ranges or []:
            nodes.append({"path": p + [n], "type": "range", "url": ru, "code": slug.split("-")[0]})
    return nodes


if __name__ == "__main__":
    save("11 Schneider", "https://www.se.com/ca/en", crawl(),
         "Category > Subcategory > Range từ sitemap ca/en. Range là cấp ngay trên mã thương mại.")
