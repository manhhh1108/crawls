# -*- coding: utf-8 -*-
"""ifm - www.ifm.com (US/EN).
L1..L3 from the navbar REST service; L4 (group headings) and L5 (tiles) are read
from the rendered category pages. L5 is the level directly above the article
number (e.g. O1D100)."""
import re, json, base64, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch
from shared.browser import Renderer
from common import save

BASE = "https://www.ifm.com"
NAV = BASE + "/restservices/us/en/menu/navbar"
B64 = re.compile(r'"([A-Za-z0-9+/=]{300,})"')
HEAD = re.compile(r"<h[1-4][^>]*>(.*?)</h[1-4]>", re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def nav_tree():
    data = json.loads(fetch(NAV, json_mode=True, referer=BASE + "/us/en/category"))
    out = []
    for bar in data:
        if bar.get("menuType") != "PRODUCTS":
            continue
        for l1 in bar.get("products") or []:
            n1 = l1.get("label")
            for l2 in l1.get("entries") or []:
                n2 = l2.get("label")
                for l3 in l2.get("entries") or []:
                    out.append((n1, n2, l3.get("label"), l3.get("url")))
    return out


def extra_rows():
    """The Accessories branch (400_*) is only in the product-line landing page,
    not in the navbar service."""
    t = fetch(BASE + "/us/en/category")
    out = []
    for m in re.finditer(r'<a[^>]+href="(/us/en/category/(400_\d+))"[^>]*>(.*?)</a>', t, re.S):
        out.append(("Accessories", txt(m.group(3)), None, m.group(1)))
    return out


def crawl():
    rows = nav_tree()
    print("navbar L3 entries:", len(rows))
    nodes, l1s, l2s = [], set(), set()
    for n1, n2, n3, url in rows:
        if n1 not in l1s:
            nodes.append({"path": [n1], "type": "category", "url": None, "code": None}); l1s.add(n1)
        if (n1, n2) not in l2s:
            nodes.append({"path": [n1, n2], "type": "category", "url": None, "code": None}); l2s.add((n1, n2))
        nodes.append({"path": [n1, n2, n3], "type": "category", "url": BASE + url,
                      "code": url.rsplit("/", 1)[-1]})

    # Accessories branch: L1 = Accessories, L2 = the 400_0xx pages themselves
    acc = extra_rows()
    print("accessories L2 entries:", len(acc))
    if acc:
        nodes.append({"path": ["Accessories"], "type": "category",
                      "url": BASE + "/us/en/category/400", "code": "400"})
    for n1, n2, _x, url in acc:
        nodes.append({"path": [n1, n2], "type": "category", "url": BASE + url,
                      "code": url.rsplit("/", 1)[-1]})

    r = Renderer()
    for i, (n1, n2, n3, url) in enumerate(rows, 1):
        t = r.html(BASE + url)
        if not t:
            continue
        heads = [(m.start(), txt(m.group(1))) for m in HEAD.finditer(t)]
        for m in B64.finditer(t):
            try:
                tiles = json.loads(base64.b64decode(m.group(1)))
            except Exception:
                continue
            if not isinstance(tiles, list) or not tiles or "categoryId" not in tiles[0]:
                continue
            prev = [h for p, h in heads if p < m.start() and h]
            group = prev[-1] if prev else n3
            nodes.append({"path": [n1, n2, n3, group], "type": "category",
                          "url": None, "code": tiles[0]["categoryId"].rsplit("_", 1)[0]})
            for tl in tiles:
                nodes.append({"path": [n1, n2, n3, group, tl.get("atKurzTextKachel")],
                              "type": "range", "url": BASE + tl.get("href", ""),
                              "code": tl.get("categoryId")})
        if i % 10 == 0:
            print("  L3 pages:", i, "/", len(rows), "nodes", len(nodes))

    for i, (n1, n2, _x, url) in enumerate(acc, 1):
        t = r.html(BASE + url)
        if not t:
            continue
        heads = [(m.start(), txt(m.group(1))) for m in HEAD.finditer(t)]
        for m in B64.finditer(t):
            try:
                tiles = json.loads(base64.b64decode(m.group(1)))
            except Exception:
                continue
            if not isinstance(tiles, list) or not tiles or "categoryId" not in tiles[0]:
                continue
            prev = [h for p, h in heads if p < m.start() and h]
            group = prev[-1] if prev else n2
            nodes.append({"path": [n1, n2, group], "type": "category", "url": None,
                          "code": tiles[0]["categoryId"].rsplit("_", 1)[0]})
            for tl in tiles:
                nodes.append({"path": [n1, n2, group, tl.get("atKurzTextKachel")],
                              "type": "range", "url": BASE + tl.get("href", ""),
                              "code": tl.get("categoryId")})
        print("  accessories pages:", i, "/", len(acc), "nodes", len(nodes))
    r.close()
    return nodes


if __name__ == "__main__":
    save("04 IFM", "https://www.ifm.com/us/en", crawl(),
         "L1-L3 từ navbar REST; L4/L5 từ trang danh mục render. Cấp cuối = product range ngay trên mã SP.")
