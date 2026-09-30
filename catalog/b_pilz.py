# -*- coding: utf-8 -*-
"""Pilz - www.pilz.com sits behind a Cloudflare challenge that rejects every
automated client, so the crawl uses Pilz's Chinese mirror www.pilz.com.cn, which
serves the identical catalogue in English under /en-CN/.
The eShop (/en-CN/eshop/.../c/<code>) carries the full category tree; the deepest
category is the level directly above the concrete order number."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

BASE = "https://www.pilz.com.cn"
SHOP = "/en-CN/eshop/"
LINK = re.compile(r'<a[^>]+href="(/en-CN/eshop/[^"?#]*/c/[^"?#/]+)"[^>]*>(.*?)</a>', re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def parts(href):
    p = href.split(SHOP, 1)[1]
    seg = p.split("/c/")[0]
    return [x for x in seg.split("/") if x], p.split("/c/")[1]


def crawl():
    root = try_fetch(BASE + SHOP + "Sensor-technology/c/A0100", retries=6, sleep=1.5, timeout=90)
    tops = {}
    for href, label in LINK.findall(root or ""):
        seg, code = parts(href)
        n = txt(label)
        if len(seg) == 1 and n and n.lower() != "downloads":
            tops.setdefault(code, (n, href))
    print("top categories:", len(tops))

    nodes, seen, queue = [], set(), []
    for code, (n, href) in sorted(tops.items()):
        nodes.append({"path": [n], "type": "category", "url": BASE + href, "code": code})
        queue.append((href, [n]))

    while queue:
        href, path = queue.pop(0)
        if href in seen:
            continue
        seen.add(href)
        t = try_fetch(BASE + href, retries=6, sleep=1.5, timeout=90)
        if not t:
            print("  [skip]", href)
            continue
        seg, _c = parts(href)
        kids = {}
        for h2, label in LINK.findall(t):
            s2, c2 = parts(h2)
            if len(s2) != len(seg) + 1 or s2[:len(seg)] != seg:
                continue
            n = txt(label)
            if n and n.lower() != "downloads":
                kids.setdefault(h2, n)
        for h2, n in sorted(kids.items()):
            s2, c2 = parts(h2)
            p = path + [n]
            nodes.append({"path": p, "type": "category", "url": BASE + h2, "code": c2})
            queue.append((h2, p))
        if len(seen) % 20 == 0:
            print("  pages:", len(seen), "nodes:", len(nodes), "queue:", len(queue))
    return nodes


if __name__ == "__main__":
    save("09 Pilz", "https://www.pilz.com/en-INT (crawl qua mirror www.pilz.com.cn/en-CN)", crawl(),
         "www.pilz.com chan bot (Cloudflare) - dung mirror .cn ban tieng Anh, eShop. Cap cuoi = cap ngay tren ma dat hang.")
