# -*- coding: utf-8 -*-
"""Supmea Automation - www.supmea.com.
The product listing page carries the full 2-level tree (<li> per level-1 with its
level-2 links). Concrete models (product_con_*.html) hang under level 2."""
import re, html as H
from fetcher import fetch
from common import save

URL = "https://www.supmea.com/product_25.html"
LI = re.compile(r'<li class="[^"]*">\s*<div class="up alltime">(.*?)</li>', re.S)
NAME = re.compile(r'<div class="text f_18">(.*?)</div>', re.S)
SUB = re.compile(r'<a href="(product_\d+_1\.html)[^"]*"[^>]*class="down_a[^"]*"[^>]*>(.*?)</a>', re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip().lstrip("·")


def crawl():
    t = fetch(URL)
    nodes, seen = [], set()
    for blk in LI.findall(t):
        m = NAME.search(blk)
        if not m:
            continue
        p1 = txt(m.group(1))
        if p1 and p1 not in seen:
            seen.add(p1)
            nodes.append({"path": [p1], "type": "category", "url": None, "code": None})
        for href, label in SUB.findall(blk):
            n = txt(label)
            k = (p1, n)
            if not n or k in seen:
                continue
            seen.add(k)
            nodes.append({"path": [p1, n], "type": "range",
                          "url": "https://www.supmea.com/" + href, "code": href})
    return nodes


if __name__ == "__main__":
    save("12 Supmea", "https://www.supmea.com/", crawl(),
         "Cay 2 cap tu trang danh muc SP; cap 2 la cap ngay tren ma SP cu the.")
