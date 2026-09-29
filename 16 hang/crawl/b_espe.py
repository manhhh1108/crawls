# -*- coding: utf-8 -*-
"""ESPE Technology - www.espetech.com.
Level 1 = the product section folders, level 2 = the product family pages inside
them (the level directly above the concrete model)."""
import re, html as H
from fetcher import fetch, try_fetch
from common import save

BASE = "https://www.espetech.com/"
SECT = {
    "safety-light-curtains": "Safety Light Curtains",
    "safety-door-switch": "Safety Door Switches",
    "measuring-light-curtains": "Measuring Light Curtains",
    "led-tower-light": "LED Tower Lights",
    "laser-scanner": "Laser Scanners",
    "accessory": "Accessories",
}
A = re.compile(r'<a[^>]+href="([^"]+\.html)"[^>]*>(.*?)</a>', re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def crawl():
    nodes, seen = [], set()
    for k, v in SECT.items():
        nodes.append({"path": [v], "type": "category", "url": BASE + k + "/", "code": k})
    pages = [BASE + "products/"] + [BASE + "products/index_%d.html" % i for i in range(2, 16)]
    for p in pages:
        t = try_fetch(p, retries=2)
        if not t:
            continue
        for href, label in A.findall(t):
            m = re.match(r"^(?:\.\./)?([a-z-]+)/(.+)\.html$", href.lstrip("/"))
            if not m or m.group(1) not in SECT:
                continue
            n = txt(label)
            key = (SECT[m.group(1)], n)
            if not n or key in seen:
                continue
            seen.add(key)
            nodes.append({"path": [SECT[m.group(1)], n], "type": "range",
                          "url": BASE + href.lstrip("/"), "code": None})
    return nodes


if __name__ == "__main__":
    save("15 ESPE Technology", "https://www.espetech.com/", crawl(),
         "Cay 2 cap tu trang /products (co phan trang); cap 2 la dong SP ngay tren ma cu the.")
