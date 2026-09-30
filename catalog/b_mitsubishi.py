# -*- coding: utf-8 -*-
"""Mitsubishi Electric FA - www.mitsubishielectric.com/fa.
L1..L3 from the site's own product mega-menu JSON; L4 = the product groups listed
on each series page (CPU module, I/O module, ...), i.e. the level directly above
the concrete model number."""
import re, json, html as H
from concurrent.futures import ThreadPoolExecutor
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

BASE = "https://www.mitsubishielectric.com"
MENU = BASE + "/fa/shared/gws0001/data/megadropdown_products.json"
A = re.compile(r'<a[^>]+href="([^"#?]+)"[^>]*>(.*?)</a>', re.S)
SKIP = ("more details", "details", "click here for details", "top")


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s.replace("<br>", " ")))).strip()


def flatten(node, path, out):
    label = txt(node.get("label") or "")
    p = path + [label] if label else path
    kids = node.get("child") or []
    if p and p != path:
        out.append((p, node.get("path")))
    for k in kids:
        flatten(k, p, out)


def crawl():
    root = json.loads(fetch(MENU, json_mode=True))
    flat = []
    for c in root.get("child") or []:
        flatten(c, [], flat)
    nodes, seen = [], set()
    leaves = []
    for p, path in flat:
        name = p[-1]
        if name.lower().endswith(" top"):
            continue
        t = tuple(p)
        if t in seen:
            continue
        seen.add(t)
        url = BASE + path if path and path != "#" else None
        nodes.append({"path": p, "type": "category", "url": url, "code": None})
        if url and len(p) >= 3:
            leaves.append((p, url))
    print("menu nodes:", len(nodes), "series pages:", len(leaves))

    def one(args):
        p, url = args
        return p, url, try_fetch(url, retries=2, sleep=0.3)

    with ThreadPoolExecutor(max_workers=5) as ex:
        res = list(ex.map(one, leaves))

    added = 0
    for p, url, t in res:
        if not t:
            continue
        branch = "/".join(url.split(BASE)[1].strip("/").split("/")[:4])
        for href, label in A.findall(t):
            if not href.startswith("/" + branch + "/"):
                continue
            if "/pmerit/" not in href and "/items/" not in href:
                continue
            name = txt(label)
            if not name or len(name) > 100 or name.lower() in SKIP:
                continue
            tt = tuple(p + [name])
            if tt in seen:
                continue
            seen.add(tt)
            nodes.append({"path": list(tt), "type": "range", "url": BASE + href, "code": None})
            added += 1
    print("  product groups added:", added)
    return nodes


if __name__ == "__main__":
    save("07 Mitsubishi Electric", "https://www.mitsubishielectric.com/fa/", crawl(),
         "L1-L3 tu megadropdown_products.json; L4 = nhom SP tren trang series (cap ngay tren ma model).")
