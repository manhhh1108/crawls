# -*- coding: utf-8 -*-
"""OTENNLUX Lighting - en.otennlux.com (EN).
The mega-menu carries the whole 2-level tree. Each category page lists its
products (_pNN.html) inside the cbp-vm-* grid; the same _pNN links also appear
in a global "hot products" block (pro_image containers) on every page, so only
links whose surrounding markup is the cbp grid are taken."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

CN = "https://cn.otennlux.com/"
EN = "https://en.otennlux.com/"
BLOCK = re.compile(r'<li class="col-sm-2">(.*?)</li>', re.S)
TITLE = re.compile(r'<a href="([^"]+)"[^>]*class="title"[^>]*>(.*?)</a>', re.S)
CHILD = re.compile(r'<a href="([^"]+)"[^>]*class="vgema-title"[^>]*>(.*?)</a>', re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def code(href):
    m = re.search(r"_(c\d+)$", href)
    return m.group(1) if m else href


def parse(home):
    t = fetch(home)
    tree = []
    for blk in BLOCK.findall(t):
        m = TITLE.search(blk)
        if not m:
            continue
        kids = [(code(h), txt(l), home.rstrip("/") + h) for h, l in CHILD.findall(blk)]
        tree.append((code(m.group(1)), txt(m.group(2)), home.rstrip("/") + m.group(1), kids))
    return tree


PROD = re.compile(r'<a[^>]+href="(/[^"]*_p(\d+)\.html)"[^>]*>(.*?)</a>', re.S)


def products(url):
    """_pNN links inside the cbp-vm grid (not the global hot-products block)."""
    t = try_fetch(url)
    if not t:
        return []
    best = {}
    for m in PROD.finditer(t):
        if "cbp-vm" not in t[max(0, m.start() - 1500):m.start()]:
            continue
        href, pid, label = m.groups()
        n = txt(label)
        if not n or n.lower() in ("view more", "more", "read more"):
            continue
        if len(n) > len(best.get(href, ("", ""))[0]):
            best[href] = (n, pid)
    return [{"name": n, "url": EN.rstrip("/") + h, "code": "p" + pid}
            for h, (n, pid) in best.items()]


def crawl():
    # crawl the English catalogue directly; its _cNN ids differ from the CN
    # version, so the trees cannot be merged - the EN site is the source.
    nodes, seen = [], set()
    for c, n, u, kids in parse(EN):
        if n not in seen:
            seen.add(n)
            nodes.append({"path": [n], "type": "category", "url": u, "code": c})
        for kc, kn, ku in kids:
            if (n, kn) in seen:
                continue
            seen.add((n, kn))
            nodes.append({"path": [n, kn], "type": "range", "url": ku, "code": kc})
    # products hang under the leaves: level-2 nodes, or level-1 without children
    parents = {tuple(x["path"][:-1]) for x in nodes if len(x["path"]) > 1}
    for nd in list(nodes):
        if tuple(nd["path"]) in parents:
            continue
        for pr in products(nd["url"]):
            k = tuple(nd["path"]) + (pr["name"],)
            if k in seen:
                continue
            seen.add(k)
            nodes.append({"path": list(k), "type": "product",
                          "url": pr["url"], "code": pr["code"]})
    # parent pages carry their own grid products that sit in no subcategory
    # (e.g. LED Work Light lists 12 lights while its 3 subcategories are empty)
    for nd in list(nodes):
        p = tuple(nd["path"])
        if p not in parents or nd.get("type") == "product":
            continue
        have = {x["url"] for x in nodes if x["type"] == "product"
                and tuple(x["path"][:len(p)]) == p}
        for pr in products(nd["url"]):
            k = p + (pr["name"],)
            if pr["url"] in have or k in seen:
                continue
            seen.add(k)
            nodes.append({"path": list(k), "type": "product",
                          "url": pr["url"], "code": pr["code"]})
    return nodes


if __name__ == "__main__":
    ns = crawl()
    print("  products:", sum(1 for x in ns if x["type"] == "product"))
    save("16 OTENNLUX Lighting", "https://en.otennlux.com/", ns,
         "Cây 2 cấp từ mega-menu bản EN + mã SP (_pNN, type=product) từ grid trên trang danh mục. "
         "Bản CN có thêm 5 dòng SP chưa đưa lên bản EN.")
