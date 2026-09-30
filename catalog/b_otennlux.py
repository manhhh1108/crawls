# -*- coding: utf-8 -*-
"""OTENNLUX Lighting - cn.otennlux.com (+ its English mirror en.otennlux.com).
The mega-menu carries the whole 2-level tree; product pages (_pNN.html) hang
directly under level 2. Category ids (_cNN) are shared between the two language
versions, so the label is "English (中文)" where both exist."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch
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
    return nodes


if __name__ == "__main__":
    save("16 OTENNLUX Lighting", "https://en.otennlux.com/", crawl(),
         "Cay 2 cap tu mega-menu ban EN. Cap 2 la cap ngay tren trang SP cu the. "
         "Ban CN co them 5 dong SP chua dua len ban EN.")
