# -*- coding: utf-8 -*-
"""OTENNLUX Lighting - cn.otennlux.com (+ its English mirror en.otennlux.com).
The mega-menu carries the whole 2-level tree; product pages (_pNN.html) hang
directly under level 2. Category ids (_cNN) are shared between the two language
versions, so the label is "English (中文)" where both exist."""
import re, html as H
from fetcher import fetch
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
    cn, en = parse(CN), parse(EN)
    en_names = {}
    for c, n, _u, kids in en:
        en_names[c] = n
        for kc, kn, _ku in kids:
            en_names[kc] = kn

    def label(c, n):
        # the two language versions use different _cNN ids, so they cannot be
        # matched reliably - keep the (authoritative) Chinese catalogue names
        return n

    nodes, seen = [], set()
    for c, n, u, kids in cn:
        p1 = label(c, n)
        if p1 not in seen:
            seen.add(p1)
            nodes.append({"path": [p1], "type": "category", "url": u, "code": c})
        for kc, kn, ku in kids:
            n2 = label(kc, kn)
            if (p1, n2) in seen:
                continue
            seen.add((p1, n2))
            nodes.append({"path": [p1, n2], "type": "range", "url": ku, "code": kc})
    return nodes


if __name__ == "__main__":
    save("16 OTENNLUX Lighting", "https://cn.otennlux.com/ (+ en.otennlux.com)", crawl(),
         "Cay 2 cap tu mega-menu, ten tieng Trung theo web goc. Cap 2 la cap ngay tren trang SP cu the.")
