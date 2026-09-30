# -*- coding: utf-8 -*-
"""Pepperl+Fuchs - www.pepperl-fuchs.com (en-sg).

Level 1/2 come from the site navigation (rendered once); deeper levels are the
"action cards" on each category page, which are built client-side - and the site
rejects both curl and headless Chromium for that markup, so a real (headed)
Chrome does the rendering. The deepest category is the level directly above the
concrete model designation."""
import os, re, gzip, time, hashlib, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch, try_fetch
from common import save

SITE = "https://www.pepperl-fuchs.com"
HUB = SITE + "/en-sg/products-gp25581"
SEED = SITE + "/en-sg/products/industrial-sensors-gp27354"
CACHE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), ".cache", "pepperl")
os.makedirs(CACHE, exist_ok=True)

CARDLINK = re.compile(r'<a class="action-card[^"]*"[^>]*href="([^"]+)"', re.S)
CLASSQ = re.compile(r"\?class=(\d+)$")
H1 = re.compile(r"<h1[^>]*>(.*?)</h1>", re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def segs(u):
    if "/en-sg/products/" not in u:
        return []
    return [x for x in re.sub(r"-gp\d+$", "", u.split("/en-sg/products", 1)[1]).split("/") if x]


class Chrome(object):
    def __init__(self):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self.b = self._pw.chromium.launch(headless=False, channel="chrome")
        self.pg = self.b.new_context(locale="en-US", viewport={"width": 1500, "height": 1000}).new_page()

    def html(self, url):
        p = os.path.join(CACHE, hashlib.sha1(url.encode()).hexdigest() + ".gz")
        if os.path.exists(p):
            with gzip.open(p, "rt", encoding="utf-8", errors="replace") as f:
                return f.read()
        for i in range(4):
            try:
                self.pg.goto(url, wait_until="domcontentloaded", timeout=90000)
                self.pg.wait_for_timeout(2000)
                t = self.pg.content()
                if len(t) > 50000:
                    with gzip.open(p, "wt", encoding="utf-8") as f:
                        f.write(t)
                    return t
            except Exception:
                pass
            time.sleep(3 * (i + 1))
        print("  [render fail]", url)
        return None

    def close(self):
        try:
            self.b.close()
        finally:
            self._pw.stop()


def nav_tree(t):
    out = []
    for seg in re.split(r'<li class="navigation__child-list-item--parent navigation__child-list-item">', t)[1:]:
        seg = seg[:40000]
        m = re.search(r'class="navigation__child-list-item-link"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', seg, re.S)
        if not m:
            continue
        subs = [(txt(b), a) for a, b in
                re.findall(r'class="navigation__subchild-list-item-link"[^>]*href="([^"]+)"[^>]*>(.*?)</a>', seg, re.S)]
        subs = [(n, a) for n, a in subs if "?class=" in a and n]
        if subs:
            out.append((txt(m.group(2)), subs))
    return out


def class_to_category(cid):
    """Resolve a ?class= filter to its canonical category page, if it has one."""
    t = try_fetch(HUB + "?class=" + cid, retries=5, sleep=2.5, timeout=90)
    if not t:
        return None
    for href in CARDLINK.findall(t):
        if href.startswith("/en-sg/products/") and len(segs(SITE + href)) >= 2:
            return SITE + href
    return None


def children_of(t, u):
    """Sub-categories shown on a rendered category page."""
    s = segs(u)
    out, seen = [], set()
    for m in re.finditer(r'<div class="action-card-container__col">', t):
        blk = t[m.start():m.start() + 14000]
        h = re.search(r'class="action-card__headline">(.*?)</h2>', blk, re.S)
        if not h:
            continue
        name = txt(h.group(1))
        deeper, cid = None, None
        for href in re.findall(r'href="([^"]+)"', blk):
            hs = segs(SITE + href) if href.startswith("/en-sg/products/") else []
            if len(hs) == len(s) + 1 and hs[:len(s)] == s:
                deeper = SITE + href
            mm = CLASSQ.search(href)
            if mm and "products-gp" in href:
                cid = mm.group(1)
        if not name or (not deeper and not cid):
            continue
        if name in seen:
            continue
        seen.add(name)
        out.append((name, deeper, cid))
    return out


def crawl():
    ch = Chrome()
    seed = ch.html(SEED)
    tree = nav_tree(seed or "")
    print("divisions:", len(tree), "L2:", sum(len(s) for _d, s in tree))

    nodes = []
    stack = []
    for dname, subs in tree:
        nodes.append({"path": [dname], "type": "category", "url": None, "code": None})
        for n2, href in subs:
            cid = CLASSQ.search(href).group(1)
            cat = class_to_category(cid)
            nodes.append({"path": [dname, n2], "type": "category",
                          "url": cat or (HUB + "?class=" + cid), "code": cid})
            # some level-2 entries have no category page of their own - their
            # sub-groups are only on the class-filtered hub page
            stack.append((cat or (HUB + "?class=" + cid), [dname, n2]))
    print("level-2 queued:", len(stack))

    done = 0
    while stack:
        u, path = stack.pop(0)
        t = ch.html(u)
        done += 1
        if not t:
            continue
        for name, deeper, cid in children_of(t, u):
            p = path + [name]
            nodes.append({"path": p, "type": "category" if deeper else "range",
                          "url": deeper or (HUB + "?class=" + cid), "code": cid})
            if deeper and len(p) < 5:
                stack.append((deeper, p))
        if done % 10 == 0:
            print("  rendered:", done, "nodes:", len(nodes), "queue:", len(stack))
    ch.close()
    return nodes


if __name__ == "__main__":
    save("02 Pepperl+Fuchs", "https://www.pepperl-fuchs.com/en-sg", crawl(),
         "L1/L2 tu navigation; cap sau lay tu the danh muc tren trang (render bang Chrome that). Cap cuoi = cap ngay tren ma SP.")
