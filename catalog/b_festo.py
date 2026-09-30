# -*- coding: utf-8 -*-
"""Festo - www.festo.com (US/EN).
Category tree via /search/categories/{pimId} (subCategories), product families
(the level directly above the concrete order code) via .../products."""
import json, sys
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.browser import Browser
from common import save

BASE = "https://www.festo.com/us/en"
ROOT = "pim1"


def crawl():
    b = Browser(BASE)
    nodes, seen = [], set()

    def subcats(pid):
        d = b.api_json("%s/search/categories/%s" % (BASE, pid))
        return d.get("subCategories") or []

    def products(pid):
        out, page, pages = [], 0, 1
        while page < pages:
            u = "%s/search/categories/%s/products?page=%d" % (BASE, pid, page)
            d = b.api_json(u)
            pg = d.get("pagination") or {}
            pages = pg.get("totalPageCount") or 1
            out.extend(d.get("productList") or [])
            page += 1
            if page > 200:
                break
        return out

    # ---- pass 1: full category tree
    stack, allcats = [(ROOT, [])], []
    n_cat = 0
    while stack:
        pid, path = stack.pop()
        if pid in seen:
            continue
        seen.add(pid)
        for k in subcats(pid):
            p = path + [k["name"]]
            nodes.append({"path": p, "type": "category", "url": k.get("url"), "code": k["pimId"]})
            allcats.append((k["pimId"], p))
            stack.append((k["pimId"], p))
            n_cat += 1
            if n_cat % 50 == 0:
                print("  categories:", n_cat, "queue", len(stack))
    print("  total categories:", n_cat)

    # ---- pass 2: product families. Query EVERY category (a family can hang on a
    # node that also has sub-categories) and place it by its own breadcrumb.
    catpaths = {c: p for c, p in allcats}
    done = 0
    for pid, path in allcats:
        for pr in products(pid):
            nm = (pr.get("name") or "").strip()
            sc = (pr.get("shortCode") or "").strip()
            label = ("%s %s" % (nm, sc)).strip() if sc and sc not in nm else nm
            bc = pr.get("breadcrumbValues") or []
            if bc and bc[0].lower() == "products":
                bc = bc[1:]
            nodes.append({"path": (bc or path) + [label], "type": "family",
                          "url": pr.get("url"), "code": sc or pr.get("code")})
        done += 1
        if done % 25 == 0:
            print("  product lists:", done, "/", len(allcats))
    b.close()
    return nodes


if __name__ == "__main__":
    ns = crawl()
    save("10 Festo", "https://www.festo.com/us/en", ns,
         "Cay danh muc lay tu API /search/categories; cap cuoi = dong san pham (product family) ngay tren ma SP cu the.")
