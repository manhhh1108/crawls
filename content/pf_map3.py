# -*- coding: utf-8 -*-
"""Pepperl+Fuchs - do BO SUNG map theo dung nhanh cha (buoc 3).

pf_map2.py gop 3 nguon san co va khop (cha, ten). Nhung van con:
  - 106 dong chua co trang
  - 38 dong chi khop duoc TEN, khong khop nhanh -> de tro nham trang
    (vd 'Accessories' duoi 18 nhanh, 8 dong tro ve cung 1 trang)

Buoc nay di dung dia chi: voi moi dong con thieu/khong chac, mo TRANG CHA cua
no (lay tu chinh map da co) roi doc cac the con tren do. Nho vay (cha, ten)
luon dung nhanh.

    python pf_map3.py            # do nhung dong chua chac
    python pf_map3.py --all      # do lai tat ca

Ket qua: ghi them vao out/02 Pepperl+Fuchs/classmap2.json
"""
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

import pf_map2 as M                                   # noqa: E402
from pf_map2 import CARD, HEAD, key, txt, full, MAP2  # noqa: E402


def parent_page(m, ln):
    """URL trang cha cua dong SP (de doc the con tren do)."""
    path = ln["path"]
    nm = ln["name"]
    idx = len(path) - 1
    while idx >= 0 and path[idx].lower() != nm.lower():
        idx -= 1
    pars = [path[idx - 1]] if idx > 0 else []
    pars += [p for p in reversed(path[:-1]) if p not in pars]
    for i, par in enumerate(pars):
        gp = pars[i + 1] if i + 1 < len(pars) else ""
        hit = m.get(key(gp, par)) or next(
            (v for k, v in m.items() if k.endswith("||" + par.lower())), None)
        if hit:
            return hit["url"], par
    return "", ""


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    redo_all = "--all" in sys.argv
    m = json.load(io.open(MAP2, encoding="utf-8"))
    lines = M.lines()

    todo = []
    for ln in lines:
        hit, how = M.lookup(m, ln)
        if redo_all or not hit or how == "chi khop ten":
            todo.append(ln)
    # gom theo trang cha de moi trang chi mo 1 lan
    by_page = {}
    nopar = []
    for ln in todo:
        u, par = parent_page(m, ln)
        if not u:
            nopar.append(ln["name"])
            continue
        by_page.setdefault(u, {"parent": par, "lines": []})["lines"].append(ln)
    print("dong can do lai: %d | trang cha phai mo: %d | khong ro cha: %d"
          % (len(todo), len(by_page), len(nopar)))
    if not by_page:
        return

    from playwright.sync_api import sync_playwright
    from c_pepperl_try import accept_cookies
    added = 0
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, channel="chrome")
        pg = b.new_page()
        first = True
        for i, (u, info) in enumerate(by_page.items(), 1):
            try:
                pg.goto(u, wait_until="networkidle", timeout=90000)
                pg.wait_for_timeout(1400)
                if first:
                    accept_cookies(pg)
                    first = False
                html = pg.content()
            except Exception as e:
                print("  %3d/%d [loi] %s %s" % (i, len(by_page), u[-40:], str(e)[:40]))
                continue
            par = M.page_parent(html) or info["parent"]
            n = 0
            for href, blk in CARD.findall(html):
                mm = HEAD.search(blk)
                if not mm:
                    continue
                nm = txt(mm.group(1))
                if not nm:
                    continue
                k = key(par, nm)
                if k not in m:
                    m[k] = {"name": nm, "parent": par,
                            "url": full(href.replace("&amp;", "&"))}
                    n += 1
            added += n
            print("  %3d/%d %-34s +%d the (cha=%s)"
                  % (i, len(by_page), info["lines"][0]["name"][:34], n, par[:22]))
            if i % 10 == 0:
                json.dump(m, io.open(MAP2, "w", encoding="utf-8"),
                          ensure_ascii=False, indent=1)
        b.close()
    json.dump(m, io.open(MAP2, "w", encoding="utf-8"), ensure_ascii=False, indent=1)

    ok = exact = loose = 0
    miss = []
    for ln in lines:
        hit, how = M.lookup(m, ln)
        if not hit:
            miss.append(ln["name"])
            continue
        ok += 1
        if how == "chi khop ten":
            loose += 1
        else:
            exact += 1
    print("\nthem %d muc | dong SP co trang: %d/%d (khop nhanh %d, chi ten %d) | thieu %d"
          % (added, ok, len(lines), exact, loose, len(miss)))


if __name__ == "__main__":
    main()
