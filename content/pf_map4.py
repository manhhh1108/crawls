# -*- coding: utf-8 -*-
"""Pepperl+Fuchs - boc NGUYEN cay danh muc + ma class tu payload bo loc.

Phat hien 20/10: HTML cac trang P+F nhung san mot payload bo loc, trong do moi
gia tri la DUONG DAN day du tu goc xuong, dang:

    Order:000032768‖Level:2‖LiDAR Sensors‖30654‖
    ¦Order:000032768‖Level:3‖2-D LiDAR Sensors‖31266‖

Nghia la ca cay danh muc + ma class cua TUNG cap deu nam san trong HTML da tai.
Khong can crawl them trang nao: chi can doc .cache/pepperl/*.gz.

Nho vay map khong con nham khi ten bi lap (Accessories xuat hien 18 lan duoi 18
nhanh -> moi nhanh mot ma class rieng).

Ket qua: out/02 Pepperl+Fuchs/classmap2.json  (ghi de, dung chung khoa 'cha||ten')
"""
import gzip
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

import pf_map2 as M                      # noqa: E402
from pf_map2 import key, MAP2             # noqa: E402

CACHE = os.path.join(REPO, ".cache", "pepperl")
SITE = "https://www.pepperl-fuchs.com"
LIST = SITE + "/en-sg/products-gp25581?class=%s"

# Level:<n>‖<ten>‖<ma>‖   - dau phan cach la U+2016 (‖)
SEG = re.compile(r"Level:(\d+)‖([^‖]+)‖(\d+)‖")


def paths_from(html):
    """Moi chuoi gia tri bo loc -> [(level, ten, ma), ...] theo thu tu goc->la."""
    out = []
    for raw in re.findall(r'value:"([^"]*Level:[^"]*)"', html):
        segs = SEG.findall(raw)
        if segs:
            out.append([(int(a), b.strip(), c) for a, b, c in segs])
    return out


def build():
    m = {}
    if not os.path.isdir(CACHE):
        return m
    for f in sorted(os.listdir(CACHE)):
        if not f.endswith(".gz"):
            continue
        try:
            html = gzip.open(os.path.join(CACHE, f), "rt",
                             encoding="utf-8", errors="replace").read()
        except Exception:
            continue
        for segs in paths_from(html):
            for i, (lv, name, code) in enumerate(segs):
                parent = segs[i - 1][1] if i > 0 else ""
                k = key(parent, name)
                if k not in m:
                    m[k] = {"name": name, "parent": parent, "level": lv,
                            "code": code, "url": LIST % code}
    return m


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    m = build()
    print("boc tu cache: %d muc (cha||ten -> ma class)" % len(m))
    if not m:
        print("khong doc duoc cache"); return

    # giu lai cac muc cu khong co trong payload (vd trang /products/ cap 1-2)
    old = {}
    if os.path.exists(MAP2):
        old = json.load(io.open(MAP2, encoding="utf-8"))
    merged = dict(old)
    merged.update(m)                      # payload chinh xac hon -> de len tren
    json.dump(merged, io.open(MAP2, "w", encoding="utf-8"),
              ensure_ascii=False, indent=1)

    lines = M.lines()
    ok = exact = loose = 0
    miss = []
    for ln in lines:
        hit, how = M.lookup(merged, ln)
        if not hit:
            miss.append(ln["name"])
            continue
        ok += 1
        if how == "chi khop ten":
            loose += 1
        else:
            exact += 1
    print("dong SP: %d | co trang: %d (khop ca nhanh: %d, chi khop ten: %d) | thieu: %d"
          % (len(lines), ok, exact, loose, len(miss)))
    if miss:
        print("con thieu:", ", ".join(sorted(set(miss))[:15]),
              "..." if len(set(miss)) > 15 else "")
    print("->", MAP2)


if __name__ == "__main__":
    main()
