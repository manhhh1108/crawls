# -*- coding: utf-8 -*-
"""Guangdong Xingguang / Starshine Drive - www.xgcd.cn.
Level 1 = gearbox families (/product/<id>/), level 2 = the series pages under
them (/product/<id>.html), which sit directly above the concrete model."""
import re, html as H
from fetcher import fetch, try_fetch
from common import save

ROOT = "https://www.xgcd.cn/product/5/"
BASE = "https://www.xgcd.cn"
A = re.compile(r'<a[^>]+href="(/product/[^"]+)"[^>]*>(.*?)</a>', re.S)


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", s))).strip()


def crawl():
    t = try_fetch(ROOT, retries=4)
    if not t:
        return []
    cats, seen = [], set()
    for href, label in A.findall(t):
        n = txt(label)
        if re.match(r"^/product/\d+/$", href) and n and n != "产品中心" and n not in seen:
            seen.add(n)
            cats.append((n, href))
    nodes = []
    for n, href in cats:
        nodes.append({"path": [n], "type": "category", "url": BASE + href, "code": href.strip("/").split("/")[-1]})
        ct = try_fetch(BASE + href, retries=4)
        if not ct:
            continue
        for h2, l2 in A.findall(ct):
            n2 = txt(l2)
            if not re.match(r"^/product/[\d-]+\.html$", h2) or not n2:
                continue
            k = (n, n2)
            if k in seen:
                continue
            seen.add(k)
            nodes.append({"path": [n, n2], "type": "range", "url": BASE + h2, "code": None})
    return nodes


if __name__ == "__main__":
    save("14 Xingguang-Starshine", "https://www.xgcd.cn/", crawl(),
         "Cay 2 cap tu trang /product/; cap 2 la dong SP ngay tren ma cu the.")
