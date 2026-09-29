# -*- coding: utf-8 -*-
"""Tree Electric (Shanghai Quyi) - www.tree-electric.com.
Flat catalogue: 9 product families, models only in the downloadable catalogues."""
import re, html as H
from fetcher import fetch
from common import save

URL = "http://www.tree-electric.com/templates/products.html"


def crawl():
    t = fetch(URL)
    out, seen = [], set()
    for href, label in re.findall(r'<a[^>]+href="([^"]*products\.html\?index=\d+)"[^>]*>(.*?)</a>', t, re.S):
        n = re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", label))).strip()
        if not n or n in seen:
            continue
        seen.add(n)
        out.append({"path": [n], "type": "category",
                    "url": "http://www.tree-electric.com/templates/" + href.split("/")[-1],
                    "code": href.rsplit("=", 1)[-1]})
    # the side menu also carries families without an index link
    for label in re.findall(r'<a[^>]*href="javascript:;"[^>]*>(.*?)</a>', t, re.S):
        n = re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", label))).strip()
        if n and n not in seen:
            seen.add(n)
            out.append({"path": [n], "type": "category", "url": URL, "code": None})
    return out


if __name__ == "__main__":
    save("13 Tree Electric", "http://www.tree-electric.com/", crawl(),
         "Danh muc phang 1 cap (web khong co cap con; model nam trong catalogue PDF).")
