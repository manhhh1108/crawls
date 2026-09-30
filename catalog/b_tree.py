# -*- coding: utf-8 -*-
"""Tree Electric (Shanghai Quyi) - www.tree-electric.com.
Flat catalogue: 9 product families, models only in the downloadable catalogues."""
import re, html as H
import os as _os, sys as _sys
_sys.path.insert(0, _os.path.dirname(_os.path.dirname(_os.path.abspath(__file__))))
from shared.fetcher import fetch
from common import save

URL = "http://www.tree-electric.com/templates/products.html"

# The site has no English version; translations follow the previous workbook,
# label style "English (中文)".
TRANS = {
    "辅材综合类": "Auxiliary Materials (General)",
    "接线端子系列": "Terminal Blocks Series",
    "重载接插件系列": "Heavy-duty Connectors Series",
    "固态继电器系列一": "Solid State Relays Series I",
    "软管/波纹管系列": "Hoses / Corrugated Tubing Series",
    "电缆接头系列": "Cable Glands Series",
    "风机过滤网系列": "Fan Filter Series",
    "线缆保护系列": "Cable Protection Series",
    "工业交换机系列": "Industrial Switches Series",
}


def bilabel(n):
    en = TRANS.get(n)
    if not en:
        print("  [warn] chua co ban dich cho:", n)
        return n
    return "%s (%s)" % (en, n)


def crawl():
    t = fetch(URL)
    out, seen = [], set()
    for href, label in re.findall(r'<a[^>]+href="([^"]*products\.html\?index=\d+)"[^>]*>(.*?)</a>', t, re.S):
        n = re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", label))).strip()
        if not n or n in seen:
            continue
        seen.add(n)
        out.append({"path": [bilabel(n)], "type": "category",
                    "url": "http://www.tree-electric.com/templates/" + href.split("/")[-1],
                    "code": href.rsplit("=", 1)[-1]})
    # the side menu also carries families without an index link
    for label in re.findall(r'<a[^>]*href="javascript:;"[^>]*>(.*?)</a>', t, re.S):
        n = re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", "", label))).strip()
        if n and n not in seen:
            seen.add(n)
            out.append({"path": [bilabel(n)], "type": "category", "url": URL, "code": None})
    return out


if __name__ == "__main__":
    save("13 Tree Electric", "http://www.tree-electric.com/", crawl(),
         "Danh mục phẳng 1 cấp (web không có cấp con; model nằm trong catalogue PDF). "
         "Web chỉ có tiếng Trung; tên EN dịch theo file cũ, dạng 'English (中文)'.")
