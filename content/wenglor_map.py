# -*- coding: utf-8 -*-
"""wenglor - lap ban do 'nhanh danh muc' -> URL trang danh sach SP.

    python wenglor_map.py

Cay danh muc trong 'Danh mục SP web NTREND.xlsx' duoc dung tu sitemap en cua
wenglor, nen duong dan trong URL khop 1-1 voi cac cot Cap 1..5. Khoa ban do la
chuoi duong dan da chuan hoa (bo dau gach, thuong hoa).

Ket qua: out/06 wenglor/urlmap.json
"""
import io
import json
import os
import re
import sys
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT = os.path.join(HERE, "out", "06 wenglor")
SITEMAP = "https://www.wenglor.com/sitemap.xml"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/129.0 Safari/537.36")
NTREND_XLSX = os.path.join(REPO, "data", "Danh mục SP web NTREND.xlsx")
SHEET = "06 wenglor"


def get(url):
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    return urllib.request.urlopen(req, timeout=90).read().decode("utf-8", "replace")


def key(parts):
    """['Sensors','Photoelectronic Sensors'] -> 'sensors/photoelectronic sensors'."""
    # bo HET ky tu khong phai chu/so: ten danh muc co '1D/2D' con URL viet
    # '1D2D', co ten co dau '-' con URL khong.
    out = []
    for p in parts:
        p = re.sub(r"[^0-9a-zA-Z]+", "", str(p or "")).lower()
        if p:
            out.append(p)
    return "/".join(out)


def url_key(u):
    """URL -> cung dang khoa (bo /en/ va duoi /c/cxmCID...)."""
    path = u.split("wenglor.com/en/", 1)[-1].split("/c/")[0]
    return key(path.split("/"))


def build():
    idx = get(SITEMAP)
    locs = re.findall(r"<loc>(.*?)</loc>", idx, re.S)
    urls = set()
    for sm in locs:
        sm = sm.strip()
        if "-en-" not in sm or not re.search(r"/(Category|CategoryLanding)-en-", sm):
            continue
        for u in re.findall(r"<loc>(.*?)</loc>", get(sm), re.S):
            u = u.strip()
            if "/c/cxmCID" in u and "/en/" in u:
                urls.add(re.sub(r"//c/", "/c/", u))
    return {url_key(u): u for u in sorted(urls)}


BRAND_VI = "wenglor"


def cat_of(v):
    """Cot danh muc Ntrend chi ghi moi ten hang = chua phan loai."""
    c = (v or "").strip()
    if c == BRAND_VI:
        return ("Không có danh mục sản phẩm,  chỉ nằm trong danh mục "
                "Thương Hiệu " + BRAND_VI)
    return c


def wanted():
    """142 dong SP (cot 'La cap ngay tren ma SP' = x)."""
    from openpyxl import load_workbook
    ws = load_workbook(NTREND_XLSX, read_only=True, data_only=True)[SHEET]
    out = []
    for r in ws.iter_rows(min_row=6, values_only=True):
        if not r or not r[0] or len(r) < 10 or str(r[8] or "").strip() != "x":
            continue
        path = [str(x).strip() for x in r[3:8] if x and str(x).strip()]
        if not path:
            continue
        out.append({"stt": str(r[0]).strip(), "name": path[-1], "path": path,
                    "cat_vi": cat_of(r[9])})
    return out


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    os.makedirs(OUT, exist_ok=True)
    m = build()
    print("sitemap: %d trang danh muc" % len(m))
    lines = wanted()
    # ten danh muc doi khi bi cat ngan so voi URL ("MLSLxxxS40" <-> trang
    # "MLSLxxxS40-for-reliable-laser-shutdown") -> doi chieu them theo tien to.
    hit, miss = [], []
    for l in lines:
        k = key(l["path"])
        u = m.get(k)
        if not u:
            cand = [v for kk, v in m.items() if kk.startswith(k)]
            if len(cand) == 1:
                u = cand[0]
        if u:
            l["url"] = u
            hit.append(l)
        else:
            miss.append(l)
    with io.open(os.path.join(OUT, "urlmap.json"), "w", encoding="utf-8") as f:
        json.dump({"map": m, "lines": lines}, f, ensure_ascii=False)
    print("dong SP: %d | khop: %d | chua khop: %d" % (len(lines), len(hit), len(miss)))
    for l in miss[:15]:
        print("   [?]", " > ".join(l["path"]))


if __name__ == "__main__":
    main()
