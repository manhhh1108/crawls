# -*- coding: utf-8 -*-
"""Bung cac dong trong urlmap Festo tro vao TRANG DANH MUC thanh dong SP con.

    python festo_cat_expand.py           # chi xem, khong ghi
    python festo_cat_expand.py --apply   # ghi vao urlmap.json

Chay sau `cat_map.py festo`, truoc `c_festo.py`. Chay lai duoc nhieu lan,
khong sinh trung (bo qua URL da co trong map).

Vi sao can: 10 dong trong urlmap tro vao `/c/...-id_pimNNN` (trang danh muc)
chu khong phai `/p/...` (trang dong SP). Trang `/c/` khong co `_ROUTER_DATA`
nen c_festo.py bao "noi dung khong hop le" va bo ca 10 dong. Trong cay
catalog/out/10 Festo.json ca 10 node nay deu type=category va KHONG CO CON
(catalog/b_festo.py dung lai o do) -> khong bung thi mat han so SP ben trong.

Lay con o dau: trang `/c/` khong de link SP trong HTML (tile la div, khong
phai <a>), nhung no goi mot API JSON tra ve dung danh sach kem URL:

    GET /us/en/search/categories/<pimId>/products?page=N
        productList[].name           ten SP (chua co ma series)
        productList[].url            URL trang dong SP /p/...
        pagination.totalPageCount    so trang

Ten dong ghi vao map chi la ten tam: c_festo.py se lay ten that tu
productSeriesData tren trang dong SP.
"""
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

URLMAP = os.path.join(HERE, "out", "10 Festo", "urlmap.json")
BASE = "https://www.festo.com/us/en"
API = BASE + "/search/categories/%s/products?page=%d"


def kids(br, pim):
    """Danh sach (ten, url) cac dong SP trong mot danh muc."""
    out, page, pages = [], 0, 1
    while page < pages:
        d = json.loads(br.api(API % (pim, page)))
        pages = (d.get("pagination") or {}).get("totalPageCount") or 1
        for p in d.get("productList") or []:
            u = p.get("url")
            if u:
                out.append((p.get("name") or "", u))
        page += 1
    return out


def code_of(url):
    """Ma dong SP nam trong URL: /p/<slug>-id_<MA>?... ."""
    m = re.search(r"-id_([^/?]+)", url or "")
    return m.group(1) if m else ""


def ident(url):
    """URL bo phan `?categoryCode=...`.

    Mot dong SP nam trong nhieu danh muc thi moi danh muc tra ve mot URL khac
    nhau o doan nay, nhung van la MOT dong SP. So nguyen ca URL thi lot trung
    (vd "Dosing elements FP_03-50-06-14" nam ca o pim525 va pim703) va se ra
    hai bai giong nhau.
    """
    return (url or "").split("?")[0]


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    apply = "--apply" in sys.argv
    doc = json.load(io.open(URLMAP, encoding="utf-8"))
    lines = doc["lines"]
    have = {ident(l["url"]) for l in lines if l.get("url")}
    cats = [l for l in lines if "/c/" in (l.get("url") or "")]
    if not cats:
        print("khong con dong nao tro vao trang danh muc")
        return

    from shared.browser import Browser
    br = Browser(BASE)
    new, skip = [], 0
    try:
        for c in cats:
            pim = code_of(c["url"])
            try:
                ks = kids(br, pim)
            except Exception as e:
                print("%-42s LOI %s" % (c["name"][:42], repr(e)[-60:]))
                continue
            print("%-42s %-9s con=%d" % (c["name"][:42], pim, len(ks)))
            for j, (nm, u) in enumerate(ks, 1):
                tag = ("%s %s" % (nm, code_of(u))).strip()
                if ident(u) in have:
                    skip += 1
                    print("      - %-46s (da co trong map)" % tag[:46])
                    continue
                have.add(ident(u))
                new.append(dict(
                    stt="%s.%d" % (c.get("stt", ""), j),
                    name=tag, path=(c.get("path") or []) + [tag],
                    cat_vi=c.get("cat_vi", ""), url=u,
                    code=code_of(u), kind="family"))
                print("      + %-46s %s" % (tag[:46], u.split("/p/")[-1][:40]))
    finally:
        br.close()

    print("\nthem %d dong moi | bo qua %d dong da co | go %d dong danh muc"
          % (len(new), skip, len(cats)))
    if not apply:
        print("chi xem - them --apply de ghi vao urlmap.json")
        return
    # Dat dong con vao dung cho cua dong danh muc, roi go dong danh muc di.
    out = []
    for l in lines:
        if l is not None and "/c/" in (l.get("url") or ""):
            out.extend(n for n in new
                       if n["stt"].rsplit(".", 1)[0] == l.get("stt", ""))
            continue
        out.append(l)
    doc["lines"] = out
    with io.open(URLMAP, "w", encoding="utf-8") as f:
        json.dump(doc, f, ensure_ascii=False)
    print("da ghi %s (%d -> %d dong)" % (URLMAP, len(lines), len(out)))


if __name__ == "__main__":
    main()
