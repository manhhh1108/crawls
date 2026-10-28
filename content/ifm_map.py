# -*- coding: utf-8 -*-
"""IFM - lap ban do 'nhanh danh muc' -> ma danh muc tren ifm.com.

    python ifm_map.py

ifm.com chan moi request khong phai trinh duyet (403), nhung ben trong trang
co REST API sach se:

    /restservices/us/en/menu/navbar                 cay danh muc cap 1-3
    /restservices/us/en/assets/category/<id>        danh muc con cua <id>
    /restservices/us/en/category/<id>/content       headline + bulletPoints + text
    /restservices/us/en/category/<id>/productsAndAttributes   bang ma SP

Ma danh muc dang 200_010_010_010 (moi cap them 1 nhom 3 so). Script nay bo cay
tu navbar xuong het cac cap con roi doi chieu voi file danh muc Ntrend.

Ket qua: out/04 IFM/urlmap.json
"""
import io
import json
import os
import re
import sys
import collections

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

SITE = "https://www.ifm.com"
HOME = SITE + "/us/en"
REST = SITE + "/restservices/us/en"
OUT = os.path.join(HERE, "out", "04 IFM")
NTREND_XLSX = os.path.join(REPO, "data", "Danh mục SP web NTREND.xlsx")
SHEET = "04 IFM"

JS = """async (u) => { const r = await fetch(u, {credentials:'include'});
                      return {s: r.status, t: await r.text()}; }"""


def norm(s):
    return re.sub(r"[^0-9a-z]+", "", (s or "").lower())


class Rest(object):
    """Goi REST API tu BEN TRONG trang (fetch cua chinh trinh duyet) - goi tu
    ngoai, ke ca bang request context cua Playwright, deu bi tra 403."""

    def __init__(self, pw):
        self.b = pw.chromium.launch(headless=False, channel="chrome")
        self.pg = self.b.new_page()
        self.pg.goto(HOME, wait_until="domcontentloaded", timeout=120000)
        self.pg.wait_for_timeout(3000)

    def get(self, path):
        r = self.pg.evaluate(JS, REST + path)
        if r["s"] != 200:
            return None
        try:
            return json.loads(r["t"])
        except ValueError:
            return None

    def page(self, url):
        """HTML da render cua mot trang (dung khi REST khong co du lieu)."""
        self.pg.goto(url, wait_until="domcontentloaded", timeout=120000)
        self.pg.wait_for_timeout(1500)
        return self.pg.content()

    def fetch_bytes(self, url):
        """Tai file nhi phan (anh trong DAM). media.ifm.com tra 403 cho urllib
        nen phai dieu hinh mot tab phu sang thang URL do roi lay body."""
        pg = self.b.new_page()
        try:
            r = pg.goto(url, wait_until="commit", timeout=60000)
            return r.body() if r and r.ok else b""
        except Exception:
            return b""
        finally:
            pg.close()

    def close(self):
        self.b.close()


def walk_navbar(nav):
    """navbar JSON -> [(duong dan ten, ma danh muc)] cho cap 1-3."""
    out = []

    def rec(node, path):
        lab = node.get("label") or ""
        p = path + [lab] if lab and not lab.startswith("NAVIGATION_") else path
        u = node.get("url") or ""
        m = re.search(r"/category/([0-9_]+)", u)
        if m:
            out.append((p, m.group(1)))
        for k in ("products", "entries", "children"):
            for ch in node.get(k) or []:
                rec(ch, p)

    for n in nav:
        rec(n, [])
    return out


HREF_CAT = 'href="/us/en/category/'
TILE_TEXT = re.compile(r'ifm-category-item-tile__text[^>]*>(.*?)</div>', re.S)


def children(rest, code):
    """Danh muc con cua <code>.

    REST khong co endpoint liet ke danh muc con (/assets/category/<id> chi tra
    ve id anh dai dien). Nhung trang danh muc render san cac the con, moi the
    co href + nhan -> doc thang tu HTML.

    Cat theo href roi lay nhan DAU TIEN trong tung doan, khong dung mot regex
    ghep ca cap: the nao thieu nhan thi regex ghep se "an" sang nhan cua the
    sau, lam mat phan lon danh muc con (da gap: 30 the chi lay duoc 12).
    """
    html = rest.page(SITE + "/us/en/category/%s" % code)
    out, seen = [], set()
    for chunk in html.split(HREF_CAT)[1:]:
        cid = chunk.split('"', 1)[0]
        if not re.fullmatch(r"[0-9_]+", cid) or cid in seen:
            continue
        # Co trang bo QUA mot cap: /category/260_020_010 lien ket thang toi
        # 260_020_010_010_010 (sau 2 cap). Nhan moi hau due, dung doi hoi dung
        # cap ke tiep - neu doi thi nhanh do mat sach danh muc con.
        if not (cid.startswith(code + "_") and
                cid.count("_") > code.count("_")):
            continue
        m = TILE_TEXT.search(chunk[:4000])
        if not m:
            continue
        lab = re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", m.group(1))).strip()
        if lab:
            seen.add(cid)
            out.append((cid, lab))
    return out


def expand(rest, code, path, seen, out):
    """Bo xuong het cac cap con cua mot ma danh muc."""
    if code in seen:
        return
    seen.add(code)
    out.append((path, code))
    for cid, lab in children(rest, code):
        expand(rest, cid, path + [lab], seen, out)


BRAND_VI = "IFM"


def cat_of(v):
    """Cot danh muc Ntrend chi ghi moi ten hang = chua phan loai."""
    c = (v or "").strip()
    if c == BRAND_VI:
        return ("Không có danh mục sản phẩm,  chỉ nằm trong danh mục "
                "Thương Hiệu " + BRAND_VI)
    return c


def wanted():
    """668 dong SP (cot 'La cap ngay tren ma SP' = x)."""
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
    cache = os.path.join(OUT, "cats.json")
    if os.path.exists(cache) and "--refresh" not in sys.argv:
        cats = json.load(io.open(cache, encoding="utf-8"))
    else:
        from playwright.sync_api import sync_playwright
        with sync_playwright() as pw:
            rest = Rest(pw)
            nav = rest.get("/menu/navbar") or []
            tops = walk_navbar(nav)
            print("navbar: %d danh muc cap 1-3" % len(tops))
            seen, out = set(), []
            for i, (p, c) in enumerate(tops, 1):
                expand(rest, c, p, seen, out)
                if i % 20 == 0:
                    print("   %d/%d nhanh, da co %d danh muc"
                          % (i, len(tops), len(out)))
            rest.close()
        cats = [{"path": p, "code": c} for p, c in out]
        json.dump(cats, io.open(cache, "w", encoding="utf-8"), ensure_ascii=False)
    print("tong cong %d danh muc" % len(cats))

    by_name = collections.defaultdict(list)
    for c in cats:
        if c["path"]:
            by_name[norm(c["path"][-1])].append(c)

    lines = wanted()
    hit = 0
    for l in lines:
        cand = by_name.get(norm(l["name"])) or []
        if len(cand) > 1:                       # chon ban trung nhieu cap cha nhat
            par = {norm(x) for x in l["path"][:-1]}
            cand = sorted(cand, key=lambda c: -len(par & {norm(x) for x in c["path"]}))
        if cand:
            l["code"] = cand[0]["code"]
            l["url"] = "%s/us/en/category/%s" % (SITE, cand[0]["code"])
            hit += 1
    json.dump({"lines": lines}, io.open(os.path.join(OUT, "urlmap.json"), "w",
              encoding="utf-8"), ensure_ascii=False)
    print("dong SP: %d | khop: %d | chua khop: %d"
          % (len(lines), hit, len(lines) - hit))
    for l in lines[:400]:
        if not l.get("url"):
            print("   [?]", " > ".join(l["path"]))


if __name__ == "__main__":
    main()
