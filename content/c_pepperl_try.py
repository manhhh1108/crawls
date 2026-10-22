# -*- coding: utf-8 -*-
"""Pepperl+Fuchs - BAN THU 1 DONG SAN PHAM (de Ntrend soat output truoc).

    python c_pepperl_try.py                      # mac dinh: Capacitive Sensors
    python c_pepperl_try.py 30710 "Inductive Sensors"

Khac SICK: pepperl-fuchs.com KHONG co API, va chan curl (tra ve HTML rong san
pham). Phai render bang Chrome that qua Playwright.

Nguon du lieu cho tung truong cua template:
    Bang ma SP + thong tin  <- trang ?class=<id>  (ma SP | ten | mo ta)
    Mo ta dai (5-8 y)       <- cac cap "Key: Value" xuat hien nhieu nhat trong
                               mo ta cua ca dong (P+F khong co Features/Benefits)
    Mo ta ngan (overview)   <- doan van tren trang danh muc cha
    Danh muc / ten tieng Viet <- data/Danh muc SP web NTREND.xlsx sheet 02
"""
import collections
import io
import os
import re
import sys
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

SITE = "https://www.pepperl-fuchs.com"
LIST_URL = SITE + "/en-sg/products-gp25581?class=%s"
NTREND_XLSX = os.path.join(REPO, "data", "Danh mục SP web NTREND.xlsx")
OUT_DIR = os.path.join(REPO, "data", "Noi dung Pepperl+Fuchs")

BRAND = "Pepperl+Fuchs"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")
MIN_HL, MAX_HL = 5, 8          # template: 5-8 thong so/dac diem


def txt(s):
    return H.unescape(re.sub(r"\s+", " ", re.sub(r"<[^>]+>", " ", s or ""))).strip()


# ---------------------------------------------------------------------------
# 1. Tai trang (Chrome that)
# ---------------------------------------------------------------------------
def accept_cookies(pg):
    """Banner cookie cua P+F phu len trang -> khong bam duoc 'Load more'."""
    for t in ("Accept all", "Allow all", "I agree", "Accept", "Agree", "Got it"):
        try:
            loc = pg.get_by_role("button", name=re.compile(t, re.I))
            if loc.count() and loc.first.is_visible():
                loc.first.click(timeout=5000)
                pg.wait_for_timeout(1200)
                return True
        except Exception:
            pass
    return False


def load_all(pg, max_clicks=80):
    """Bam 'Load more' den khi khong ra them ma nao (trang chi hien 50 ma dau)."""
    last = pg.content().count('class="product-card"')
    same = 0
    for _ in range(max_clicks):
        try:
            loc = pg.locator("[class*=loadmore]")
            if loc.count() == 0 or not loc.first.is_visible():
                break
            loc.first.scroll_into_view_if_needed()
            pg.wait_for_timeout(300)
            loc.first.click(timeout=8000)
            pg.wait_for_timeout(2200)
        except Exception:
            break
        n = pg.content().count('class="product-card"')
        if n <= last:
            same += 1
            if same >= 2:
                break
        else:
            same = 0
        last = n
    return last


def render(url, wait=2500, expand=True):
    from playwright.sync_api import sync_playwright
    with sync_playwright() as pw:
        b = pw.chromium.launch(headless=False, channel="chrome")
        pg = b.new_page()
        pg.goto(url, wait_until="networkidle", timeout=90000)
        pg.wait_for_timeout(wait)
        accept_cookies(pg)
        if expand:
            load_all(pg)
        html = pg.content()
        b.close()
    return html


CARD = re.compile(r'<div class="product-card">(.*?)(?=<div class="product-card">|$)', re.S)
TOPLINE = re.compile(r'product-card__topline">(.*?)</span>', re.S)
HEADLINE = re.compile(r'product-card__headline">(.*?)</h2>', re.S)
DESC_FULL = re.compile(r'product-card__description--list"><div><p>(.*?)</p>', re.S)
HREF = re.compile(r'product-card__container" href="([^"]+)"')
IMG = re.compile(r'<img[^>]+src="([^"]+)"')


def parse_items(html):
    """Trang ?class= -> [{art, name, desc, url}] theo dung 3 cot trong template."""
    out, seen = [], set()
    for blk in CARD.findall(html):
        art = txt((TOPLINE.search(blk) or [None, ""]) and
                  (TOPLINE.search(blk).group(1) if TOPLINE.search(blk) else ""))
        if not art or art in seen:
            continue
        seen.add(art)
        nm = HEADLINE.search(blk)
        d = DESC_FULL.search(blk)
        u = HREF.search(blk)
        im = IMG.search(blk)
        src = im.group(1) if im else ""
        if src.startswith("/_ipx/"):          # xin ban to hon ban thumbnail w_200
            src = re.sub(r"/w_\d+/", "/w_740/", src)
        out.append({"art": art,
                    "name": txt(nm.group(1)) if nm else "",
                    "desc": txt(d.group(1)) if d else "",
                    "img": SITE + src if src.startswith("/") else src,
                    "url": SITE + u.group(1) if u else ""})
    return out



# ---------------------------------------------------------------------------
# 2. "Mo ta dai" = cac cap Key: Value hay gap nhat cua ca dong
# ---------------------------------------------------------------------------
def _groups(items):
    """Gom cac cap 'Key: Value' trong mo ta cua ca dong SP.

    P+F lap lai cung mot ten thong so nhieu lan trong 1 mo ta (vd 'Output type'
    vua la PNP/NPN vua la 3-wire/4-wire) -> gop chung lai lam mot.
    """
    pairs = collections.defaultdict(collections.Counter)
    for it in items:
        for part in re.split(r",(?![^(]*\))", it.get("desc") or ""):
            part = part.strip()
            if ":" not in part:
                continue
            k, v = part.split(":", 1)
            k, v = k.strip(), v.strip()
            if k and v:
                pairs[k][v] += 1
    return sorted(pairs.items(), key=lambda kv: -sum(kv[1].values()))


# Thong so chi phan loai catalog / lap lai ten dong SP - khong phai thong so
# ky thuat, bo khoi ca Mo ta dai lan Mo ta ngan (chot 19/10).
CLASS_KEY = re.compile(r"^\s*(series|construction\s*type|design|product|type|"
                       r"product\s*(sub)?segment|product\s*family|"
                       r"product\s*categor(y|ies))\s*$", re.I)


def _norm_vals(cnt):
    """Tap gia tri da chuan hoa, de so 2 nhom co trung noi dung khong."""
    out = set()
    for v in cnt:
        w = re.sub(r"(type|series|version)", " ", v.lower())
        w = re.sub(r"[^a-z0-9 ]+", " ", w)
        w = " ".join(w.split())
        if w:
            out.add(w)
    return out


def _dup_of_line(key, cnt, line_en):
    """Thong so chi lap lai chinh ten dong SP -> bo (theo ghi chu trong template)."""
    ln = re.sub(r"[^a-z ]", " ", (line_en or "").lower()).split()
    ln = {w.rstrip("s") for w in ln if len(w) > 3}
    if not ln:
        return False
    for v in cnt:
        words = {w.rstrip("s") for w in re.sub(r"[^a-z ]", " ", v.lower()).split()
                 if len(w) > 3}
        if words and words <= ln:
            return True
    return False


def fmt_group(key, cnt, max_vals=6):
    vals = [v for v, _c in cnt.most_common(max_vals)]
    more = len(cnt) - len(vals)
    return "%s: %s%s" % (key, "; ".join(vals),
                         " (+%d tuỳ chọn khác)" % more if more > 0 else "")


def prose_descs(items, limit=15):
    """Mo ta dang VAN XUOI (khong co cap 'Key: Value').

    Nhieu dong cua P+F - nhat la nhom Enterprise Mobility - khong liet ke thong
    so ma viet mot cau ngan cho tung ma ('Lightweight duplex headset'). Khi do
    lay chinh cac cau do, bo trung, lam gach dau dong.
    """
    out = []
    for it in items:
        d = (it.get("desc") or "").strip()
        if not d or ":" in d:
            continue
        if d not in out:
            out.append(d)
        if len(out) >= limit:
            break
    return out


def long_desc(items, line_en):
    """Mo ta DAI = FULL thong so, bo:
       - nhom chi phan loai catalog (Series, Construction type, Product...)
       - nhom lap lai chinh ten dong SP
    """
    out = []
    for k, cnt in _groups(items):
        # CHI bo nhom phan loai catalog + nhom lap lai ten dong SP.
        # KHONG bo theo trung gia tri: 'Output polarity' va 'Type of voltage'
        # deu la DC/AC nhung la hai thong so khac nhau.
        if CLASS_KEY.match(k) or _dup_of_line(k, cnt, line_en):
            continue
        out.append(fmt_group(k, cnt))
    return out or prose_descs(items)


def short_desc(items, line_en, n=MAX_HL):
    """Mo ta NGAN = 5-8 thong so quan trong nhat (xep theo so ma dung)."""
    return long_desc(items, line_en)[:n]


# ---------------------------------------------------------------------------
# 3. Danh muc Ntrend + ten tieng Viet
# ---------------------------------------------------------------------------
def ntrend_row(name_en):
    from openpyxl import load_workbook
    ws = load_workbook(NTREND_XLSX, read_only=True, data_only=True)["02 Pepperl+Fuchs"]
    for r in ws.iter_rows(min_row=6, values_only=True):
        if not r or not r[0] or len(r) < 10:
            continue
        if str(r[8] or "").strip() != "x":
            continue
        if (r[1] or "").strip() == name_en.strip():
            cat = (r[9] or "").strip()
            if cat == BRAND:        # chi ghi ten hang = khong co danh muc rieng
                cat = ("Không có danh mục sản phẩm,  chỉ nằm trong danh mục "
                       "Thương Hiệu Pepperl+Fuchs")
            return {"name_vi": (r[2] or "").strip(), "cat_vi": cat}
    return {"name_vi": "", "cat_vi": ""}


# ---------------------------------------------------------------------------
# 4. Dung file Excel theo template
# ---------------------------------------------------------------------------
META_TPL = ("Mua {type_l} {name_brand} chính hãng – thông số kỹ thuật đầy đủ, "
            "giá tốt, hỗ trợ tư vấn và lựa chọn model phù hợp. Liên hệ Ntrend ngay!")
SPECDESC = (
    "{type_c} {name_brand} là thiết bị công nghiệp được sử dụng trong các hệ "
    "thống tự động hóa và điều khiển công nghiệp, đáp ứng các yêu cầu về độ "
    "chính xác, độ ổn định và độ tin cậy trong quá trình vận hành. Sản phẩm "
    "được thiết kế phù hợp với nhiều ứng dụng trong nhà máy và dây chuyền sản "
    "xuất, hỗ trợ đo lường, phát hiện, giám sát, điều khiển và tối ưu hóa quy trình.\n\n"
    "{type_c} được sản xuất theo các tiêu chuẩn kỹ thuật của {brand}, với các "
    "thông số kỹ thuật được công bố rõ ràng, giúp kỹ sư và bộ phận kỹ thuật dễ "
    "dàng lựa chọn, tích hợp và triển khai vào hệ thống tự động hóa.\n\n"
    "Tùy thuộc vào từng model, sản phẩm có thể có các tùy chọn khác nhau, đáp "
    "ứng đa dạng yêu cầu của các ứng dụng công nghiệp.\n\n"
    "Dưới đây là thông tin chi tiết của sản phẩm:")


import unicodedata


def slugify(s):
    """Keyword chinh (tieng Viet) -> URL: bo dau, thuong hoa, noi bang '-'.

    Chuan SEO: keyword chinh phai nam trong URL.
        'Cam bien dien dung Pepperl+Fuchs' -> 'cam-bien-dien-dung-pepperl-fuchs'
    """
    s = (s or "").replace("Đ", "D").replace("đ", "d")
    s = unicodedata.normalize("NFD", s)
    s = "".join(c for c in s if unicodedata.category(c) != "Mn")
    s = re.sub(r"[^a-zA-Z0-9]+", "-", s).lower().strip("-")
    return re.sub(r"-{2,}", "-", s) or "sp"


def build(line_en, data, out_dir):
    from openpyxl import Workbook
    from openpyxl.styles import Font, Alignment, PatternFill
    from openpyxl.utils import get_column_letter
    HDR = PatternFill("solid", fgColor="1F4E79")
    FLD = PatternFill("solid", fgColor="DDEBF7")
    WRAP = Alignment(wrap_text=True, vertical="top")
    os.makedirs(out_dir, exist_ok=True)
    wb = Workbook(); ws = wb.active; ws.title = "NOI DUNG"
    ws["A1"] = "NỘI DUNG SẢN PHẨM NTREND - %s" % line_en
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["Trường", "Tiếng Việt (bản đăng web)", "Tiếng Anh (soát chính tả)"])
    for c in range(1, 4):
        ws.cell(row=3, column=c).fill = HDR
        ws.cell(row=3, column=c).font = Font(bold=True, color="FFFFFF")

    def row(k, vi, en=""):
        ws.append([k, vi, en])
        r = ws.max_row
        ws.cell(row=r, column=1).fill = FLD
        ws.cell(row=r, column=1).font = Font(bold=True)
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).alignment = WRAP

    nv = data["name_vi"] or line_en
    # Ten dong SP cua P+F CHINH LA loai san pham ("Capacitive Sensors"), nen
    # {type_c} va {name} trung nhau -> {name_brand} chi giu thuong hieu, neu
    # khong se ra "Mua cam bien dien dung Cam bien dien dung Pepperl+Fuchs".
    first = nv.split(" ")[0]
    cv = {"type_c": nv[:1].upper() + nv[1:],
          "type_l": nv if (len(first) > 1 and first.isupper()) else nv[:1].lower() + nv[1:],
          "name": nv, "name_brand": BRAND, "brand": BRAND}
    ce = {"type_c": line_en, "type_l": line_en.lower(),
          "name": line_en, "name_brand": BRAND, "brand": BRAND}

    row("Ảnh sản phẩm", data.get("img") or "(chưa có ảnh)")
    row("Danh mục sản phẩm", data["cat_vi"] or "(chưa phân loại)")
    row("Thẻ tag", "%s, %s" % (nv, BRAND), "%s, %s" % (line_en, BRAND))
    row("Thương hiệu", BRAND, BRAND)
    kw_vi = "%s %s" % (nv, BRAND)
    kw_en = "%s %s" % (line_en, BRAND)
    row("URL", slugify(kw_vi), slugify(kw_en))
    row("Keyword chính", kw_vi, kw_en)
    row("Meta title", "%s %s" % (cv["type_c"], BRAND), "%s %s" % (line_en, BRAND))
    row("Meta description (155 - 165 characters)",
        META_TPL.format(**cv), META_TPL.format(**ce))
    row("Mô tả dài của sản phẩm",
        "\n".join("- " + x for x in data["long_vi"]),
        "\n".join("- " + x for x in data["long_en"]))
    row("Mô tả ngắn của sản phẩm",
        "\n".join("- " + x for x in data["short_vi"]),
        "\n".join("- " + x for x in data["short_en"]))
    row("Mô tả thông số", SPECDESC.format(**cv), SPECDESC.format(**ce))

    ws.append(["Bảng mã sản phẩm + Thông tin sản phẩm", "Mã sản phẩm",
               "Tên sản phẩm", "Mô tả (VI)", "Description (EN)"])
    r = ws.max_row
    ws.cell(row=r, column=1).fill = FLD
    ws.cell(row=r, column=1).font = Font(bold=True)
    for c in range(2, 6):
        ws.cell(row=r, column=c).font = Font(bold=True, italic=True)
    for it in data["items"]:
        ws.append(["", it["art"], it["name"], it.get("desc_vi", ""), it["desc"]])
        for c in range(2, 6):
            ws.cell(row=ws.max_row, column=c).alignment = WRAP
    ws.append([])
    ws.append(["Nguồn", data["url"], ""])
    for i, w in enumerate([34, 58, 58, 58, 58], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    p = os.path.join(out_dir, "THU - %s.xlsx" % re.sub(r'[\/:*?"<>|]+', "-", line_en))
    wb.save(p)
    return p


def line_image(items, line_en):
    """Anh dai dien cua dong SP = anh ma SP dau tien, ghep nen NTREND."""
    import urllib.request
    url = next((it.get("img") for it in items if it.get("img")), "")
    if not url:
        return ""
    img_dir = os.path.join(OUT_DIR, "images")
    src_dir = os.path.join(HERE, "out", "02 Pepperl+Fuchs", "src_img")
    for d in (img_dir, src_dir):
        os.makedirs(d, exist_ok=True)
    base = slugify(line_en)
    src = os.path.join(src_dir, base + os.path.splitext(url.split("?")[0])[1][:5])
    try:
        req = urllib.request.Request(url, headers={"User-Agent": UA})
        data = urllib.request.urlopen(req, timeout=60).read()
        if len(data) < 500:
            return ""
        with open(src, "wb") as f:
            f.write(data)
        import brandimg
        out = brandimg.compose(src, os.path.join(img_dir, base + ".jpg"))
        return os.path.basename(out)
    except Exception as e:
        print("   [anh loi]", repr(e)[:70])
        return ""


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    cls = sys.argv[1] if len(sys.argv) > 1 else "30736"
    line_en = sys.argv[2] if len(sys.argv) > 2 else "Capacitive Sensors"
    url = LIST_URL % cls
    print("1. Tai trang danh sach SP (Chrome that):", url)
    html = render(url)
    items = parse_items(html)
    print("   -> %d ma san pham" % len(items))


    long_en = long_desc(items, line_en)
    short_en = short_desc(items, line_en)
    print("3. Mo ta dai: %d thong so (full) | Mo ta ngan: %d thong so"
          % (len(long_en), len(short_en)))
    for x in long_en:
        print("     -", x[:110])

    meta = ntrend_row(line_en)
    print("4. Ntrend: ten VI = %r | danh muc = %r" % (meta["name_vi"], meta["cat_vi"][:50]))

    print("5. Anh san pham...")
    img_name = line_image(items, line_en)
    print("   ->", img_name or "(khong lay duoc)")

    print("6. Dich...")
    from vi_engine import translate_many
    long_vi = translate_many(long_en, "value") if long_en else []
    short_vi = long_vi[:len(short_en)]
    d_en = [it["desc"] for it in items]
    d_vi = translate_many(d_en, "value") if d_en else []
    for it, v in zip(items, d_vi):
        it["desc_vi"] = v
    from translate import _flush
    _flush(True)

    data = dict(meta, items=items, long_en=long_en, long_vi=long_vi,
                short_en=short_en, short_vi=short_vi, url=url,
                img=img_name)
    p = build(line_en, data, OUT_DIR)
    print("7. ->", p)


if __name__ == "__main__":
    main()
