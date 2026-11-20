# -*- coding: utf-8 -*-
"""Tree Electric / Shanghai Quyi (hang 13) - crawl noi dung bai dang.

    python c_tree.py              # chay tiep nhung dong chua co
    python c_tree.py --force      # lam lai tu dau

Can chay truoc: python cat_map.py tree

CA 9 DONG SP NAM TREN MOT TRANG. www.tree-electric.com/templates/products.html
tai ve day du 9 khoi san pham; tham so ?index=N chi la de JavaScript hien khoi
nao, khong doi noi dung tai ve. Nen crawler chi tai MOT trang roi cat ra.

Moi khoi chi co ba thu: ten muc (tieng Trung), anh bia catalogue, va link
catalogue PDF:

    <div name="products-sort-listN" ...>
        <div class="products-detail-title"> 辅材综合类
            <a href="../Image/support/list_pdf1.pdf">点击立即阅读>>></a>
        <a href="../Image/support/list_pdf1.pdf">
            <img src="../Image/products/list1.png">

Cho nen theo dung 'Template nội dung Tree Electric.docx':
    Mo ta dai   = khong lay (template ghi "Không cần lấy vì trên web không có
                  thông tin") -> dung mau DESCDESC nhu cac bai Festo khong co
                  mo ta, de bai dang len web khong trong tron
    Mo ta ngan  = cau dau cua mau do
    Link datasheet = link catalogue PDF
Template cua hang nay KHONG co o "Bang ma SP" lan o "Thong so ky thuat": ma SP
cua ho chi nam trong catalogue PDF, trang web khong dang ma nao.

ANH: anh duy nhat trong khoi la ANH BIA CATALOGUE - mot to A4 900x1143, nua
tren la khoi anh (anh nganh + anh san pham tren nen xanh), nua duoi la ten muc
bang chu Trung, logo hang va khoang trang. Khach khoanh dung NUA TREN va yeu
cau dung no lam anh bai (review 02/11: "hien tai dang la anh logo"), nen
`cat_trang()` cat bo nua duoi roi moi ghep khung NTREND.

Ket qua: out/13 Tree Electric/content.json -> build_tree_excel.py dung tiep.
"""
import io
import json
import os
import re
import sys
import html as H

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)
sys.path.insert(0, REPO)

import out_dir                                         # noqa: E402

BRAND = "Tree Electric"
SITE = "http://www.tree-electric.com"
PAGE = SITE + "/templates/products.html"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0 Safari/537.36")
OUT = os.path.join(HERE, "out", "13 Tree Electric")
CONTENT = os.path.join(OUT, "content.json")
URLMAP = os.path.join(OUT, "urlmap.json")
SRC_IMG = os.path.join(OUT, "src_img")
IMG_DIR = os.path.join(out_dir.brand_dir("Noi dung Tree Electric"), "images")
for _d in (OUT, SRC_IMG, IMG_DIR):
    os.makedirs(_d, exist_ok=True)

# Khoi san pham ket thuc khi gap khoi ke tiep hoac het vung san pham.
BLOCK = re.compile(r'(?s)<div name="products-sort-list(\d+)"(.*?)'
                   r'(?=<div name="products-sort-list|<!--end)')
TITLE = re.compile(r'(?s)class="products-detail-title">\s*([^<]+)')
PDF = re.compile(r'href="([^"]+\.pdf)"')
IMG = re.compile(r'<img src="([^"]+)"')


def txt(s):
    return re.sub(r"\s+", " ", H.unescape(re.sub(r"<[^>]+>", " ", s or ""))).strip()


def abs_url(u):
    """Link trong trang la '../Image/...' - trang nam o /templates/."""
    return SITE + "/" + u.lstrip("./") if u.startswith("..") else u


def blocks():
    """{ten tieng Trung: {pdf, img}} - doc ca phan bi <!-- --> boc.

    Muc "工业交换机系列" (Industrial Switches) bi comment o THANH BEN nhung khoi
    noi dung cua no van con trong trang, va danh muc Ntrend co dong do - nen
    khong duoc bo qua phan trong comment.
    """
    from shared.fetcher import try_fetch
    s = try_fetch(PAGE, retries=2)
    if not s:
        return {}
    if isinstance(s, bytes):
        s = s.decode("utf-8", "replace")
    out = {}
    for _n, body in BLOCK.findall(s):
        t = TITLE.search(body)
        if not t:
            continue
        p, im = PDF.search(body), IMG.search(body)
        out[txt(t.group(1))] = {
            "pdf": abs_url(p.group(1)) if p else "",
            "img": abs_url(im.group(1)) if im else ""}
    return out


def cat_trang(im):
    """Vi tri cat bo nua duoi to catalogue.

    Tim dai khoang trang dau tien sau 1/6 chieu cao - do la khoang giua khoi
    anh va dong ten muc chu Trung. Ca 9 to deu cat o 57-60% chieu cao.
    """
    g = im.convert("L")
    w, h = g.size
    px = g.load()
    buoc = max(1, w // 120)
    dai = max(8, h // 60)
    trong = 0
    for y in range(h):
        if min(px[x, y] for x in range(0, w, buoc)) >= 244:
            trong += 1
            if trong >= dai and y - trong > h // 6:
                return y - trong
        else:
            trong = 0
    return h


def grab_image(url, line_en, slugify):
    """Tai anh bia, cat nua duoi, ghep khung NTREND -> ten file anh bai."""
    import urllib.request
    from PIL import Image
    if not url:
        return ""
    base = slugify(line_en)[:60].strip("-") or "sp"
    src = os.path.join(SRC_IMG, base + ".png")
    try:
        if not (os.path.exists(src) and os.path.getsize(src) > 2000):
            req = urllib.request.Request(
                url, headers={"User-Agent": UA, "Referer": PAGE})
            data = urllib.request.urlopen(req, timeout=90).read()
            if len(data) < 2000:
                return ""
            with open(src, "wb") as f:
                f.write(data)
        cut = os.path.join(SRC_IMG, base + "-cat.jpg")
        with Image.open(src) as im:
            im = im.convert("RGB")
            im.crop((0, 0, im.size[0], cat_trang(im))).save(cut, quality=95)
        import brandimg
        return os.path.basename(
            brandimg.compose(cut, os.path.join(IMG_DIR, base + ".jpg")))
    except Exception as e:
        print("     [anh loi]", repr(e)[:60])
        return ""


# Ten trong danh muc Ntrend la "English (中文)" - cat ra de doi chieu voi trang.
BILABEL = re.compile(r"^(.*?)\s*\(([^()]*[一-鿿][^()]*)\)\s*$")


def tach_ten(name):
    m = BILABEL.match(name or "")
    return (m.group(1).strip(), m.group(2).strip()) if m else (name, "")


def save(arts):
    with io.open(CONTENT, "w", encoding="utf-8") as f:
        json.dump({"brand": "13 Tree Electric", "articles": arts}, f,
                  ensure_ascii=False)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    force = "--force" in sys.argv

    from vi_engine import translate_many
    from translate import _flush
    from c_pepperl_try import slugify

    ls = [l for l in json.load(io.open(URLMAP, encoding="utf-8"))["lines"]
          if l.get("url")]
    done = {}
    if os.path.exists(CONTENT) and not force:
        done = {a["name"]: a for a in
                json.load(io.open(CONTENT, encoding="utf-8"))["articles"]}

    bl = blocks()
    print("trang hang co %d khoi san pham | danh muc Ntrend co %d dong"
          % (len(bl), len(ls)))
    if not bl:
        print("KHONG TAI DUOC trang san pham - dung lai")
        return

    todo = [l for l in ls if l["name"] not in done]
    if not todo:
        print("da xong het %d bai" % len(done))
        return

    tens = [tach_ten(l["name"])[0] for l in todo]
    vi = translate_many(tens, "key")
    for ln, en, nv in zip(todo, tens, vi):
        _en, cn = tach_ten(ln["name"])
        d = bl.get(cn) or {}
        if not d:
            print("  %-42s KHONG KHOP KHOI NAO (%s)" % (en[:42], cn))
        img = grab_image(d.get("img", ""), en, slugify)
        done[ln["name"]] = dict(
            ln, url_src=ln["url"], title_en=en, name_vi=nv,
            series="", codes=[], img_file=img,
            datasheet=d.get("pdf", ""), cat_img=d.get("img", ""),
            long_en=[], long_vi=[], short_en=[], short_vi=[])
        print("  %-42s %-12s %s"
              % (en[:42], "pdf" if d.get("pdf") else "KHONG CO PDF",
                 img or "KHONG CO ANH"))
    _flush(True)
    save([done[l["name"]] for l in ls if l["name"] in done])
    print("\nxong %d bai -> %s" % (len(done), CONTENT))


if __name__ == "__main__":
    main()
