# -*- coding: utf-8 -*-
"""Tree Electric - PHA 3: content.json -> moi bai 1 file Excel.

    python build_tree_excel.py

Layout theo 'Template nội dung Tree Electric.docx'. Template nay ngan nhat
trong ca bo: khong co khoi "Bang ma SP", khong co khoi "Thong so ky thuat",
doi lai co them dong "Link datasheet sản phẩm".

Ly do: trang hang chi dang anh bia + catalogue PDF cho tung dong SP, khong dang
ma SP lan thong so nao - chinh template cung ghi o o mo ta dai la "Không cần
lấy vì trên web không có thông tin".
"""
import io
import json
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
sys.path.insert(0, HERE)

from openpyxl import Workbook                              # noqa: E402
from openpyxl.styles import Font, Alignment, PatternFill   # noqa: E402
from openpyxl.utils import get_column_letter               # noqa: E402

import out_dir                                         # noqa: E402
import logo_fallback                                  # noqa: E402
from bai_prune import prune                             # noqa: E402
from c_pepperl_try import (slugify, SPECDESC, META_TPL,    # noqa: E402
                           SPECDESC_EN, META_TPL_EN,
                           DESCDESC, DESCDESC_EN)

BRAND = "Tree Electric"
BRAND_FULL = "Tree Electric (Shanghai Quyi Electric)"
SRC = os.path.join(HERE, "out", "13 Tree Electric", "content.json")
OUT_DIR = out_dir.brand_dir("Noi dung Tree Electric")

# Bai cua hang nay khong co bang thong so nao - thu dung ngay duoi o "Mo ta
# thong so" la LINK CATALOGUE PDF, nen cau dan phai noi dung cai sap den
# (review khach 02/11), khong phai "thong tin chi tiet" chung chung.
_CU = "Dưới đây là thông tin chi tiết của sản phẩm:"
_MOI = "Dưới đây là datasheet chi tiết của sản phẩm:"
_CU_EN = "Detailed product information is listed below:"
_MOI_EN = "The detailed product datasheet is provided below:"
assert _CU in SPECDESC and _CU_EN in SPECDESC_EN, "mau SPECDESC da doi"
SPEC_DS = SPECDESC.replace(_CU, _MOI)
SPEC_DS_EN = SPECDESC_EN.replace(_CU_EN, _MOI_EN)

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
FLD_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")


def safe(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s or "").strip()[:n].rstrip(" .")


def thuong(s):
    """Ha thuong cau tieng Anh nhung GIU NGUYEN ten dong SP.

    `en.lower()` bien "Solid State Relays Series I" thanh "...series i" - doc
    ra nhu loi chinh ta ngay trong o giao khach. Chi ha thuong tu viet hoa chu
    dau thuong; tu co chu hoa ben trong hoac co so giu nguyen.
    """
    out = []
    for w in (s or "").split(" "):
        lot = w.strip(",.;:()")
        out.append(w.lower() if lot[:1].isupper() and lot[1:].islower() else w)
    return " ".join(out)


def ctx(a):
    en = a.get("title_en") or a["name"]
    nv = a.get("name_vi") or en
    first = nv.split(" ")[0].strip(",.;:()")
    # Chu dau chi ha thuong khi tu dau la tu thuong, khong thi ten rieng
    # nhu "TUK"/"PT" bi ha thanh "tUK".
    cv = {"type_c": nv[:1].upper() + nv[1:],
          "type_l": nv[:1].lower() + nv[1:] if first[1:].islower() else nv,
          "name": nv, "name_brand": BRAND, "brand": BRAND}
    ce = {"type_c": en, "type_l": thuong(en),
          "name": en, "name_brand": BRAND, "brand": BRAND}
    return cv, ce


def build_one(a, idx, out_dir):
    cv, ce = ctx(a)
    nv, en = cv["name"], ce["name"]
    wb = Workbook()
    ws = wb.active
    ws.title = "NOI DUNG"
    ws["A1"] = "NỘI DUNG SẢN PHẨM NTREND - %s" % en
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["Trường", "Tiếng Việt (bản đăng web)", "Tiếng Anh (soát chính tả)"])
    for c in range(1, 4):
        ws.cell(row=3, column=c).fill = HDR_FILL
        ws.cell(row=3, column=c).font = Font(bold=True, color="FFFFFF")

    def row(k, vi, en_=""):
        ws.append([k, vi, en_])
        r = ws.max_row
        ws.cell(row=r, column=1).fill = FLD_FILL
        ws.cell(row=r, column=1).font = Font(bold=True)
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).alignment = WRAP

    kw_vi = "%s %s" % (nv, BRAND)
    kw_en = "%s %s" % (en, BRAND)
    row("Ảnh sản phẩm", a.get("img_file") or "(chưa có ảnh)")
    row("Danh mục sản phẩm", a.get("cat_vi") or "(chưa phân loại)")
    # Trang hang khong dang ma SP nao -> the tag lay ten dong SP.
    row("Thẻ tag", "%s, %s" % (nv, BRAND), "%s, %s" % (en, BRAND))
    row("Thương hiệu", BRAND_FULL, BRAND_FULL)
    row("URL", slugify(kw_vi), slugify(kw_en))
    row("Keyword chính", kw_vi, kw_en)
    row("Meta title", "%s %s" % (cv["type_c"], BRAND), kw_en)
    row("Meta description (155 - 165 characters)",
        META_TPL.format(**cv), META_TPL_EN.format(**ce))
    # Trang hang khong co chu nao gioi thieu (template ghi ro "Không cần lấy vì
    # trên web không có thông tin") -> dung mau nhu cac bai Festo cung canh.
    # Mo ta ngan dung DU ca hai cau cua mau (review khach 02/11): trang hang
    # khong co chu nao khac bu vao nua cau bi cat.
    dv, de = DESCDESC.format(**cv), DESCDESC_EN.format(**ce)
    row("Mô tả dài của sản phẩm", dv, de)
    row("Mô tả ngắn của sản phẩm", dv, de)
    row("Mô tả thông số", SPEC_DS.format(**cv), SPEC_DS_EN.format(**ce))
    row("Link datasheet sản phẩm",
        a.get("datasheet") or "(trang hãng không có catalogue PDF)")

    ws.append([])
    ws.append(["Nhánh danh mục", " > ".join(a.get("path") or []), ""])
    ws.append(["Nguồn", a.get("url_src", ""), ""])
    ws.append(["Ảnh bìa catalogue (tham khảo)", a.get("cat_img", ""), ""])
    for i, w in enumerate([34, 70, 70], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    fname = "%02d - %s.xlsx" % (idx, safe(en))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_index(arts, names, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "MUC LUC"
    ws["A1"] = "MỤC LỤC BÀI ĐĂNG - 13 Tree Electric"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["STT", "Dòng sản phẩm", "Tên tiếng Việt", "File", "Ảnh",
               "Link datasheet", "Danh mục Ntrend", "Nhánh danh mục", "Nguồn"])
    for c in range(1, 10):
        ws.cell(row=3, column=c).fill = FLD_FILL
        ws.cell(row=3, column=c).font = Font(bold=True)
    for i, (a, f) in enumerate(zip(arts, names), 1):
        ws.append([i, a.get("title_en") or a["name"], a.get("name_vi", ""), f,
                   "x" if a.get("img_file") else "",
                   a.get("datasheet", ""), a.get("cat_vi", ""),
                   " > ".join(a.get("path") or []), a.get("url_src", "")])
    for i, w in enumerate([6, 40, 40, 48, 6, 56, 54, 40, 50], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    # Anh duy nhat tren trang hang la anh bia catalogue chu khong phai anh san
    # pham -> moi bai deu lay logo hang.
    logo_fallback.run("tree")
    arts = json.load(io.open(SRC, encoding="utf-8"))["articles"]
    os.makedirs(OUT_DIR, exist_ok=True)
    names = [build_one(a, i, OUT_DIR) for i, a in enumerate(arts, 1)]
    build_index(arts, names, OUT_DIR)
    prune(OUT_DIR, names)
    print("Saved %d bai -> %s" % (len(arts), OUT_DIR))
    bad = [a for a in arts if not a.get("datasheet")]
    if bad:
        print("  khong co link datasheet (%d): %s"
              % (len(bad), ", ".join(x["title_en"] for x in bad)))


if __name__ == "__main__":
    main()
