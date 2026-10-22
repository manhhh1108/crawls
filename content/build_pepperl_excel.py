# -*- coding: utf-8 -*-
"""Pepperl+Fuchs - PHA 3: content.json -> moi bai 1 file Excel.

    python build_pepperl_excel.py

Layout theo 'data/template_pepper fuch.docx' (chot 19-20/10):
    Mo ta dai   = FULL thong so cua dong SP (bo nhom chi phan loai catalog)
    Mo ta ngan  = 5-8 thong so quan trong nhat
    URL         = keyword chinh tieng Viet, bo dau, noi bang '-'
    Bang ma SP  = Ma SP | Ten SP | Mo ta   (nhu bang Variants tren web)

Ten file phai kem nhanh danh muc: 16 nhom trung ten nhung khac trang (vd
'Accessories' o 16 nhanh) - neu dat ten chi theo ten dong SP thi de len nhau.
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

from c_pepperl_try import BRAND, slugify, SPECDESC, META_TPL   # noqa: E402

SRC = os.path.join(HERE, "out", "02 Pepperl+Fuchs", "content.json")
OUT_DIR = os.path.join(REPO, "data", "Noi dung Pepperl+Fuchs")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
FLD_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")


def safe(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s or "").strip()[:n].rstrip(" .")


def ctx(a):
    """{type_c}/{type_l}/{name_brand} cho 2 cot VI va EN."""
    nv = a.get("name_vi") or a["name"]
    first = nv.split(" ")[0]
    cv = {"type_c": nv[:1].upper() + nv[1:],
          "type_l": nv if (len(first) > 1 and first.isupper())
                    else nv[:1].lower() + nv[1:],
          "name": nv, "name_brand": BRAND, "brand": BRAND}
    ce = {"type_c": a["name"], "type_l": a["name"].lower(),
          "name": a["name"], "name_brand": BRAND, "brand": BRAND}
    return cv, ce


def build_one(a, idx, out_dir):
    cv, ce = ctx(a)
    nv = cv["name"]
    wb = Workbook()
    ws = wb.active
    ws.title = "NOI DUNG"
    ws["A1"] = "NỘI DUNG SẢN PHẨM NTREND - %s" % a["name"]
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["Trường", "Tiếng Việt (bản đăng web)", "Tiếng Anh (soát chính tả)"])
    for c in range(1, 4):
        ws.cell(row=3, column=c).fill = HDR_FILL
        ws.cell(row=3, column=c).font = Font(bold=True, color="FFFFFF")

    def row(k, vi, en=""):
        ws.append([k, vi, en])
        r = ws.max_row
        ws.cell(row=r, column=1).fill = FLD_FILL
        ws.cell(row=r, column=1).font = Font(bold=True)
        for c in (1, 2, 3):
            ws.cell(row=r, column=c).alignment = WRAP

    kw_vi = "%s %s" % (nv, BRAND)
    kw_en = "%s %s" % (a["name"], BRAND)
    row("Ảnh sản phẩm", a.get("img_file") or "(chưa có ảnh)")
    row("Danh mục sản phẩm", a.get("cat_vi") or "(chưa phân loại)")
    row("Thẻ tag", "%s, %s" % (nv, BRAND), "%s, %s" % (a["name"], BRAND))
    row("Thương hiệu", BRAND, BRAND)
    row("URL", slugify(kw_vi), slugify(kw_en))
    row("Keyword chính", kw_vi, kw_en)
    row("Meta title", "%s %s" % (cv["type_c"], BRAND), kw_en)
    row("Meta description (155 - 165 characters)",
        META_TPL.format(**cv), META_TPL.format(**ce))
    row("Mô tả dài của sản phẩm",
        "\n".join("- " + x for x in a.get("long_vi") or []),
        "\n".join("- " + x for x in a.get("long_en") or []))
    row("Mô tả ngắn của sản phẩm",
        "\n".join("- " + x for x in a.get("short_vi") or []),
        "\n".join("- " + x for x in a.get("short_en") or []))
    row("Mô tả thông số", SPECDESC.format(**cv), SPECDESC.format(**ce))

    items = a.get("items") or []
    if items:
        ws.append(["Bảng mã sản phẩm + Thông tin sản phẩm", "Mã sản phẩm",
                   "Tên sản phẩm", "Mô tả (VI)", "Description (EN)"])
        r = ws.max_row
        ws.cell(row=r, column=1).fill = FLD_FILL
        ws.cell(row=r, column=1).font = Font(bold=True)
        for c in range(2, 6):
            ws.cell(row=r, column=c).font = Font(bold=True, italic=True)
        for it in items:
            ws.append(["", it.get("art", ""), it.get("name", ""),
                       it.get("desc_vi", ""), it.get("desc", "")])
            for c in range(2, 6):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP
    else:
        ws.append(["Bảng mã sản phẩm + Thông tin sản phẩm",
                   "(trang hãng không có dữ liệu)", "", "", ""])
        ws.cell(row=ws.max_row, column=1).fill = FLD_FILL
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True)

    ws.append([])
    ws.append(["Nhánh danh mục", " > ".join(a.get("path") or []), ""])
    ws.append(["Nguồn", a.get("url_src", ""), ""])
    for i, w in enumerate([34, 58, 58, 58, 58], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"

    fname = "%03d - %s.xlsx" % (idx, safe(a["name"]))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_index(arts, names, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "MUC LUC"
    ws["A1"] = "MỤC LỤC BÀI ĐĂNG - 02 Pepperl+Fuchs"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["STT", "Dòng sản phẩm", "Tên tiếng Việt", "File", "Ảnh",
               "Số mã SP", "Số thông số", "Danh mục Ntrend", "Nhánh danh mục",
               "Nguồn"])
    for c in range(1, 11):
        ws.cell(row=3, column=c).fill = FLD_FILL
        ws.cell(row=3, column=c).font = Font(bold=True)
    for i, (a, f) in enumerate(zip(arts, names), 1):
        ws.append([i, a["name"], a.get("name_vi", ""), f,
                   "x" if a.get("img_file") else "",
                   len(a.get("items") or []), len(a.get("long_en") or []),
                   a.get("cat_vi", ""), " > ".join(a.get("path") or []),
                   a.get("url_src", "")])
    for i, w in enumerate([6, 38, 34, 46, 6, 10, 12, 50, 52, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    arts = json.load(io.open(SRC, encoding="utf-8"))["articles"]
    os.makedirs(OUT_DIR, exist_ok=True)
    names = []
    for i, a in enumerate(arts, 1):
        names.append(build_one(a, i, OUT_DIR))
    build_index(arts, names, OUT_DIR)
    nitem = sum(len(a.get("items") or []) for a in arts)
    print("Saved %d bai (%d ma SP) -> %s" % (len(arts), nitem, OUT_DIR))
    empty = [a["name"] for a in arts if not a.get("items")]
    if empty:
        print("  bai khong co ma SP (%d): %s" % (len(empty), ", ".join(empty[:8])))


if __name__ == "__main__":
    main()
