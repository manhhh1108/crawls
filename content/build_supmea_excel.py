# -*- coding: utf-8 -*-
"""Supmea - PHA 3: content.json -> moi bai 1 file Excel.

    python build_supmea_excel.py

Layout theo 'Template nội dung Supmea.docx'. Khac cac hang truoc o mot diem:
template cua Supmea KHONG co o "Bang ma san pham" - trang hang khong ban theo
ma dat hang ma moi trang la mot san pham, nen thay vao do la khoi
"Thong so ky thuat" lay tu bang specifications.

Bang nhieu cot tren trang (bang tra ma tin hieu dau vao cua cac bo hien thi)
khong phai thong so cua san pham -> de rieng thanh khoi "Bang tra cuu".
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

BRAND = "Supmea"
BRAND_FULL = "Supmea Automation"
SRC = os.path.join(HERE, "out", "12 Supmea", "content.json")
OUT_DIR = out_dir.brand_dir("Noi dung Supmea")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
FLD_FILL = PatternFill("solid", fgColor="DDEBF7")
GRP_FILL = PatternFill("solid", fgColor="F2F2F2")
WRAP = Alignment(wrap_text=True, vertical="top")
MAX_TAGS = 12


def safe(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s or "").strip()[:n].rstrip(" .")


def cau_dau(s):
    """Cau dau tien - de mo ta ngan khong chep y nguyen mo ta dai."""
    m = re.search(r"\.\s+(?=[A-ZÀ-Ỹ])", s or "")
    return (s[:m.start() + 1] if m else (s or "")).strip()


def ctx(a):
    en = a.get("title_en") or a["name"]
    nv = a.get("name_vi") or en
    first = nv.split(" ")[0]
    cv = {"type_c": nv[:1].upper() + nv[1:],
          "type_l": nv if (len(first) > 1 and first.isupper())
                    else nv[:1].lower() + nv[1:],
          "name": nv, "name_brand": BRAND, "brand": BRAND}
    ce = {"type_c": en, "type_l": en.lower(),
          "name": en, "name_brand": BRAND, "brand": BRAND}
    return cv, ce


def tag_text(a, name):
    """The tag = [Ma san pham], [Thuong hieu]. Trang nao khong ghi ma trong
    bang thong so thi lay ten san pham thay vao."""
    codes = a.get("codes") or []
    head = ", ".join(codes[:MAX_TAGS]) if codes else name
    return "%s, %s" % (head, BRAND)


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

    def block(hdr):
        ws.append(hdr)
        r = ws.max_row
        ws.cell(row=r, column=1).fill = FLD_FILL
        ws.cell(row=r, column=1).font = Font(bold=True)
        for c in range(2, len(hdr) + 1):
            ws.cell(row=r, column=c).font = Font(bold=True, italic=True)
            ws.cell(row=r, column=c).alignment = WRAP

    kw_vi = "%s %s" % (nv, BRAND)
    kw_en = "%s %s" % (en, BRAND)
    row("Ảnh sản phẩm", a.get("img_file") or "(chưa có ảnh)")
    row("Danh mục sản phẩm", a.get("cat_vi") or "(chưa phân loại)")
    row("Thẻ tag", tag_text(a, nv), tag_text(a, en))
    row("Thương hiệu", BRAND_FULL, BRAND_FULL)
    row("URL", slugify(kw_vi), slugify(kw_en))
    row("Keyword chính", kw_vi, kw_en)
    row("Meta title", "%s %s" % (cv["type_c"], BRAND), kw_en)
    row("Meta description (155 - 165 characters)",
        META_TPL.format(**cv), META_TPL_EN.format(**ce))

    long_vi, long_en = a.get("long_vi") or [], a.get("long_en") or []
    short_vi, short_en = a.get("short_vi") or [], a.get("short_en") or []
    if not long_vi and not short_vi:
        # Trang hang khong co chu nao gioi thieu -> dung mau chung.
        dv, de = DESCDESC.format(**cv), DESCDESC_EN.format(**ce)
        row("Mô tả dài của sản phẩm", dv, de)
        row("Mô tả ngắn của sản phẩm", cau_dau(dv), cau_dau(de))
    else:
        row("Mô tả dài của sản phẩm", "\n".join(long_vi), "\n".join(long_en))
        row("Mô tả ngắn của sản phẩm",
            "\n".join("- " + x for x in short_vi),
            "\n".join("- " + x for x in short_en))
    row("Mô tả thông số", SPECDESC.format(**cv), SPECDESC_EN.format(**ce))

    # ---- Bang thong so ky thuat ----
    specs = a.get("specs_en") or []
    if specs:
        block(["Thông số kỹ thuật", "Tên thông số", "Giá trị",
               "Parameter (EN)", "Value (EN)"])
        sv = a.get("specs_vi") or specs
        for (ke, ve), (kv, vv) in zip(specs, sv):
            if ke == "#":                 # dong tieu de nhom trong bang goc
                ws.append(["", vv, "", ve, ""])
                for c in (2, 4):
                    ws.cell(row=ws.max_row, column=c).font = Font(bold=True)
                for c in range(2, 6):
                    ws.cell(row=ws.max_row, column=c).fill = GRP_FILL
            else:
                ws.append(["", kv, vv, ke, ve])
            for c in range(2, 6):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP
    else:
        block(["Thông số kỹ thuật", "(trang hãng không có bảng thông số)"])

    # ---- Bang tra cuu nhieu cot (giu nguyen dang luoi cua trang hang) ----
    extra = a.get("extra_en") or []
    ev = a.get("extra_vi") or extra
    for n, (t_en, t_vi) in enumerate(zip(extra, ev), 1):
        ncol = len(t_en["head"])
        block(["Bảng tra cứu %d" % n if len(extra) > 1 else "Bảng tra cứu"]
              + list(t_vi["head"]) + ["%s (EN)" % c for c in t_en["head"]])
        for r_en, r_vi in zip(t_en["rows"], t_vi["rows"]):
            line = [""] + (list(r_vi) + [""] * ncol)[:ncol] \
                        + (list(r_en) + [""] * ncol)[:ncol]
            ws.append(line)
            for c in range(2, len(line) + 1):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP

    ws.append([])
    ws.append(["Nhánh danh mục", " > ".join(a.get("path") or []), ""])
    ws.append(["Mã sản phẩm", a.get("series", ""), ""])
    ws.append(["Nguồn", a.get("url_src", ""), ""])
    ncol = max([len(t["head"]) for t in extra] + [2])
    for i, w in enumerate([34, 40, 46] + [24] * (ncol * 2), 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    fname = "%03d - %s.xlsx" % (idx, safe(en))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_index(arts, names, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "MUC LUC"
    ws["A1"] = "MỤC LỤC BÀI ĐĂNG - 12 Supmea"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["STT", "Sản phẩm", "Tên tiếng Việt", "Mã sản phẩm", "File",
               "Ảnh", "Số thông số", "Số dòng mô tả dài", "Danh mục Ntrend",
               "Nhánh danh mục", "Nguồn"])
    for c in range(1, 12):
        ws.cell(row=3, column=c).fill = FLD_FILL
        ws.cell(row=3, column=c).font = Font(bold=True)
    for i, (a, f) in enumerate(zip(arts, names), 1):
        ws.append([i, a.get("title_en") or a["name"], a.get("name_vi", ""),
                   ", ".join(a.get("codes") or []), f,
                   "x" if a.get("img_file") else "",
                   sum(1 for k, _v in (a.get("specs_en") or []) if k != "#"),
                   len(a.get("long_en") or []),
                   a.get("cat_vi", ""), " > ".join(a.get("path") or []),
                   a.get("url_src", "")])
    for i, w in enumerate([6, 44, 40, 18, 48, 6, 12, 16, 50, 46, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    # Bai nao khong co anh thi lay logo hang (xem logo_fallback.py).
    logo_fallback.run("supmea")
    arts = json.load(io.open(SRC, encoding="utf-8"))["articles"]
    os.makedirs(OUT_DIR, exist_ok=True)
    names = [build_one(a, i, OUT_DIR) for i, a in enumerate(arts, 1)]
    build_index(arts, names, OUT_DIR)
    prune(OUT_DIR, names)
    print("Saved %d bai (%d bai co bang thong so) -> %s"
          % (len(arts), sum(1 for a in arts if a.get("specs_en")), OUT_DIR))
    for lbl, bad in (("khong co bang thong so",
                      [a for a in arts if not a.get("specs_en")]),
                     ("khong co mo ta",
                      [a for a in arts if not a.get("long_en")
                       and not a.get("short_en")]),
                     ("khong co anh", [a for a in arts if not a.get("img_file")])):
        if bad:
            print("  %s (%d): %s" % (lbl, len(bad),
                  ", ".join((x.get("title_en") or x["name"])[:28] for x in bad[:8])))


if __name__ == "__main__":
    main()
