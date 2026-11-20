# -*- coding: utf-8 -*-
"""Xingguang / Starshine Drive - PHA 3: content.json -> moi bai 1 file Excel.

    python build_xingguang_excel.py

Layout theo 'Template nội dung Xingguang-Starshine.docx'. O "Bang ma san pham"
trong template de TRONG, va dung la trang hang khong dang ma SP nao - moi trang
la mot DONG SP, ma dong nam ngay trong ten bai (SVF, NCJ, SHMF...). Nen o do
thay bang khoi "Thong so ky thuat" lam tu cac dong "ten: gia tri" o dau phan
mo ta (Output speed, Input power, Input).
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
from openpyxl.drawing.image import Image as XLImage        # noqa: E402
from openpyxl.styles import Font, Alignment, PatternFill   # noqa: E402
from openpyxl.utils import get_column_letter               # noqa: E402

import out_dir                                         # noqa: E402
import logo_fallback                                  # noqa: E402
from bai_prune import prune                             # noqa: E402
from c_pepperl_try import (slugify, SPECDESC, META_TPL,    # noqa: E402
                           SPECDESC_EN, META_TPL_EN,
                           DESCDESC, DESCDESC_EN)

BRAND = "Starshine Drive"
BRAND_FULL = "Guangdong Xingguang Transmission (Starshine Drive)"
SRC = os.path.join(HERE, "out", "14 Xingguang-Starshine", "content.json")
OUT_DIR = out_dir.brand_dir("Noi dung Xingguang-Starshine")
IMG_DIR = os.path.join(OUT_DIR, "images")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
FLD_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")
MAX_TAGS = 12


def safe(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s or "").strip()[:n].rstrip(" .")


def thuong(s):
    """Ha thuong cau tieng Anh nhung GIU NGUYEN ma dong SP.

    `en.lower()` bien "SNFEG3000 GEAR REDUCER" thanh "snfeg3000 gear reducer",
    "XGK Series" thanh "xgk series" - doc ra nhu loi chinh ta ngay trong o
    giao khach. Chi ha thuong tu viet hoa chu dau thuong; tu co chu hoa ben
    trong hoac co so (SNFEG3000, XGK, SMQ340) giu nguyen.
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
    # Chu dau chi ha thuong khi tu dau la tu thuong, khong thi ma dong
    # ("SNFEG3000") bi ha thanh "sNFEG3000".
    cv = {"type_c": nv[:1].upper() + nv[1:],
          "type_l": nv[:1].lower() + nv[1:] if first[1:].islower() else nv,
          "name": nv, "name_brand": BRAND, "brand": BRAND}
    ce = {"type_c": en, "type_l": thuong(en),
          "name": en, "name_brand": BRAND, "brand": BRAND}
    return cv, ce


def than_bai(lines, bul):
    """Mot o mo ta dai: dong nao la gach dau dong thi them "- " o dau.

    Khach yeu cau giu dung cach trinh bay cua trang - cac muc danh so "1. 2."
    thanh gach dau dong, con doan ket luan ("Do do, thiet bi co the duoc su
    dung...") de nguyen la mot doan van (review 02/11).
    """
    return "\n".join(("- " + t) if i < len(bul) and bul[i] else t
                      for i, t in enumerate(lines))


def tag_text(a, name):
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
    bul = a.get("long_bul") or []
    if not long_vi and not short_vi:
        # Mo ta ngan dung DU ca hai cau cua mau - trang hang khong co chu nao
        # khac bu vao nua cau bi cat.
        dv, de = DESCDESC.format(**cv), DESCDESC_EN.format(**ce)
        row("Mô tả dài của sản phẩm", dv, de)
        row("Mô tả ngắn của sản phẩm", dv, de)
    else:
        row("Mô tả dài của sản phẩm",
            than_bai(long_vi, bul), than_bai(long_en, bul))
        row("Mô tả ngắn của sản phẩm",
            "\n".join("- " + x for x in short_vi),
            "\n".join("- " + x for x in short_en))
    row("Mô tả thông số", SPECDESC.format(**cv), SPECDESC_EN.format(**ce))

    specs = a.get("specs_en") or []
    if specs:
        block(["Thông số kỹ thuật", "Tên thông số", "Giá trị",
               "Parameter (EN)", "Value (EN)"])
        sv = a.get("specs_vi") or specs
        for (ke, ve), (kv, vv) in zip(specs, sv):
            ws.append(["", kv, vv, ke, ve])
            for c in range(2, 6):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP
    else:
        block(["Thông số kỹ thuật", "(trang hãng không có bảng thông số dạng chữ)"])
    # Bang thong so day du cua hang CHI CO o dang anh. Khach xin tai ve kem
    # bai de dang thang len web (review 02/11), nen vua ghi ten file vua nhung
    # anh xuong duoi de soat cho nhanh.
    si = a.get("spec_img_file") or ""
    row("Ảnh bảng thông số (hãng chỉ đăng dạng ảnh)",
        si or "(trang hãng không có)", a.get("spec_img_url") or "")
    # Template de trong o nay vi trang hang khong dang ma dat hang nao - moi
    # trang la mot dong SP, ma dong da nam o the tag.
    block(["Bảng mã sản phẩm + Thông tin sản phẩm",
           "(trang hãng không đăng mã đặt hàng - mỗi trang là một dòng SP)"])

    ws.append([])
    ws.append(["Nhánh danh mục", " > ".join(a.get("path") or []), ""])
    ws.append(["Mã dòng sản phẩm", ", ".join(a.get("codes") or []), ""])
    ws.append(["Nguồn", a.get("url_src", ""), ""])
    for i, w in enumerate([34, 46, 50, 34, 40], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    p = os.path.join(IMG_DIR, si)
    if si and os.path.exists(p):
        ws.append([])
        r = ws.max_row + 1
        ws.cell(row=r, column=1).value = "(ảnh bảng thông số – file %s)" % si
        ws.cell(row=r, column=1).font = Font(bold=True)
        ws.add_image(XLImage(p), "B%d" % r)
    fname = "%02d - %s.xlsx" % (idx, safe(en))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_index(arts, names, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "MUC LUC"
    ws["A1"] = "MỤC LỤC BÀI ĐĂNG - 14 Xingguang-Starshine"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["STT", "Dòng sản phẩm", "Tên tiếng Việt", "Mã dòng", "File",
               "Ảnh", "Ảnh bảng TS", "Số thông số", "Số dòng mô tả dài",
               "Danh mục Ntrend", "Nhánh danh mục", "Nguồn"])
    for c in range(1, 13):
        ws.cell(row=3, column=c).fill = FLD_FILL
        ws.cell(row=3, column=c).font = Font(bold=True)
    for i, (a, f) in enumerate(zip(arts, names), 1):
        ws.append([i, a.get("title_en") or a["name"], a.get("name_vi", ""),
                   ", ".join(a.get("codes") or []), f,
                   "x" if a.get("img_file") else "",
                   a.get("spec_img_file", ""),
                   len(a.get("specs_en") or []), len(a.get("long_en") or []),
                   a.get("cat_vi", ""), " > ".join(a.get("path") or []),
                   a.get("url_src", "")])
    for i, w in enumerate([6, 44, 42, 14, 48, 6, 40, 12, 16, 50, 48, 56], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    logo_fallback.run("xingguang")
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
                      [a for a in arts if not a.get("long_en")]),
                     ("khong co anh", [a for a in arts if not a.get("img_file")]),
                     ("khong co anh bang thong so",
                      [a for a in arts if not a.get("spec_img_file")])):
        if bad:
            print("  %s (%d): %s" % (lbl, len(bad),
                  ", ".join((x.get("title_en") or x["name"])[:28] for x in bad[:8])))


if __name__ == "__main__":
    main()
