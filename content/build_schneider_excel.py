# -*- coding: utf-8 -*-
"""Schneider Electric - PHA 3: content.json -> moi bai 1 file Excel.

    python build_schneider_excel.py

Layout theo 'Template nội dung Schneider.docx': 12 dong truong + mot bang
"Model | Mã đặt hàng (Part No.)".

se.com khong tach "model" voi "ma dat hang": moi san pham chi co MOT commercial
reference (MGL36800) kem mot cau mo ta ("Circuit breaker, PowerPacT M, 800A,
3 pole..."). Nen cot Model lay cau mo ta do - giong cach bang ma SP cua Festo
dat ten san pham o cot Model - con cot "Ma dat hang" lay commercial reference.
Them mot cot "Mo ta (VI)" cho ban tieng Viet cua cau do.
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

BRAND = "Schneider Electric"
SRC = os.path.join(HERE, "out", "11 Schneider", "content.json")
OUT_DIR = out_dir.brand_dir("Noi dung Schneider")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
FLD_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")
MAX_TAGS = 12
MAX_ROWS = 300


def safe(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s or "").strip()[:n].rstrip(" .")


def cau_dau(s):
    m = re.search(r"\.\s+(?=[A-ZÀ-Ỹ])", s or "")
    return (s[:m.start() + 1] if m else (s or "")).strip()


DAU_DONG = re.compile("^[\u2022\u25cf\u25aa\u00b7*\\-\u2013]+\\s*")


def gach(x):
    """Mot dong mo ta dai -> mot gach dau dong.

    Van ban cua hang da co san dau "\u2022" o cac y con. Them "- " may moc nua
    thi ra "- \u2022 5 current ranges from 0.1 to 110A" - hai dau dong chong
    nhau. Doi dau cua hang thanh "-" de ca o dong nhat.
    """
    return "- " + DAU_DONG.sub("", x or "").strip()


def thuong(s):
    """Ha thuong cau tieng Anh nhung GIU NGUYEN ten dong SP.

    `en.lower()` bien "SpaceLogic Central I/O Module" thanh "spacelogic...",
    "EcoFit" thanh "ecofit" - doc ra nhu loi chinh ta ngay trong o giao khach.
    Chi ha thuong tu viet hoa chu dau thuong; tu co chu hoa ben trong hoac co
    so (SpaceLogic, PowerPacT, TeSys, LR9D) giu nguyen.
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
    # Chu dau chi ha thuong khi tu dau la tu thuong: "PowerPacT" khong phai
    # VIET HOA HET nen luat cu ha thanh "powerPacT".
    cv = {"type_c": nv[:1].upper() + nv[1:],
          "type_l": nv[:1].lower() + nv[1:] if first[1:].islower() else nv,
          "name": nv, "name_brand": BRAND, "brand": BRAND}
    ce = {"type_c": en, "type_l": thuong(en),
          "name": en, "name_brand": BRAND, "brand": BRAND}
    return cv, ce


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
    row("Thương hiệu", BRAND, BRAND)
    row("URL", slugify(kw_vi), slugify(kw_en))
    row("Keyword chính", kw_vi, kw_en)
    row("Meta title", "%s %s" % (cv["type_c"], BRAND), kw_en)
    row("Meta description (155 - 165 characters)",
        META_TPL.format(**cv), META_TPL_EN.format(**ce))

    # Mo ta ngan dung DU ca hai cau cua mau (review khach): cat con cau dau
    # la mat nua thong tin, ma trang hang khong co cau nao khac bu vao.
    dv, de = DESCDESC.format(**cv), DESCDESC_EN.format(**ce)
    long_vi, long_en = a.get("long_vi") or [], a.get("long_en") or []
    if long_en:
        row("Mô tả dài của sản phẩm",
            "\n".join(gach(x) for x in long_vi),
            "\n".join(gach(x) for x in long_en))
        row("Mô tả ngắn của sản phẩm", dv, de)
    else:
        # Trang hang khong co phan Overview -> dung mau chung.
        row("Mô tả dài của sản phẩm", dv, de)
        row("Mô tả ngắn của sản phẩm", dv, de)
    row("Mô tả thông số", SPECDESC.format(**cv), SPECDESC_EN.format(**ce))

    items = a.get("items") or []
    if items:
        block(["Bảng mã sản phẩm + Thông tin sản phẩm", "Model",
               "Mã đặt hàng (Part No.)", "Mô tả (VI)"])
        for it in items[:MAX_ROWS]:
            ws.append(["", it.get("name", ""), it.get("art", ""),
                       it.get("desc_vi") or it.get("name", "")])
            for c in range(2, 5):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP
        # So mã hãng đăng (total) có thể lớn hơn số mã đã cào (MAX_ITEMS) và
        # lớn hơn số dòng in được (MAX_ROWS) - ghi theo con số của hãng, đừng
        # trừ theo số đã cào, kẻo báo thiếu ít hơn thực tế.
        printed = min(len(items), MAX_ROWS)
        total = max(a.get("total") or 0, len(items))
        if total > printed:
            ws.append(["", "(dòng này hãng đăng %d mã, bảng in %d mã đầu)"
                       % (total, printed)])
    else:
        block(["Bảng mã sản phẩm + Thông tin sản phẩm",
               "(trang hãng không đăng mã SP cho dòng này)"])

    ws.append([])
    ws.append(["Nhánh danh mục", " > ".join(a.get("path") or []), ""])
    ws.append(["Nguồn", a.get("url_src", ""), ""])
    for i, w in enumerate([34, 56, 22, 56], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    fname = "%03d - %s.xlsx" % (idx, safe(en))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_index(arts, names, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "MUC LUC"
    ws["A1"] = "MỤC LỤC BÀI ĐĂNG - 11 Schneider"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["STT", "Dòng sản phẩm", "Tên tiếng Việt", "File", "Ảnh",
               "Số mã SP", "Tổng mã trên web", "Số dòng mô tả dài",
               "Danh mục Ntrend", "Nhánh danh mục", "Nguồn"])
    for c in range(1, 12):
        ws.cell(row=3, column=c).fill = FLD_FILL
        ws.cell(row=3, column=c).font = Font(bold=True)
    for i, (a, f) in enumerate(zip(arts, names), 1):
        ws.append([i, a.get("title_en") or a["name"], a.get("name_vi", ""), f,
                   "x" if a.get("img_file") else "",
                   len(a.get("items") or []), a.get("total") or 0,
                   len(a.get("long_en") or []), a.get("cat_vi", ""),
                   " > ".join(a.get("path") or []), a.get("url_src", "")])
    for i, w in enumerate([6, 44, 44, 48, 6, 10, 14, 16, 50, 50, 56], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    logo_fallback.run("schneider")
    arts = json.load(io.open(SRC, encoding="utf-8"))["articles"]
    os.makedirs(OUT_DIR, exist_ok=True)
    names = [build_one(a, i, OUT_DIR) for i, a in enumerate(arts, 1)]
    build_index(arts, names, OUT_DIR)
    prune(OUT_DIR, names)
    nit = sum(len(a.get("items") or []) for a in arts)
    print("Saved %d bai (%d ma SP) -> %s" % (len(arts), nit, OUT_DIR))
    for lbl, bad in (("khong co ma SP", [a for a in arts if not a.get("items")]),
                     ("khong co mo ta dai",
                      [a for a in arts if not a.get("long_en")]),
                     ("khong co anh", [a for a in arts if not a.get("img_file")])):
        if bad:
            print("  %s (%d): %s" % (lbl, len(bad),
                  ", ".join((x.get("title_en") or x["name"])[:26] for x in bad[:8])))


if __name__ == "__main__":
    main()
