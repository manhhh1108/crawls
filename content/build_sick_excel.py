# -*- coding: utf-8 -*-
"""SICK - PHA 3: content.json -> moi bai 1 file Excel trong data/Noi dung SICK/.

Giu dung layout da duyet o Perma: cot A = truong, cot B = TIENG VIET (ban dang
web), cot C = tieng Anh (de soat chinh ta ban dich). Mau van theo
'Crawl noi dung Ntrend.docx' + 'ntrend template.xlsx'.

Them so voi Perma:
  - "Danh muc san pham" dien san tu cot "Danh muc SP Tieng Viet" cua file thong ke
  - co CA bang thong so ky thuat CA bang ma san pham
  - sinh them "00 CAN SOAT DICH.xlsx" = cac cau may dich con dang ngo
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
from openpyxl.styles import Font, Alignment, PatternFill   # noqa: E402  # noqa
from openpyxl.utils import get_column_letter               # noqa: E402
from openpyxl.drawing.image import Image as XImage         # noqa: E402

from build_content_excel import (LONG_TPL_VI, LONG_TPL_EN, SHORT_TPL_VI,       # noqa: E402
                                 SHORT_TPL_EN, SPECDESC_TPL_VI, SPECDESC_TPL_EN,
                                 HDR_FILL, FIELD_FILL, HDR_FONT, FIELD_FONT, WRAP)

BRAND_FULL_VI = "SICK (Đức)"
BRAND_FULL_EN = "SICK (Germany)"
MAX_ITEM_ROWS = 60          # so dong toi da cua bang ma san pham trong 1 bai

NOSPEC_DESC_VI = ("{type_c} {name_brand} là sản phẩm công nghiệp chính hãng của "
                  + BRAND_FULL_VI + "; bảng mã sản phẩm chi tiết như bảng dưới:")
NOSPEC_DESC_EN = ("The {name_brand} {type_l} is a genuine industrial product from "
                  + BRAND_FULL_EN + "; the detailed part-number table is below:")
META_TPL_VI = ("Mua {type_l} {name_brand} chính hãng – thông số kỹ thuật đầy đủ, "
               "giá tốt, hỗ trợ tư vấn và lựa chọn model phù hợp. "
               "Liên hệ Ntrend ngay để được tư vấn!")
META_TPL_EN = ("Buy genuine {name_brand} {type_l} – full specifications, good "
               "price, support in selecting the right model. "
               "Contact Ntrend now for a consultation!")


def with_brand(a):
    return a["name"] if a["brand"].lower() in a["name"].lower() \
        else "%s %s" % (a["name"], a["brand"])


def singular(s):
    """'safety controllers' -> 'safety controller' (nhom ho SP cua SICK deu so nhieu)."""
    w = (s or "").split()
    if w and len(w[-1]) > 3 and w[-1].endswith("s") and not w[-1].endswith(("ss", "us", "is")):
        w[-1] = w[-1][:-1] if not w[-1].endswith("ies") else w[-1][:-3] + "y"
    return " ".join(w)


def ctx(a):
    t = singular(a["type_en"] or "sensor")
    return {"name": a["name"], "name_brand": with_brand(a), "brand": a["brand"],
            "type_c": (a["type_vi"] or "Sản phẩm")[:1].upper() + (a["type_vi"] or "Sản phẩm")[1:],
            "type_l": t.lower()}


def ctx_vi(a):
    c = ctx(a)
    c["type_l"] = (a["type_vi"] or "sản phẩm").lower()
    return c


def safe_name(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s).strip()[:n].rstrip(" .")


def add_row(ws, label, vn, en):
    ws.append([label, vn, en])
    r = ws.max_row
    ws.cell(row=r, column=1).fill = FIELD_FILL
    ws.cell(row=r, column=1).font = FIELD_FONT
    for c in (1, 2, 3):
        ws.cell(row=r, column=c).alignment = WRAP
    return r


def table_header(ws, label, cols):
    ws.append([label] + cols)
    r = ws.max_row
    ws.cell(row=r, column=1).fill = FIELD_FILL
    ws.cell(row=r, column=1).font = FIELD_FONT
    for c in range(2, len(cols) + 2):
        ws.cell(row=r, column=c).font = Font(bold=True, italic=True)
    return r


def build_article(a, i, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "NOI DUNG"
    ws["A1"] = "NỘI DUNG SẢN PHẨM NTREND - %s" % a["name"]
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["Trường", "Tiếng Việt (bản đăng web)", "Tiếng Anh (soát chính tả)"])
    hr = ws.max_row
    for c in range(1, 4):
        ws.cell(row=hr, column=c).fill = HDR_FILL
        ws.cell(row=hr, column=c).font = HDR_FONT

    cv, ce = ctx_vi(a), ctx(a)
    has_specs = bool(a.get("specs_en"))
    slug = a["slug"] if a["brand"].lower() in a["slug"].lower() \
        else "%s-%s" % (a["slug"], a["brand"].lower())

    if a.get("overview_vi"):
        long_vi, long_en = "\n\n".join(a["overview_vi"]), "\n\n".join(a["overview_en"])
    else:
        long_vi, long_en = LONG_TPL_VI.format(**cv), LONG_TPL_EN.format(**ce)
    if a.get("features_vi"):
        short_vi = "\n".join("- " + x for x in a["features_vi"])
        short_en = "\n".join("- " + x for x in a["features_en"])
    else:
        short_vi, short_en = SHORT_TPL_VI.format(**cv), SHORT_TPL_EN.format(**ce)
    if has_specs:
        sd_vi, sd_en = SPECDESC_TPL_VI.format(**cv), SPECDESC_TPL_EN.format(**ce)
    elif a.get("items"):
        sd_vi, sd_en = NOSPEC_DESC_VI.format(**cv), NOSPEC_DESC_EN.format(**ce)
    else:
        sd_vi = ("{type_c} {name_brand} là sản phẩm công nghiệp chính hãng của "
                 + BRAND_FULL_VI + ".").format(**cv)
        sd_en = ("The {name_brand} {type_l} is a genuine industrial product from "
                 + BRAND_FULL_EN + ".").format(**ce)

    add_row(ws, "Ảnh sản phẩm",
            os.path.basename(a["img_file"]) if a.get("img_file") else "(chưa có ảnh)", "")
    add_row(ws, "Danh mục sản phẩm", a.get("cat_vi") or "(chưa phân loại)",
            a.get("cat_en") or " > ".join(a.get("path", [])[1:]))
    add_row(ws, "Thẻ tag", "%s, %s" % (a["name"], a["brand"]),
            "%s, %s" % (a["name"], a["brand"]))
    add_row(ws, "Thương hiệu", a.get("brand_field") or a["brand"],
            a.get("brand_field") or a["brand"])
    add_row(ws, "URL", slug, slug)
    add_row(ws, "Keyword chính", with_brand(a), with_brand(a))
    add_row(ws, "Meta title", "%s %s" % (cv["type_c"], with_brand(a)),
            "%s %s" % (with_brand(a), ce["type_l"]))
    add_row(ws, "Meta description (155 - 165 characters)",
            META_TPL_VI.format(**cv), META_TPL_EN.format(**ce))
    if a.get("slogan_vi"):
        add_row(ws, "Slogan (hãng)", a["slogan_vi"], a.get("slogan_en", ""))
    add_row(ws, "Mô tả dài của sản phẩm", long_vi, long_en)
    add_row(ws, "Mô tả ngắn của sản phẩm", short_vi, short_en)
    if a.get("applications_vi"):
        add_row(ws, "Lĩnh vực ứng dụng", "; ".join(a["applications_vi"]),
                "; ".join(a["applications_en"]))
    add_row(ws, "Mô tả thông số", sd_vi, sd_en)

    if has_specs:
        table_header(ws, "Thông số chi tiết",
                     ["Tên thông số", "Giá trị", "Parameter (EN)", "Value (EN)"])
        for (ke, ve), (kv, vv) in zip(a["specs_en"], a["specs_vi"]):
            ws.append(["", kv, vv, ke, ve])
            for c in range(2, 6):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP

    items = a.get("items") or []
    if items:
        shown = items[:MAX_ITEM_ROWS]
        label = "Bảng mã sản phẩm"
        if len(items) > len(shown):
            label += " (%d/%d model)" % (len(shown), len(items))
        table_header(ws, label, ["Model", "Mã đặt hàng (Part No.)"])
        for it in shown:
            ws.append(["", it["name"], it["art"]])
        if len(items) > len(shown):
            ws.append(["", "... còn %d model khác, xem đầy đủ tại trang nguồn"
                       % (len(items) - len(shown)), ""])

    if not has_specs and not items:
        table_header(ws, "Bảng mã sản phẩm", ["(trang hãng không có dữ liệu)"])

    ws.append([])
    if a.get("not_on_web"):
        ws.append(["CẢNH BÁO",
                   "Dòng này không còn trên sick.com/sg/en (đã ngừng hoặc đổi tên) – "
                   "nội dung đang dùng mẫu văn, cần kiểm tra lại trước khi đăng", ""])
        ws.cell(row=ws.max_row, column=1).fill = PatternFill("solid", fgColor="FFC7CE")
        ws.cell(row=ws.max_row, column=1).font = Font(bold=True, color="9C0006")
        ws.cell(row=ws.max_row, column=2).alignment = WRAP
    ws.append(["Nguồn", a["url_src"] or "(không còn trang trên web SICK)", ""])
    ws.append(["Mã danh mục SICK", a.get("code", ""), ""])
    for i2, w in enumerate([34, 62, 62, 34, 46], 1):
        ws.column_dimensions[get_column_letter(i2)].width = w

    if a.get("img_file") and os.path.exists(a["img_file"]):
        try:
            img = XImage(a["img_file"])
            img.width, img.height = 170, 170
            ws.add_image(img, "E4")
        except Exception:
            pass

    fname = "%03d - %s.xlsx" % (i, safe_name(a["name"]))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_review(out_dir, src_dir):
    p = os.path.join(src_dir, "review_dich.json")
    if not os.path.exists(p):
        return 0
    with io.open(p, encoding="utf-8") as f:
        rows = json.load(f)
    wb = Workbook()
    ws = wb.active
    ws.title = "CAN SOAT"
    ws["A1"] = "CÁC CÂU DỊCH CẦN SOÁT TAY - SICK"
    ws["A1"].font = Font(bold=True, size=13, color="C00000")
    ws.append([])
    ws.append(["Dòng sản phẩm", "Trường", "Lý do nghi ngờ", "Tiếng Anh", "Bản dịch"])
    for c in range(1, 6):
        ws.cell(row=3, column=c).fill = HDR_FILL
        ws.cell(row=3, column=c).font = HDR_FONT
    for r in rows:
        ws.append(r)
        for c in range(1, 6):
            ws.cell(row=ws.max_row, column=c).alignment = WRAP
    for i, w in enumerate([26, 20, 26, 70, 70], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 CAN SOAT DICH.xlsx"))
    return len(rows)


def build(src_dir, out_dir):
    with io.open(os.path.join(src_dir, "content.json"), encoding="utf-8") as f:
        arts = json.load(f)["articles"]
    os.makedirs(out_dir, exist_ok=True)

    idx = Workbook()
    iw = idx.active
    iw.title = "MUC LUC"
    iw["A1"] = "MỤC LỤC BÀI ĐĂNG - 01 SICK"
    iw["A1"].font = Font(bold=True, size=13, color="1F4E79")
    iw.append([])
    iw.append(["STT", "Dòng sản phẩm", "File", "Ảnh", "Có TSKT", "Số model",
               "Danh mục Ntrend (VI)", "Nguồn"])
    for c in range(1, 9):
        iw.cell(row=3, column=c).fill = HDR_FILL
        iw.cell(row=3, column=c).font = HDR_FONT
    for i, w in enumerate([6, 34, 50, 6, 9, 10, 52, 62], 1):
        iw.column_dimensions[get_column_letter(i)].width = w
    iw.freeze_panes = "A4"

    for i, a in enumerate(arts, 1):
        fname = build_article(a, i, out_dir)
        iw.append([i, a["name"], fname, "x" if a.get("img_file") else "",
                   "x" if a.get("specs_en") else "", len(a.get("items") or []),
                   a.get("cat_vi", ""), a["url_src"]])
    idx.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))
    n = build_review(out_dir, src_dir)
    print("Saved %d bai -> %s" % (len(arts), out_dir))
    if n:
        print("  + 00 CAN SOAT DICH.xlsx (%d dong can soat tay)" % n)


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    build(os.path.join(HERE, "out", "01 SICK"),
          os.path.join(REPO, "data", "Noi dung SICK"))
