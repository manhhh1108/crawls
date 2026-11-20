# -*- coding: utf-8 -*-
"""Pilz - PHA 3: content.json -> moi bai 1 file Excel.

    python build_pilz_excel.py

Layout theo 'Template nội dung Pilz.docx': 12 dong truong + mot bang 3 cot
"Model | Mã đặt hàng (Part No.) | Mô tả".

Template khong co khoi "Thong so ky thuat" rieng: o o mo ta dai no ghi "Lấy
full thông số/đặc điểm của sản phẩm... (Loại các thông số trùng tên dòng)".
Nen thong so (lay tu khoi bo loc cua trang muc) duoc viet thang thanh cac gach
dau dong trong o mo ta dai, khong lam bang rieng - khong thi bai bi lap hai
lan cung mot noi dung.
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

BRAND = "Pilz"
SRC = os.path.join(HERE, "out", "09 Pilz", "content.json")
OUT_DIR = out_dir.brand_dir("Noi dung Pilz")

HDR_FILL = PatternFill("solid", fgColor="1F4E79")
FLD_FILL = PatternFill("solid", fgColor="DDEBF7")
WRAP = Alignment(wrap_text=True, vertical="top")
MAX_TAGS = 12
MAX_ROWS = 300          # so dong ma SP in ra moi bai


def safe(s, n=58):
    return re.sub(r'[\\/:*?"<>|]+', "-", s or "").strip()[:n].rstrip(" .")


MA_DONG = re.compile(r"[A-Z]{2,}|[A-Za-z]+[0-9]")


def ten_dong(a):
    """Ten dong SP, lay tu muc CHA.

    Khach review doi mo ta ngan dung mau "[Loai SP] [Ten dong SP] ...".
    56/180 muc Pilz dat ten thuan loai ("Two-hand monitoring", "Accessoires")
    nen doc mot minh khong biet thuoc dong nao; muc cha moi co ten dong:
    "PNOZsigma - safety relays (77)" -> "PNOZsigma".

    Chi nhan khi phan dau cua cha trong nhu MA DONG. Cha kieu "Cables and
    plug-in connectors" chi la mot ten loai nua - ghep vao cau tieng Viet
    thanh cau nua Anh nua Viet ma khong them thong tin gi.
    """
    path = a.get("path") or []
    if len(path) < 2:
        return ""
    cha = re.sub(r"\s*\(\d+\)\s*$", "", path[-2] or "").strip()
    cha = cha.split(" - ")[0].strip()
    return cha if MA_DONG.search(cha) else ""


def chu_de(ten, dong):
    """Ghep "[Loai SP] [Ten dong SP]", bo qua neu ten da chua ten dong."""
    if not dong or dong.lower() in (ten or "").lower():
        return ten
    return "%s %s" % (ten, dong)


def thuong(s):
    """Ha thuong cau tieng Anh nhung GIU NGUYEN ma dong SP.

    `en.lower()` bien "PSSuniversal - base modules" thanh "pssuniversal..." -
    doc ra nhu loi chinh ta. Chi ha thuong tu viet hoa chu dau thuong, giu
    nguyen tu co chu hoa ben trong hoac co so (PSSuniversal, PNOZ, PDP67).
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
    # Ha thuong chu dau chi khi tu dau la tu thuong. "PSSuniversal" khong
    # phai VIET HOA HET nen luat cu ha thanh "pSSuniversal" - loi chinh ta
    # ngay o Meta description.
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


def spec_lines(a, line_en):
    """Thong so -> (gach dau dong VI, gach dau dong EN).

    Loc theo BAN TIENG ANH roi lay dong tieng Viet CUNG CHI SO - loc rieng hai
    ban thi hai cot lech nhau, ten thong so cua dong nay ghep voi gia tri cua
    dong khac.

    Template ghi ro "(Loại các thông số trùng tên dòng)": muc "Plug-in
    connectors" co bo loc "Type of general accessories: Connector" - noi lai
    chinh ten muc, in ra khong them gi.
    """
    se = a.get("specs_en") or []
    sv = a.get("specs_vi") or se
    low = (line_en or "").lower()
    vi, en = [], []
    for n, (k, v) in enumerate(se):
        if not v or (v.lower() in low and k.lower() in low):
            continue
        kv, vv = sv[n] if n < len(sv) else (k, v)
        vi.append("%s: %s" % (kv, vv))
        en.append("%s: %s" % (k, v))
    return vi, en


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

    # Mo ta ngan dung DU ca hai cau cua mau, khong cat con cau dau: trang Pilz
    # khong co doan gioi thieu nao nen day la tat ca thong tin bai co.
    dong = ten_dong(a)
    dv = DESCDESC.format(**dict(cv, type_c=chu_de(cv["type_c"], dong)))
    de = DESCDESC_EN.format(**dict(ce, type_l=chu_de(ce["type_l"], dong)))

    lv, le = spec_lines(a, en)
    if le:
        row("Mô tả dài của sản phẩm",
            "\n".join("- " + x for x in lv),
            "\n".join("- " + x for x in le))
    else:
        # Muc khong co bo loc nao (muc chi vai ma) -> mo ta dai cung dung mau.
        row("Mô tả dài của sản phẩm", dv, de)
    row("Mô tả ngắn của sản phẩm", dv, de)
    row("Mô tả thông số", SPECDESC.format(**cv), SPECDESC_EN.format(**ce))

    items = a.get("items") or []
    if items:
        block(["Bảng mã sản phẩm + Thông tin sản phẩm", "Model",
               "Mã đặt hàng (Part No.)", "Mô tả", "Description (EN)"])
        for it in items[:MAX_ROWS]:
            ws.append(["", it.get("name", ""), it.get("art", ""),
                       it.get("desc_vi") or it.get("desc", ""),
                       it.get("desc", "")])
            for c in range(2, 6):
                ws.cell(row=ws.max_row, column=c).alignment = WRAP
        total = a.get("total") or len(items)
        if total > len(items):
            ws.append(["", "(trang hãng còn %d mã nữa, bài lấy %d mã đầu)"
                       % (total - len(items), len(items))])
        elif len(items) > MAX_ROWS:
            ws.append(["", "(bài còn %d mã nữa, bảng in %d mã đầu)"
                       % (len(items) - MAX_ROWS, MAX_ROWS)])
    else:
        block(["Bảng mã sản phẩm + Thông tin sản phẩm",
               "(trang hãng không có mã SP cụ thể)"])

    ws.append([])
    ws.append(["Nhánh danh mục", " > ".join(a.get("path") or []), ""])
    ws.append(["Nguồn", a.get("url_src", ""), ""])
    for i, w in enumerate([34, 40, 22, 50, 50], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    fname = "%03d - %s.xlsx" % (idx, safe(en))
    wb.save(os.path.join(out_dir, fname))
    return fname


def build_index(arts, names, out_dir):
    wb = Workbook()
    ws = wb.active
    ws.title = "MUC LUC"
    ws["A1"] = "MỤC LỤC BÀI ĐĂNG - 09 Pilz"
    ws["A1"].font = Font(bold=True, size=13, color="1F4E79")
    ws.append([])
    ws.append(["STT", "Mục sản phẩm", "Tên tiếng Việt", "File", "Ảnh",
               "Số mã SP", "Tổng mã trên web", "Số thông số",
               "Danh mục Ntrend", "Nhánh danh mục", "Nguồn"])
    for c in range(1, 12):
        ws.cell(row=3, column=c).fill = FLD_FILL
        ws.cell(row=3, column=c).font = Font(bold=True)
    for i, (a, f) in enumerate(zip(arts, names), 1):
        ws.append([i, a.get("title_en") or a["name"], a.get("name_vi", ""), f,
                   "x" if a.get("img_file") else "",
                   len(a.get("items") or []), a.get("total") or 0,
                   len(a.get("specs_en") or []), a.get("cat_vi", ""),
                   " > ".join(a.get("path") or []), a.get("url_src", "")])
    for i, w in enumerate([6, 40, 40, 48, 6, 10, 14, 12, 50, 60, 60], 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A4"
    wb.save(os.path.join(out_dir, "00 MUC LUC.xlsx"))


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    logo_fallback.run("pilz")
    arts = json.load(io.open(SRC, encoding="utf-8"))["articles"]
    os.makedirs(OUT_DIR, exist_ok=True)
    names = [build_one(a, i, OUT_DIR) for i, a in enumerate(arts, 1)]
    build_index(arts, names, OUT_DIR)
    prune(OUT_DIR, names)
    nit = sum(len(a.get("items") or []) for a in arts)
    print("Saved %d bai (%d ma SP) -> %s" % (len(arts), nit, OUT_DIR))
    for lbl, bad in (("khong co ma SP", [a for a in arts if not a.get("items")]),
                     ("khong co thong so",
                      [a for a in arts if not a.get("specs_en")]),
                     ("khong co anh", [a for a in arts if not a.get("img_file")])):
        if bad:
            print("  %s (%d): %s" % (lbl, len(bad),
                  ", ".join((x.get("title_en") or x["name"])[:26] for x in bad[:8])))


if __name__ == "__main__":
    main()
