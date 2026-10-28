# -*- coding: utf-8 -*-
"""Ghep anh san pham theo nhan dien NTREND (ti le can theo file mau Canva:
logo trai ~45% be rong, hotline phai voi icon tron + ong nghe, nen trang)."""
import os
import math
from PIL import Image, ImageChops, ImageDraw, ImageFont

HERE = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(HERE, "assets", "ntrend-logo.png")
PHONE = "(+84) 24 3232 1532"
_logo_cache = None

# ti le can theo anh mau Canva (829px): logo trai 10%/10%, rong ~19.5%;
# hotline phai: font ~3.1%, icon tron ~5.5%, cach le phai ~8.5%
L_MARGIN = 0.103
T_MARGIN = 0.102
LOGO_W = 0.195
R_MARGIN = 0.085
FONT_FRAC = 0.031
ICON_FRAC = 0.055


def _logo():
    global _logo_cache
    if _logo_cache is None:
        im = Image.open(LOGO_PATH).convert("RGB")
        bg = Image.new("RGB", im.size, (255, 255, 255))
        bbox = ImageChops.difference(im, bg).getbbox()
        _logo_cache = im.crop(bbox)
    return _logo_cache


def _font(size, bold=True):
    name = "arialbd.ttf" if bold else "arial.ttf"
    return ImageFont.truetype("C:/Windows/Fonts/" + name, size)


PHONE_ICON = os.path.join(HERE, "assets", "phone-icon.png")
_icon_cache = {}


def _phone_icon(dia):
    """Icon dien thoai cat tu anh mau cua team, scale theo duong kinh can ve."""
    if dia not in _icon_cache:
        ic = Image.open(PHONE_ICON).convert("L")
        ic = ic.resize((dia, dia), Image.LANCZOS)
        _icon_cache[dia] = ic
    return _icon_cache[dia]


def _handset(d, cx, cy, r, color, lw):
    dia = 2 * r
    ic = _phone_icon(dia)
    # dung anh xam lam mask: pixel toi = net icon
    from PIL import ImageOps
    mask = ImageOps.invert(ic)
    tint = Image.new("RGB", ic.size, color)
    d._image.paste(tint, (int(cx - r), int(cy - r)), mask)


MAX_UPSCALE = 1.4       # phong to anh SP qua muc nay la nhin thay ro vo


def _fit_size(src_path, size):
    """Chon canh khung sao cho anh SP khong bi phong to qua MAX_UPSCALE.

    Anh goc cua vai hang chi 200-300px (the thu nho tren trang danh sach). Ep
    len khung 800 la phong 3-4 lan, nhin vo het. Thu nho khung lai thi ty le bo
    cuc khong doi (moi thu deu tinh theo `size`) ma anh van net.
    """
    try:
        from PIL import Image
        w, h = Image.open(src_path).size
    except Exception:
        return size
    for n in range(size, 399, -20):
        top = int(n * T_MARGIN) + max(1, int(_logo().height * int(n * LOGO_W)
                                             / _logo().width)) + int(n * 0.035)
        m = int(n * L_MARGIN)
        box_w, box_h = n - 2 * m, n - top - m
        if min(box_w / w, box_h / h) <= MAX_UPSCALE:
            return n
    return 400


def compose(src_path, dst_path, size=800):
    size = _fit_size(src_path, size)
    canvas = Image.new("RGB", (size, size), "white")
    d = ImageDraw.Draw(canvas)
    color = (35, 35, 35)

    # logo trai
    logo = _logo()
    lw_px = int(size * LOGO_W)
    lh_px = max(1, int(logo.height * lw_px / logo.width))
    lx, ly = int(size * L_MARGIN), int(size * T_MARGIN)
    canvas.paste(logo.resize((lw_px, lh_px), Image.LANCZOS), (lx, ly))
    cy = ly + lh_px // 2  # truc giua header

    # hotline phai: text truoc (can phai), icon ben trai text
    f = _font(int(size * FONT_FRAC))
    tb = d.textbbox((0, 0), PHONE, font=f)
    tw, th = tb[2] - tb[0], tb[3] - tb[1]
    tx = size - int(size * R_MARGIN) - tw
    d.text((tx, cy - th // 2 - tb[1]), PHONE, fill=color, font=f)
    r = int(size * ICON_FRAC) // 2
    icx = tx - r - int(size * 0.022)
    _handset(d, icx, cy, r, color, max(2, size // 260))

    # anh san pham: can giua phan con lai
    top = ly + lh_px + int(size * 0.035)
    margin = int(size * L_MARGIN)
    box_w = size - 2 * margin
    box_h = size - top - margin
    prod = Image.open(src_path)
    if prod.mode in ("RGBA", "LA", "P"):
        prod = prod.convert("RGBA")
        white = Image.new("RGBA", prod.size, (255, 255, 255, 255))
        prod = Image.alpha_composite(white, prod).convert("RGB")
    else:
        prod = prod.convert("RGB")
    scale = min(box_w / prod.width, box_h / prod.height)
    nw, nh = max(1, int(prod.width * scale)), max(1, int(prod.height * scale))
    prod = prod.resize((nw, nh), Image.LANCZOS)
    canvas.paste(prod, (margin + (box_w - nw) // 2, top + (box_h - nh) // 2))

    os.makedirs(os.path.dirname(dst_path), exist_ok=True)
    ext = os.path.splitext(dst_path)[1].lower()
    if ext in (".jpg", ".jpeg"):
        canvas.save(dst_path, "JPEG", quality=90, optimize=True, subsampling=0)
    else:
        canvas.save(dst_path, "PNG")
    return dst_path
