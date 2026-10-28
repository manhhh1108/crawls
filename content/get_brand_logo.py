# -*- coding: utf-8 -*-
"""Lay logo mot hang tu trang chu hang do, luu vao assets/<key>-logo.png.

    python get_brand_logo.py espe https://www.espetech.com/
    python get_brand_logo.py espe https://www.espetech.com/ --inplace

Phai dung trinh duyet vi phan lon logo la SVG ma may khong co thu vien doi
SVG sang anh. Cach lam: tim phan tu logo tren trang chu, roi DUNG LAI no tren
mot trang trang tron o co lon va chup - lam vay thi khong vuong banner cookie,
va SVG ra net o bat ky co nao.

    --inplace   chup thang tren trang, giu nguyen nen cua hang.
                Dung khi logo la chu TRANG: dung lai tren nen trang thi mat
                chu (da gap voi ifm va OTENNLUX).

Chup xong nho MO RA XEM. Script khong biet no vo phai icon gio hang hay anh
bia; no chi biet do la phan tu dau tien hop le.
"""
import os
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(HERE, "assets")

# Tim theo thu tu: cho nao chac chan la logo truoc, chung chung sau.
FIND = """() => {
  const sels = ['header a[href="/"] img', 'a[class*="logo" i] img',
    '[class*="logo" i] img', 'img[alt*="logo" i]', 'img[src*="logo" i]',
    'header img', 'a[class*="logo" i] svg', '[class*="logo" i] svg',
    'header svg'];
  for (const s of sels) {
    for (const el of document.querySelectorAll(s)) {
      const r = el.getBoundingClientRect();
      if (r.width < 40 || r.height < 10) continue;
      if (el.tagName.toLowerCase() === 'img') {
        const u = el.currentSrc || el.src;
        if (!u || u.startsWith('data:image/gif')) continue;
        return {kind: 'img', url: new URL(u, location.href).href,
                w: r.width, h: r.height, sel: s};
      }
      return {kind: 'svg', html: el.outerHTML, w: r.width, h: r.height, sel: s};
    }
  }
  return null;
}"""


def trim(path, pad_frac=0.02, cap=1600):
    """Cat vien trang quanh logo roi ha ve co vua phai."""
    from PIL import Image
    im = Image.open(path)
    if im.mode in ("RGBA", "LA", "P"):
        im = im.convert("RGBA")
        bg = Image.new("RGBA", im.size, (255, 255, 255, 255))
        im = Image.alpha_composite(bg, im)
    im = im.convert("RGB")
    ink = im.convert("L").point(lambda v: 255 if v < 248 else 0)
    bb = ink.getbbox()
    if bb:
        pad = max(4, int(pad_frac * max(im.size)))
        im = im.crop((max(0, bb[0] - pad), max(0, bb[1] - pad),
                      min(im.size[0], bb[2] + pad),
                      min(im.size[1], bb[3] + pad)))
    if max(im.size) > cap:
        im.thumbnail((cap, cap), Image.LANCZOS)
    im.save(path)
    return im.size


def grab(key, home, inplace=False, width=1400):
    from playwright.sync_api import sync_playwright
    os.makedirs(ASSETS, exist_ok=True)
    dst = os.path.join(ASSETS, key + "-logo.png")
    pw = sync_playwright().start()
    b = pw.chromium.launch(headless=False)
    ctx = b.new_context(viewport={"width": 1500, "height": 900},
                        device_scale_factor=3, locale="en-US")
    try:
        pg = ctx.new_page()
        pg.goto(home, wait_until="domcontentloaded", timeout=90000)
        pg.wait_for_timeout(4000)
        info = pg.evaluate(FIND)
        if not info:
            print("khong tim thay phan tu logo tren %s" % home)
            return ""
        print("thay %s %dx%d qua '%s'"
              % (info["kind"], info["w"], info["h"], info["sel"]))
        if info["kind"] == "img":
            print("  ", info["url"][:120])
        if inplace:
            pg.locator(info["sel"]).first.screenshot(path=dst)
        pg.close()
        if not inplace:
            t = ctx.new_page()
            body = ('<img id="L" src="%s" style="width:%dpx;display:block">'
                    % (info["url"], width)) if info["kind"] == "img" else \
                   ('<div id="L" style="width:%dpx">%s</div>'
                    % (width, info["html"]))
            t.set_content(
                '<html><body style="margin:0;background:#fff">%s'
                '<style>#L svg{width:100%%;height:auto;display:block}</style>'
                '</body></html>' % body, wait_until="load")
            t.wait_for_timeout(1500)
            t.locator("#L").screenshot(path=dst)
            t.close()
    finally:
        b.close()
        pw.stop()
    size = trim(dst, pad_frac=0 if inplace else 0.02)
    print("-> %s  %dx%d" % (dst, size[0], size[1]))
    print("   MO RA XEM truoc khi dung: chup nham icon/anh bia la chuyen thuong.")
    return dst


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    if len(args) < 2:
        print(__doc__)
        return
    grab(args[0], args[1], inplace="--inplace" in sys.argv)


if __name__ == "__main__":
    main()
