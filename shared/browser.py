# -*- coding: utf-8 -*-
"""Headed-Chromium helper: some vendor sites (Akamai/Cloudflare bot mitigation)
reject curl and headless browsers, so we drive a real browser and issue the
JSON/API calls from inside the page (same TLS + cookie context as the site)."""
import json, os, time, gzip, hashlib

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
BCACHE = os.path.join(REPO, ".cache", "browser")
os.makedirs(BCACHE, exist_ok=True)

FETCH_JS = """async (args) => {
  const [u, opts] = args;
  const r = await fetch(u, Object.assign({credentials:'include'}, opts||{}));
  return [r.status, await r.text()];
}"""


class Browser(object):
    def __init__(self, home, headless=False, locale="en-US", tries=4):
        """Chrome co luc mo len roi tat ngay ("Target page, context or browser
        has been closed"), lan sau lai binh thuong - nen thu lai vai lan, khong
        thi ca lenh cao chet tu dong 0.

        Thu lai PHAI nam trong day, dung mot phien Playwright: goi
        sync_playwright().start() lan thu hai trong cung mot luong se bao
        "Playwright Sync API inside the asyncio loop" va khong bao gio mo duoc.
        """
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self.home = home
        self.b = self.ctx = self.pg = None
        last = None
        for k in range(tries):
            try:
                self.b = self._pw.chromium.launch(headless=headless)
                self.ctx = self.b.new_context(
                    locale=locale, viewport={"width": 1400, "height": 900})
                self.pg = self.ctx.new_page()
                self.pg.goto(home, wait_until="domcontentloaded", timeout=90000)
                self.pg.wait_for_timeout(3500)
                return
            except Exception as e:
                last = repr(e)[:70]
                print("  [mo Chrome that bai lan %d: %s]" % (k + 1, last))
                try:
                    if self.b:
                        self.b.close()
                except Exception:
                    pass
                self.b = self.ctx = self.pg = None
                time.sleep(5 * (k + 1))
        self._pw.stop()
        raise RuntimeError("khong mo duoc Chrome sau %d lan: %s" % (tries, last))

    def goto(self, url, wait="domcontentloaded", pause=2500, timeout=90000):
        self.pg.goto(url, wait_until=wait, timeout=timeout)
        self.pg.wait_for_timeout(pause)
        return self.pg.content()

    def _key(self, url, opts):
        h = hashlib.sha1((url + "|" + json.dumps(opts, sort_keys=True)).encode()).hexdigest()
        return os.path.join(BCACHE, h + ".gz")

    def api(self, url, opts=None, cache=True, retries=3, sleep=0.15, accept=None):
        """`accept(text) -> bool` kiem tra noi dung truoc khi nhan va luu cache.

        Can cai nay vi co trang tra ve HTTP 200 kem mot trang LOI (festo tra ve
        vo micro-frontend 4.5KB khong co _ROUTER_DATA khi bi chan) - khong kiem
        thi vo loi do vao cache va moi lan chay sau deu doc lai tu cache, dong
        san pham do mat noi dung vinh vien."""
        path = self._key(url, opts or {})
        if cache and os.path.exists(path):
            with gzip.open(path, "rt", encoding="utf-8", errors="replace") as f:
                t = f.read()
            if accept is None or accept(t):
                return t
            os.remove(path)          # cache cu la trang loi -> bo, tai lai
        last = None
        for i in range(retries):
            try:
                st, txt = self.pg.evaluate(FETCH_JS, [url, opts or {}])
                if st == 200 and accept is not None and not accept(txt):
                    # 200 nhung la trang loi: lam moi phien o tab chinh roi thu lai
                    last = "noi dung khong hop le (%d bytes)" % len(txt)
                    try:
                        self.pg.goto(self.home, wait_until="domcontentloaded",
                                     timeout=60000)
                        self.pg.wait_for_timeout(2000)
                    except Exception:
                        pass
                elif st == 200:
                    if cache:
                        with gzip.open(path, "wt", encoding="utf-8") as f:
                            f.write(txt)
                    time.sleep(sleep)
                    return txt
                else:
                    last = "HTTP %s" % st
                    if st in (404, 410):
                        break
            except Exception as e:
                last = repr(e)
                try:
                    self.pg.goto(self.home, wait_until="domcontentloaded", timeout=60000)
                except Exception:
                    pass
            time.sleep(1.0 * (i + 1))
        raise RuntimeError("API FAIL %s -> %s" % (url, last))

    def has_cache(self, url, opts=None):
        """Trang nay da nam trong cache dia chua.

        De crawler dem duoc con bao nhieu trang phai tai that, tu do cat lo
        ma chay - doc lai tu cache thi gan nhu khong ton gi.
        """
        return os.path.exists(self._key(url, opts or {}))

    def api_json(self, url, **kw):
        return json.loads(self.api(url, **kw))

    def close(self):
        try:
            self.b.close()
        finally:
            self._pw.stop()


class Renderer(object):
    """Headless renderer with on-disk cache, for sites that build the DOM in JS
    but do not block headless Chromium."""

    def __init__(self, headless=True, locale="en-US", block_media=True):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self.b = self._pw.chromium.launch(headless=headless)
        self.ctx = self.b.new_context(locale=locale, viewport={"width": 1400, "height": 1000})
        if block_media:
            self.ctx.route("**/*", self._route)
        self.pg = self.ctx.new_page()

    @staticmethod
    def _route(route):
        if route.request.resource_type in ("image", "media", "font"):
            return route.abort()
        return route.continue_()

    def html(self, url, cache=True, wait="load", pause=2500, timeout=90000, retries=2):
        path = os.path.join(BCACHE, hashlib.sha1(("R|" + url).encode()).hexdigest() + ".gz")
        if cache and os.path.exists(path):
            with gzip.open(path, "rt", encoding="utf-8", errors="replace") as f:
                return f.read()
        last = None
        for i in range(retries):
            try:
                self.pg.goto(url, wait_until=wait, timeout=timeout)
                self.pg.wait_for_timeout(pause)
                t = self.pg.content()
                if cache:
                    with gzip.open(path, "wt", encoding="utf-8") as f:
                        f.write(t)
                return t
            except Exception as e:
                last = repr(e)
                time.sleep(1.5 * (i + 1))
        print("  [render fail]", url, last)
        return None

    def close(self):
        try:
            self.b.close()
        finally:
            self._pw.stop()
