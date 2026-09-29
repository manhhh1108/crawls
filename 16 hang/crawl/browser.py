# -*- coding: utf-8 -*-
"""Headed-Chromium helper: some vendor sites (Akamai/Cloudflare bot mitigation)
reject curl and headless browsers, so we drive a real browser and issue the
JSON/API calls from inside the page (same TLS + cookie context as the site)."""
import json, os, time, gzip, hashlib

HERE = os.path.dirname(os.path.abspath(__file__))
BCACHE = os.path.join(HERE, "bcache")
os.makedirs(BCACHE, exist_ok=True)

FETCH_JS = """async (args) => {
  const [u, opts] = args;
  const r = await fetch(u, Object.assign({credentials:'include'}, opts||{}));
  return [r.status, await r.text()];
}"""


class Browser(object):
    def __init__(self, home, headless=False, locale="en-US"):
        from playwright.sync_api import sync_playwright
        self._pw = sync_playwright().start()
        self.b = self._pw.chromium.launch(headless=headless)
        self.ctx = self.b.new_context(locale=locale, viewport={"width": 1400, "height": 900})
        self.pg = self.ctx.new_page()
        self.home = home
        self.pg.goto(home, wait_until="domcontentloaded", timeout=90000)
        self.pg.wait_for_timeout(3500)

    def goto(self, url, wait="domcontentloaded", pause=2500, timeout=90000):
        self.pg.goto(url, wait_until=wait, timeout=timeout)
        self.pg.wait_for_timeout(pause)
        return self.pg.content()

    def _key(self, url, opts):
        h = hashlib.sha1((url + "|" + json.dumps(opts, sort_keys=True)).encode()).hexdigest()
        return os.path.join(BCACHE, h + ".gz")

    def api(self, url, opts=None, cache=True, retries=3, sleep=0.15):
        path = self._key(url, opts or {})
        if cache and os.path.exists(path):
            with gzip.open(path, "rt", encoding="utf-8", errors="replace") as f:
                return f.read()
        last = None
        for i in range(retries):
            try:
                st, txt = self.pg.evaluate(FETCH_JS, [url, opts or {}])
                if st == 200:
                    if cache:
                        with gzip.open(path, "wt", encoding="utf-8") as f:
                            f.write(txt)
                    time.sleep(sleep)
                    return txt
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
