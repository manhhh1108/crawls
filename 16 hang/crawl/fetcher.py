# -*- coding: utf-8 -*-
"""Shared HTTP fetch helper (curl-backed, browser-like headers) + on-disk cache.

Uses curl because several vendor sites (Festo/Akamai, Schneider, ...) fingerprint
the TLS handshake and reject python-requests outright.
"""
import os, time, json, hashlib, gzip, random, subprocess, tempfile

HERE = os.path.dirname(os.path.abspath(__file__))
CACHE_DIR = os.path.join(HERE, "cache")
os.makedirs(CACHE_DIR, exist_ok=True)

UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

HTML_HEADERS = [
    ("Accept", "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8"),
    ("Accept-Language", "en-US,en;q=0.9"),
    ("sec-ch-ua", '"Chromium";v="131", "Not_A Brand";v="24"'),
    ("sec-ch-ua-platform", '"Windows"'),
    ("sec-ch-ua-mobile", "?0"),
    ("Sec-Fetch-Site", "none"),
    ("Sec-Fetch-Mode", "navigate"),
    ("Sec-Fetch-User", "?1"),
    ("Sec-Fetch-Dest", "document"),
    ("Upgrade-Insecure-Requests", "1"),
]

JSON_HEADERS = [
    ("Accept", "application/json, text/plain, */*"),
    ("Accept-Language", "en-US,en;q=0.9"),
    ("sec-ch-ua", '"Chromium";v="131", "Not_A Brand";v="24"'),
    ("sec-ch-ua-platform", '"Windows"'),
    ("sec-ch-ua-mobile", "?0"),
    ("Sec-Fetch-Site", "same-origin"),
    ("Sec-Fetch-Mode", "cors"),
    ("Sec-Fetch-Dest", "empty"),
]


def _key(url, extra):
    h = hashlib.sha1((url + "|" + json.dumps(extra, sort_keys=True, default=str)).encode("utf-8")).hexdigest()
    return os.path.join(CACHE_DIR, h + ".gz")


def fetch(url, headers=None, json_mode=False, timeout=60, retries=3, cache=True,
          sleep=0.3, referer=None, method="GET", data=None, cookie_jar=None):
    extra = {"m": method, "d": data, "h": headers, "j": json_mode, "r": referer}
    path = _key(url, extra)
    if cache and os.path.exists(path):
        with gzip.open(path, "rt", encoding="utf-8", errors="replace") as f:
            return f.read()

    hdrs = list(JSON_HEADERS if json_mode else HTML_HEADERS)
    if referer:
        hdrs = [(k, v) for k, v in hdrs if k != "Sec-Fetch-Site"]
        hdrs += [("Sec-Fetch-Site", "same-origin"), ("Referer", referer)]
    if headers:
        keys = {k.lower() for k in headers}
        hdrs = [(k, v) for k, v in hdrs if k.lower() not in keys] + list(headers.items())

    last = None
    for i in range(retries):
        cmd = ["curl", "-sS", "-L", "--compressed", "--max-time", str(timeout),
               "-A", UA, "-w", "\n__HTTP_CODE__%{http_code}"]
        for k, v in hdrs:
            cmd += ["-H", "%s: %s" % (k, v)]
        if cookie_jar:
            cmd += ["-b", cookie_jar, "-c", cookie_jar]
        if method != "GET":
            cmd += ["-X", method]
        if data is not None:
            cmd += ["--data-binary", data]
        cmd.append(url)
        try:
            p = subprocess.run(cmd, capture_output=True, timeout=timeout + 20)
            out = p.stdout.decode("utf-8", "replace")
            marker = out.rfind("\n__HTTP_CODE__")
            code = out[marker + len("\n__HTTP_CODE__"):].strip() if marker >= 0 else "000"
            body = out[:marker] if marker >= 0 else out
            if code == "200":
                if cache:
                    with gzip.open(path, "wt", encoding="utf-8") as f:
                        f.write(body)
                time.sleep(sleep + random.random() * 0.2)
                return body
            last = "HTTP %s" % code
            if code in ("404", "410"):
                break
        except Exception as e:  # noqa
            last = repr(e)
        time.sleep(1.2 * (i + 1))
    raise RuntimeError("FAIL %s -> %s" % (url, last))


def fetch_json(url, **kw):
    kw["json_mode"] = True
    return json.loads(fetch(url, **kw))


def try_fetch(url, **kw):
    try:
        return fetch(url, **kw)
    except Exception as e:
        print("  [warn]", e)
        return None
