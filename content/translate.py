# -*- coding: utf-8 -*-
"""EN -> VI qua endpoint free translate.googleapis.com (cach lam giong pipeline
SICK cu), co cache tren dia o .cache/translate_vi.json."""
import json, os, io, time, hashlib, subprocess, urllib.parse, re

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
CACHE_PATH = os.path.join(REPO, ".cache", "translate_vi.json")
os.makedirs(os.path.dirname(CACHE_PATH), exist_ok=True)
try:
    with io.open(CACHE_PATH, encoding="utf-8") as f:
        _CACHE = json.load(f)
except Exception:
    _CACHE = {}
_dirty = 0


def _flush(force=False):
    global _dirty
    if _dirty == 0 and not force:
        return
    with io.open(CACHE_PATH, "w", encoding="utf-8") as f:
        json.dump(_CACHE, f, ensure_ascii=False)
    _dirty = 0


def _call(text):
    url = ("https://translate.googleapis.com/translate_a/single"
           "?client=gtx&sl=en&tl=vi&dt=t&q=" + urllib.parse.quote(text))
    p = subprocess.run(["curl", "-sS", "--max-time", "30", url],
                       capture_output=True, timeout=45)
    data = json.loads(p.stdout.decode("utf-8", "replace"))
    return "".join(seg[0] for seg in data[0] if seg and seg[0])


def vi(text, sleep=0.12):
    """Translate one EN string to Vietnamese (cached)."""
    global _dirty
    text = (text or "").strip()
    if not text:
        return ""
    key = hashlib.sha1(text.encode("utf-8")).hexdigest()
    if key in _CACHE:
        return _CACHE[key]
    parts, cur = [], ""
    for s in re.split(r"(?<=[.!?])\s+", text):
        if len(cur) + len(s) > 1500 and cur:
            parts.append(cur)
            cur = s
        else:
            cur = (cur + " " + s).strip()
    if cur:
        parts.append(cur)
    out = []
    for pt in parts:
        for attempt in range(3):
            try:
                out.append(_call(pt))
                break
            except Exception:
                time.sleep(1.5 * (attempt + 1))
        else:
            out.append(pt)  # bo tay: giu nguyen EN
        time.sleep(sleep)
    res = " ".join(out).strip()
    _CACHE[key] = res
    _dirty += 1
    if _dirty >= 20:
        _flush()
    return res
