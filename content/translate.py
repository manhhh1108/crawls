# -*- coding: utf-8 -*-
"""EN -> VI qua endpoint free translate.googleapis.com (cach lam giong pipeline
SICK cu), co cache tren dia o .cache/translate_vi.json."""
import json, os, io, time, hashlib, subprocess, urllib.parse, re
from concurrent.futures import ThreadPoolExecutor

WORKERS = 6

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


def _call_blob(items):
    """Dich nhieu chuoi NGAN trong 1 request bang cach noi bang xuong dong.
    Tra ve list cung do dai, hoac None neu Google tra ve lech dong."""
    blob = "\n".join(items)
    out = _call(blob)
    lines = out.split("\n")
    if len(lines) != len(items):
        return None
    return [l.strip() for l in lines]


_RX_WORD = re.compile(r"(?<![A-Za-z])[a-z]{3,}(?![A-Za-z])")   # chi tu thuong:
#   "Safety Designer", "Flexi Net", "PROFIBUS, CANopen" la ten rieng -> bo qua
_RX_VI = re.compile(r"[ăâđêôơưĂÂĐÊÔƠƯáàảãạắằẳẵặấầẩẫậéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộ"
                    r"ớờởỡợúùủũụứừửữựýỳỷỹỵ]")


def _looks_untranslated(src, out):
    """Google thi thoang tra ve nguyen van 1 dong khi dich theo lo -> dich lai."""
    if not out or out.strip() != src.strip():
        return False
    return len(_RX_WORD.findall(src)) >= 2 and not _RX_VI.search(src)


def vi_many(texts, sleep=0.1, batch=60, maxchars=4000):
    """Dich 1 loat chuoi (co cache). Nhanh hon vi() ~20 lan nho gop request."""
    global _dirty
    texts = [(t or "").strip() for t in texts]
    keys = [hashlib.sha1(t.encode("utf-8")).hexdigest() if t else "" for t in texts]
    todo = []
    for t, k in zip(texts, keys):
        if t and k not in _CACHE and "\n" not in t and len(t) < 900 and t not in todo:
            todo.append(t)

    chunks, i = [], 0
    while i < len(todo):
        chunk, n = [], 0
        while i < len(todo) and len(chunk) < batch and n + len(todo[i]) < maxchars:
            chunk.append(todo[i])
            n += len(todo[i]) + 1
            i += 1
        chunks.append(chunk)

    def work(chunk):
        got = None
        for attempt in range(2):
            try:
                got = _call_blob(chunk)
                break
            except Exception:
                time.sleep(1.5 * (attempt + 1))
        if got is None:                      # lech dong -> dich tung cai
            return [(s, None) for s in chunk]
        return list(zip(chunk, got))

    if chunks:
        retry = []
        with ThreadPoolExecutor(max_workers=WORKERS) as ex:
            for res in ex.map(work, chunks):
                for s, r in res:
                    if r is None or _looks_untranslated(s, r):
                        retry.append(s)
                        continue
                    _CACHE[hashlib.sha1(s.encode("utf-8")).hexdigest()] = r or s
                    _dirty += 1
        for s in retry:                      # dich lai tung cai cho chac
            vi(s)
        _flush()

    out = []
    for t, k in zip(texts, keys):
        if not t:
            out.append("")
            continue
        r = _CACHE.get(k)
        if r is None or _looks_untranslated(t, r):   # tu sua cache hong tu lan truoc
            _CACHE.pop(k, None)
            r = vi(t)
        out.append(r)
    return out


def vi(text, sleep=0.12):
    """Translate one EN string to Vietnamese (cached)."""
    global _dirty
    text = (text or "").strip()
    if not text:
        return ""
    key = hashlib.sha1(text.encode("utf-8")).hexdigest()
    if key in _CACHE and _looks_untranslated(text, _CACHE[key]):
        del _CACHE[key]          # cache hong tu lan dich theo lo -> bo, dich lai
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
