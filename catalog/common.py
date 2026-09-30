# -*- coding: utf-8 -*-
import json, os, io

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")
os.makedirs(OUT, exist_ok=True)


def save(brand, site, nodes, note=""):
    """nodes: list of dicts {path:[...], type:'category'|'family', url, code}"""
    seen, clean = set(), []
    for n in nodes:
        k = tuple(n["path"])
        if k in seen:
            continue
        seen.add(k)
        n = dict(n)
        n["level"] = len(n["path"])
        clean.append(n)
    clean.sort(key=lambda n: [str(x).lower() for x in n["path"]])
    p = os.path.join(OUT, brand + ".json")
    with io.open(p, "w", encoding="utf-8") as f:
        json.dump({"brand": brand, "site": site, "note": note, "nodes": clean},
                  f, ensure_ascii=False, indent=1)
    lv = {}
    for n in clean:
        lv[n["level"]] = lv.get(n["level"], 0) + 1
    print("[%s] saved %d nodes -> %s" % (brand, len(clean), p))
    print("        levels:", dict(sorted(lv.items())))
    return p


def load(brand):
    with io.open(os.path.join(OUT, brand + ".json"), encoding="utf-8") as f:
        return json.load(f)


import re as _re, json as _json, html as _html

_LD = _re.compile(r'<script[^>]+type="application/ld\+json"[^>]*>(.*?)</script>', _re.S)


def ld_breadcrumb(htmltext, drop=("homepage", "home", "products", "product")):
    """Return the breadcrumb name list from a JSON-LD BreadcrumbList, if any."""
    for m in _LD.finditer(htmltext or ""):
        raw = _html.unescape(m.group(1)).strip()
        try:
            d = _json.loads(raw)
        except Exception:
            continue
        pool = d if isinstance(d, list) else [d]
        if isinstance(d, dict) and isinstance(d.get("@graph"), list):
            pool = pool + d["@graph"]
        for obj in pool:
            if not isinstance(obj, dict):
                continue
            if obj.get("@type") == "BreadcrumbList":
                items = sorted(obj.get("itemListElement") or [],
                               key=lambda x: x.get("position", 0))
                names = []
                for i in items:
                    nm = i.get("name")
                    if not nm and isinstance(i.get("item"), dict):
                        nm = i["item"].get("name")
                    names.append(_html.unescape(str(nm or "")).strip())
                while names and names[0].lower() in drop:
                    names = names[1:]
                return [n for n in names if n]
    return None


_BCOL = _re.compile(r'<ol[^>]*class="[^"]*breadcrumb[^"]*"[^>]*>(.*?)</ol>', _re.S)
_NAME = _re.compile(r'itemprop="name"[^>]*>(.*?)<', _re.S)


def micro_breadcrumb(htmltext, drop=("home", "homepage", "products", "product")):
    """Breadcrumb names from schema.org microdata markup."""
    m = _BCOL.search(htmltext or "")
    if not m:
        return None
    names = [_html.unescape(_re.sub(r"\s+", " ", x)).strip() for x in _NAME.findall(m.group(1))]
    while names and names[0].lower() in drop:
        names = names[1:]
    return [n for n in names if n] or None
