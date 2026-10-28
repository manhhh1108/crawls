# -*- coding: utf-8 -*-
"""Dich truoc theo lo lon cho cac crawler moi (wenglor / Leuze / IFM / Mitsubishi).

Dich tung bai thi moi lo chi co vai chuc chuoi -> moi request API cho duoc rat
it, dot sach hang muc request/ngay (free tier Gemini = 500 req/ngay/model) roi
phai nam cho het cooldown. Gom chuoi cua TAT CA cac bai lai dich mot luot thi
mot request cho duoc hang tram chuoi, nhanh hon hang chuc lan. Dich xong nam
trong cache tren dia, vong lap chinh cua crawler chi viec lay ra.

Dung trong crawler:

    import prewarm
    prewarm.run(bag)        # bag = {"key": set(), "value": set(), "prose": set()}
"""
import sys

CHUNK = 400


def run(bag, chunk=CHUNK):
    """bag: {"key": iterable, "value": iterable, "prose": iterable}."""
    from vi_engine import translate_many
    from translate import _flush

    bag = {k: sorted(x for x in (v or ()) if x and str(x).strip())
           for k, v in bag.items()}
    tot = sum(len(v) for v in bag.values())
    if not tot:
        return
    print("prewarm: gom %d chuoi (key=%d value=%d prose=%d), dich theo lo %d"
          % (tot, len(bag.get("key") or []), len(bag.get("value") or []),
             len(bag.get("prose") or []), chunk))
    done = 0
    for kind in ("key", "value", "prose"):
        items = bag.get(kind) or []
        for i in range(0, len(items), chunk):
            translate_many(items[i:i + chunk], kind)
            done += len(items[i:i + chunk])
            _flush(True)
            print("   prewarm %d/%d (%s)" % (done, tot, kind))
            sys.stdout.flush()
    _flush(True)
