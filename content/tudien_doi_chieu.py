# -*- coding: utf-8 -*-
"""Doi chieu TU DIEN voi du lieu da crawl: cau tieng Anh co tu khoa ma ban
dich tieng Viet CHUA co ban dich mong muon -> in ra cac kieu dang bi dung.

    python tudien_doi_chieu.py

Dung truoc khi viet luat moi cho apply_tudien.py. Mot tu thuong bi dich ra
nhieu kieu khac nhau (fork -> "chac" / "dang chu U" / "dang fork" / "dang
khe"), sua mot kieu la chua xong.

CANH BAO: ket qua co RAT NHIEU bao dong gia - "NO" khop "No. of pins",
"flush" khop "non-flush" (ma "lap noi" la dung). Phai doc mau roi moi viet
luat, dung sua hang loat theo con so.
"""
import collections
import glob
import io
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
from apply_tudien import duyet                              # noqa: E402

from openpyxl import load_workbook                          # noqa: E402

TD = os.path.join(os.path.dirname(HERE), "data", "Tu dien dich EN-VI.xlsx")
ws = load_workbook(TD, read_only=True, data_only=True).active
CAP = []
for r in ws.iter_rows(values_only=True):
    if not r or len(r) < 2:
        continue
    en = str(r[0] or "").strip()
    vi = str(r[1] or "").strip()
    if en and vi and en != "Ti\u1ebfng Anh" and len(en) < 60:
        CAP.append((en, vi))
print("tu dien: %d cap\n" % len(CAP))

RX = [(en, vi,
       re.compile(r"(?<![A-Za-z])" + re.escape(en) + r"(?![A-Za-z])", re.I),
       re.compile(re.escape(vi), re.I))
      for en, vi in CAP]

thieu = collections.Counter()
mau = collections.defaultdict(collections.Counter)
for p in sorted(glob.glob(os.path.join(HERE, "out", "*", "content.json"))):
    h = os.path.basename(os.path.dirname(p))
    for a in json.load(io.open(p, encoding="utf-8"))["articles"]:
        for en, chu, kh, _tr in duyet(a):
            v = chu[kh]
            if not en or not isinstance(v, str) or not v.strip():
                continue
            for ten, vi, rxe, rxv in RX:
                if rxe.search(en) and not rxv.search(v):
                    thieu[(ten, h)] += 1
                    mau[ten][(en[:85], v[:85])] += 1

print("### Cap tu dien CON SOT (EN co tu, VI chua co ban dich)")
tong = collections.Counter()
for (ten, _h), n in thieu.items():
    tong[ten] += n
for ten, n in tong.most_common():
    vi = dict(CAP)[ten]
    hs = {h: c for (t, h), c in thieu.items() if t == ten}
    print("\n=== %-34s -> %-34s %d chuoi" % (ten, vi, n))
    print("    hang:", hs)
    for (e, v), c in mau[ten].most_common(6):
        print("      x%-3d EN: %s" % (c, e))
        print("           VI: %s" % v)
