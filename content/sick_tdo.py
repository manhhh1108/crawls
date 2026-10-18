# -*- coding: utf-8 -*-
"""SICK - bang TECHNICAL DATA OVERVIEW lay tu file PDF 'Product overview'.

Theo review 13/10: thong so ky thuat cua 1 dong SP phai lay o muc
"TECHNICAL DATA OVERVIEW" trong tai lieu Product overview cua hang, KHONG phai
cac facet loc o tab Variants (Certificate / New products / Product segment /
Task ... - deu la metadata catalog, dang web vo nghia).

Bang nay chi co trong PDF: truong `TechData` o API cap dong SP luon rong.
Link PDF nam trong `DownloadJson` cua ban ghi detail (muc productoverview_<g>_en).

Chay doc lap:
    python sick_tdo.py            # bo sung tdo.json cho moi dong con thieu
    python sick_tdo.py --force    # lam lai tu dau

Ket qua: out/01 SICK/tdo.json = {"<gcode>": [[ten_thong_so, gia_tri], ...]}
PDF chi tai ve thu muc tam roi xoa ngay sau khi boc (500+ file, ~1 GB).
"""
import io
import json
import os
import re
import subprocess
import sys
import tempfile
import threading

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
OUT_DIR = os.path.join(HERE, "out", "01 SICK")
TDO_JSON = os.path.join(OUT_DIR, "tdo.json")
SITE = "https://www.sick.com"
HEADING = "TECHNICAL DATA OVERVIEW"
UA = ("Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 "
      "(KHTML, like Gecko) Chrome/131.0.0.0 Safari/537.36")

WRAP_TOL = 16.0        # dong ket thuc cach le phai duoi nguong nay = bi ngat dong

_lock = threading.Lock()
# PyMuPDF dung 1 context MuPDF chung -> mo/doc PDF tu nhieu luong cung luc cho
# ket qua rong bua bai (lan chay 18/10: 284/510 file bi rong oan). Tai thi chay
# song song duoc, nhung BOC phai xep hang.
_pdf_lock = threading.Lock()


def pdf_url(path):
    """Path trong DownloadJson -> URL tai that."""
    p = (path or "").strip()
    if not p:
        return ""
    if p.startswith("http"):
        return p
    return SITE + "/media" + (p if p.startswith("/") else "/" + p)


def _norm(s):
    return re.sub(r"[ \t\xa0]+", " ", str(s or "").replace("\xad", "")).strip()


def _cell_text(page, bbox):
    """Doc 1 o cua bang, phan biet XUONG DONG DO NGAT CHU va GIA TRI KHAC NHAU.

    PDF cua SICK khong dung ky tu ngan cach giua cac gia tri: moi gia tri la 1
    dong rieng. Dong nao chay sat le phai cua o -> la cau bi ngat (noi lai bang
    dau cach); dong ngan hon han -> la gia tri rieng (noi bang '; ').
    """
    import fitz

    right = bbox[2]
    lines = []
    for blk in page.get_text("dict", clip=fitz.Rect(bbox))["blocks"]:
        for ln in blk.get("lines", []):
            t = _norm("".join(sp["text"] for sp in ln["spans"]))
            if t:
                lines.append((ln["bbox"][1], ln["bbox"][2], t))
    lines.sort(key=lambda x: x[0])
    if not lines:
        return ""

    out = lines[0][2]
    for i in range(1, len(lines)):
        prev_right = lines[i - 1][1]
        t = lines[i][2]
        if out.endswith("-") and t[:1].islower():
            out = out[:-1] + t                      # tu bi ngat giua dong
        elif prev_right >= right - WRAP_TOL:
            out += " " + t                          # cau dai bi xuong dong
        else:
            out += "; " + t                         # gia tri khac
    return _norm(out)


def parse_pdf(path):
    """Boc bang ngay duoi tieu de TECHNICAL DATA OVERVIEW -> [[ten, gia tri], ...]."""
    import fitz

    rows = []
    with fitz.open(path) as doc:
        for page in doc:
            hits = page.search_for(HEADING)
            if not hits:
                continue
            top = min(r.y1 for r in hits)
            try:
                tables = page.find_tables().tables
            except Exception:
                tables = []
            # bang dung ngay duoi tieu de (gan nhat ve phia duoi)
            below = sorted((t for t in tables if t.bbox[1] >= top - 2),
                           key=lambda t: t.bbox[1])
            if not below:
                continue
            for trow in below[0].rows:
                boxes = [b for b in trow.cells if b is not None]
                if len(boxes) < 2:
                    continue
                k = _cell_text(page, boxes[0])
                v = "; ".join(x for x in (_cell_text(page, b) for b in boxes[1:]) if x)
                if k and v and k.upper() != HEADING:
                    rows.append([k, v])
            if rows:
                break
    return rows


def fetch_one(url, timeout=120):
    """Tai PDF ve file tam, boc bang, xoa file. Tra [[ten, gia tri], ...]."""
    fd, tmp = tempfile.mkstemp(suffix=".pdf")
    os.close(fd)
    try:
        p = subprocess.run(["curl", "-sS", "-L", "--max-time", str(timeout),
                            "-A", UA, "-o", tmp, url],
                           capture_output=True, timeout=timeout + 30)
        if p.returncode != 0 or os.path.getsize(tmp) < 2000:
            return []
        with _pdf_lock:
            return parse_pdf(tmp)
    except Exception:
        return []
    finally:
        try:
            os.remove(tmp)
        except OSError:
            pass


def load():
    if os.path.exists(TDO_JSON):
        with io.open(TDO_JSON, encoding="utf-8") as f:
            return json.load(f)
    return {}


def save(d):
    os.makedirs(OUT_DIR, exist_ok=True)
    with io.open(TDO_JSON, "w", encoding="utf-8") as f:
        json.dump(d, f, ensure_ascii=False)


def main():
    sys.stdout.reconfigure(encoding="utf-8")
    force = "--force" in sys.argv
    with io.open(os.path.join(OUT_DIR, "raw.json"), encoding="utf-8") as f:
        lines = json.load(f)["lines"]

    done = {} if force else load()
    todo = [l for l in lines if l.get("pdf_overview") and l["code"] not in done]
    nopdf = [l for l in lines if not l.get("pdf_overview")]
    print("co link Product overview: %d | da boc: %d | con lai: %d | khong co PDF: %d"
          % (len(lines) - len(nopdf), len(done), len(todo), len(nopdf)))

    from concurrent.futures import ThreadPoolExecutor
    cnt = [0]

    def work(ln):
        rows = fetch_one(pdf_url(ln["pdf_overview"]))
        with _lock:
            done[ln["code"]] = rows
            cnt[0] += 1
            print("  %4d/%d %-34s %2d thong so%s"
                  % (cnt[0], len(todo), ln["name"][:34], len(rows),
                     "" if rows else "   <-- PDF khong co muc nay"))
            if cnt[0] % 25 == 0:
                save(done)

    with ThreadPoolExecutor(max_workers=6) as ex:
        list(ex.map(work, todo))
    save(done)

    ok = sum(1 for v in done.values() if v)
    print("\nxong: %d dong co bang TECHNICAL DATA OVERVIEW / %d dong co PDF"
          % (ok, len(done)))


if __name__ == "__main__":
    main()
