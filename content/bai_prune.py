# -*- coding: utf-8 -*-
"""Xoa file bai dang cu khong con trong lan build nay.

Ten file bai la "<so thu tu> - <ten dong SP>.xlsx". Khi so bai hoac thu tu doi
(crawl lai, cao duoc them dong moi), lan build sau sinh ten khac ma file cu van
nam do -> thu muc giao hang co ca hai ban, khong biet ban nao dung. Da gap:

    * SICK: con "001 - Through-Beam Sensors.xlsx" cua mot lan chay thu
    * Balluff: build 113 bai roi build lai 150 bai -> sot 99 file cu

Builder goi prune(OUT_DIR, names) ngay sau khi ghi xong mucluc.
"""
import os


def prune(out_dir, names, index="00 MUC LUC.xlsx", quiet=False):
    """Giu lai dung `names` + file muc luc, xoa cac .xlsx con lai."""
    keep = set(names) | {index}
    gone = [f for f in os.listdir(out_dir)
            if f.endswith(".xlsx") and f not in keep and not f.startswith("~$")]
    for f in gone:
        try:
            os.remove(os.path.join(out_dir, f))
        except OSError as e:
            print("  [khong xoa duoc] %s: %s" % (f, e))
    if gone and not quiet:
        print("  da xoa %d file bai cu: %s"
              % (len(gone), ", ".join(sorted(gone)[:4])
                 + (" ..." if len(gone) > 4 else "")))
    return len(gone)
