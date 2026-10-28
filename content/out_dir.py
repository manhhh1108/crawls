# -*- coding: utf-8 -*-
"""Thu muc giao hang cua mot hang: ban dang lam hay ban da check.

Hang nao lam xong va da soat thi toan bo file duoc chuyen sang
`data/File đã crawls/Noi dung <hang>/`. Script nao van ghi cung vao
`data/Noi dung <hang>/` se sinh ra BAN THU HAI, roi hai ban lech nhau.

Da xay ra that: sau khi them anh logo cho cac bai khong co anh, 5 hang (IFM,
Leuze, Pepperl+Fuchs, Perma, wenglor) nam o ca hai cho - ban da check thi
thieu anh logo, ban moi thi thieu toan bo anh that.

Nen moi cho ghi file giao hang deu goi qua day: thu muc "da check" ton tai
thi ghi thang vao do, khong thi ghi vao thu muc dang lam.
"""
import os

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
DATA = os.path.join(REPO, "data")
DONE = os.path.join(DATA, "File đã crawls")


def brand_dir(outsub, create=True):
    """`outsub` la ten thu muc hang, vd 'Noi dung IFM'."""
    done = os.path.join(DONE, outsub)
    if os.path.isdir(done):
        return done
    live = os.path.join(DATA, outsub)
    if create:
        os.makedirs(live, exist_ok=True)
    return live


def img_dir(outsub, create=True):
    d = os.path.join(brand_dir(outsub, create), "images")
    if create:
        os.makedirs(d, exist_ok=True)
    return d


def where(outsub):
    """'da check' hay 'dang lam' - de in ra cho de doi chieu."""
    return "da check" if os.path.isdir(os.path.join(DONE, outsub)) \
        else "dang lam"
