# -*- coding: utf-8 -*-
"""Ap ban cap nhat tu dien EN-VI vao DU LIEU DA CRAWL (khong phai cao lai).

    python apply_tudien.py --xem     # chi liet ke, khong sua
    python apply_tudien.py           # sua that (co sao luu .bak)

Sau do chay lai build_<hang>_excel.py cho cac hang bi doi (script in ra san).

VI SAO KHONG CAO LAI
Tu dien chi anh huong khau DICH. Cao lai ca 17 hang mat nhieu gio, phu thuoc
mang, va cache dich danh theo sha1 cua chuoi tieng Anh nen phan lon se tra ve
y nguyen ban cu - tuc la van khong ap duoc tu moi. Sua thang tren content.json
thi chinh xac, nhanh, va doi chieu duoc tung cau.

CACH CHAN DICH SAI
Moi luat co hai tang:

* `toan_cuc` - nhung cum CHAC CHAN la nghia can sua, thay o moi truong `*_vi`.
  Vi du "Cam bien chac" chac chan la fork sensor.
* `cap` - nhung tu DE NHAM, chi thay khi cau tieng Anh goc (truong `*_en`,
  hoac `name`/`vals`/`desc` trong `items`) co dung tu khoa do.

Khong co tang hai thi hong ngay: Festo dich "clevis" (khop chac) thanh "chac",
"Mat bich chac" ma doi thanh "Mat bich kim" la sai. Tuong tu "cell" -> "cell"
khong duoc dong vao "pin mat troi" (solar cell), va "cam bien cam ung" cua
"touch sensor" khong phai "inductive sensor".
"""
import io
import json
import os
import re
import shutil
import sys

sys.stdout.reconfigure(encoding="utf-8")

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "out")

# ---------------------------------------------------------------------------
# Luat - moi lan khach sua tu dien thi them/bot o day.
# Dot cap nhat 17/11: 10 dong moi trong data/Tu dien dich EN-VI.xlsx.
# ---------------------------------------------------------------------------
KIM = "dạng kìm"
LUAT = [
    {
        "ten": "fork -> dạng kìm",
        "toan_cuc": [
            ("Cảm biến chạc", "Cảm biến " + KIM),
            ("cảm biến chạc", "cảm biến " + KIM),
            ("Cảm biến dạng chữ U", "Cảm biến " + KIM),
            ("cảm biến dạng chữ U", "cảm biến " + KIM),
            ("Cảm biến dạng fork", "Cảm biến " + KIM),
            ("cảm biến dạng fork", "cảm biến " + KIM),
        ],
        # "chac" tran chi doi khi CA BAI noi ve fork sensor. Khong co chan
        # nay thi "Helmet mount attachement forks" -> "Kim gan mu bao ho" va
        # "mounting fork" cua gia do encoder -> "kim lap", deu sai: do la chac
        # CO KHI, khong phai cam bien.
        # Khong dong vao "chu U" tran: no dich "U design", vd "thiet ke chu U,
        # chieu rong khe 5 mm" cua Pepperl+Fuchs - dich dung roi.
        "en_bai": r"fork\s+sensors?",
        "en": r"(?<![A-Za-z])forks?(?![A-Za-z])",
        "cap": [
            ("Chạc", "Kìm"),
            ("chạc", "kìm"),
        ],
    },
    {
        "ten": "cell -> giữ nguyên \"cell\"",
        "toan_cuc": [],
        "en": r"(?<![A-Za-z])cells?(?![A-Za-z])",
        "cap": [("Tế bào", "Cell"), ("tế bào", "cell")],
    },
    {
        "ten": "line scanner -> giữ nguyên",
        "toan_cuc": [],
        "en": r"line\s+scanners?",
        "cap": [
            ("Máy quét mã vạch dạng dòng",
             "Line scanner mã vạch"),
            ("máy quét mã vạch dạng dòng",
             "line scanner mã vạch"),
            ("Máy quét dòng", "Line scanner"),
            ("máy quét dòng", "line scanner"),
            ("Máy quét đường", "Line scanner"),
            ("máy quét đường", "line scanner"),
        ],
    },
    {
        "ten": "inductive sensor -> cảm biến tiệm cận điện cảm",
        "toan_cuc": [],
        # Chan bang "inductive" tran chu khong phai "inductive sensor": trang
        # hang viet nhieu kieu - "Inductive analogue sensors", "Inductive
        # sensors for ferrous metal detection", "inductive active principle".
        # Van an toan vi "cam bien cam ung" cua TOUCH sensor thi cau tieng Anh
        # khong he co chu "inductive".
        "en": r"(?<![A-Za-z])inductive(?![A-Za-z])",
        "cap": [
            ("Cảm biến cảm ứng",
             "Cảm biến tiệm cận điện cảm"),
            ("cảm biến cảm ứng",
             "cảm biến tiệm cận điện cảm"),
        ],
    },
    {
        "ten": "swivel mounting -> gá lắp đặt xoay",
        "toan_cuc": [],
        "en": r"swivel\s+mounting",
        # Phai dung regex: thay chuoi thuong thi "Lap dat xoay" -> "Ga lap dat
        # xoay", roi luat chu thuong lai an vao chinh ket qua do -> "Ga ga lap
        # dat xoay". Lookbehind chan khong cho them "ga" lan hai.
        "cap": [],
        "cap_re": [
            (r"(?<![Gg]á )Lắp đặt xoay",
             "Gá lắp đặt xoay"),
            (r"(?<![Gg]á )lắp đặt xoay",
             "gá lắp đặt xoay"),
        ],
    },
    {
        # Khach chi doi dong "moderate", nhung ca ho CRC nam CUNG MOT O
        # ("0 - ...; 1 - ...; 2 - ..."), de lan lon hai cach goi thi doc rat
        # ky. Nen doi dong bo "ung suat an mon" -> "muc do an mon".
        "ten": "corrosion stress -> mức độ ăn mòn",
        "toan_cuc": [],
        "en": r"corrosion\s+stress",
        "cap": [
            ("Không có ứng suất ăn mòn",
             "Không bị ăn mòn"),
            ("không có ứng suất ăn mòn",
             "không bị ăn mòn"),
            ("Ứng suất ăn mòn trung bình",
             "Mức độ ăn mòn trung tính"),
            ("ứng suất ăn mòn trung bình",
             "mức độ ăn mòn trung tính"),
            ("Ứng suất ăn mòn", "Mức độ ăn mòn"),
            ("ứng suất ăn mòn", "mức độ ăn mòn"),
        ],
    },
]
# ---------------------------------------------------------------------------
# Dot 20/11 - khach review file giao, bat ra may cho con sot.
# ---------------------------------------------------------------------------
MAN = "Màn chắn sáng"
LUAT += [
    {
        # Leuze "Fork photoelectric sensor" -> "Cam bien quang dang khe",
        # Festo "Fork light barrier" cung vay. Nhung "dang khe" con dich tu
        # SLOT sensor (Pepperl+Fuchs "slot grid sensor", ESPE "Slot-type
        # sensor") - cho do dung la "khe", nen phai chan nguoc chu "slot".
        "ten": "fork -> dạng kìm (thêm \"dạng khe\")",
        "toan_cuc": [],
        "en": r"(?<![A-Za-z])forks?(?![A-Za-z])",
        "en_khong": r"slot",
        "cap": [("dạng khe", KIM)],
    },
    {
        "ten": "line profile sensor -> cảm biến quét biên dạng",
        "toan_cuc": [
            ("Cảm biến biên dạng đường",
             "Cảm biến quét biên dạng"),
            ("cảm biến biên dạng đường",
             "cảm biến quét biên dạng"),
        ],
        "en": r"line\s+profile",
        "cap": [],
    },
    {
        "ten": "optical positioning sensor -> xác định vị trí",
        "toan_cuc": [
            ("Cảm biến định vị quang học",
             "Cảm biến xác định vị trí bằng quang học"),
            ("cảm biến định vị quang học",
             "cảm biến xác định vị trí bằng quang học"),
            ("Hệ thống định vị quang học",
             "Hệ thống xác định vị trí bằng quang học"),
        ],
        # Leuze con viet "Optical laser positioning systems" - cung mot ho,
        # cho phep mot tu chen giua.
        "en": r"optical\s+(?:\w+\s+)?positioning",
        "cap": [],
    },
    {
        # "ren quang" / "luoi quang" deu phai ve "man chan sang" KHI tieng Anh
        # la light curtain. Chan nguoc "light grid": LIGHT GRID la mat hang
        # KHAC (IFM "Luoi quang an toan dong OY", SICK "Automation light
        # grids"), dich "luoi quang" la dung, doi la sai hang.
        # "Rem quang" LUON la light curtain - khong can chan nguoc. Nhieu
        # dong cua Pepperl+Fuchs noi ca hai mat hang trong cung mot cau
        # ("for safety light grids SLP, for safety light curtains SLC"), chan
        # nguoc o day la bo sot ca dong.
        "ten": "light curtain -> màn chắn sáng",
        "toan_cuc": [],
        "en": r"light\s+curtains?",
        "cap": [("Rèm quang", MAN), ("rèm quang", MAN.lower())],
    },
    {
        # Rieng "Luoi quang" thi phai chan: LIGHT GRID la mat hang KHAC va
        # dich "luoi quang" moi dung (IFM "Luoi quang an toan dong OY").
        "ten": "light curtain -> màn chắn sáng (lưới quang)",
        "toan_cuc": [],
        "en": r"light\s+curtains?",
        "en_khong": r"light\s+grid|safety\s+grid|switching\s+grid",
        "cap": [("Lưới quang", MAN),
                ("lưới quang", MAN.lower())],
    },
    {
        # Ten DANH MUC Ntrend - khong co ban tieng Anh di kem nen khong chan
        # bang EN duoc, phai gioi han theo TRUONG. Chi `cat_vi` moi co cum
        # nay; cho khac "Luoi quang an toan" la light grid that.
        "ten": "danh mục: Lưới quang an toàn -> Màn chắn sáng an toàn",
        "toan_cuc": [("Lưới quang an toàn",
                      MAN + " an toàn")],
        "chi_truong": ("cat_vi",),
        "en": r"$^",
        "cap": [],
    },
    {
        "ten": "screw terminal -> cầu đấu vít",
        # Trong tieng Viet "dau cuc vit" khong con nghia nao khac ngoai screw
        # terminal, nen de o toan_cuc - con bat duoc ca nhung dong KHONG co
        # ban tieng Anh di kem (vd mo ta contactor TeSys cua Schneider).
        "toan_cuc": [
            ("Đầu cực vít", "Cầu đấu vít"),
            ("đầu cực vít", "cầu đấu vít"),
        ],
        # Pilz lay thang tieng Duc tu web hang: Schraubklemme = screw terminal.
        # Khong chi "screw terminal": Pilz viet "plug-in terminals of screw
        # type", SICK "screw- and spring-loaded terminals", Schneider
        # "screwed terminal connections". Chan = cau nao co CA "screw" lan
        # "terminal" (hoac tieng Duc Schraubklemme).
        "en": r"^(?=[\s\S]*screw)(?=[\s\S]*terminal)|schraubklemme",
        "cap": [
            ("Đầu cực vít", "Cầu đấu vít"),
            ("đầu cực vít", "cầu đấu vít"),
        ],
    },
    {
        # IFM: may dich bo sot han chu "closed", de nguyen tieng Anh giua cau
        # tieng Viet. 1.220 chuoi, hien ra file giao.
        "ten": "IFM: \"/ closed\" còn tiếng Anh",
        "toan_cuc": [],
        "en": r"normally\s+open\s*/\s*closed",
        "cap": [
            ("Thường mở (NO) / closed",
             "Thường mở (NO) / Thường đóng (NC)"),
            ("Thường mở (NO) / đóng",
             "Thường mở (NO) / Thường đóng (NC)"),
        ],
    },
]

for L in LUAT:
    L["rx_en"] = re.compile(L["en"], re.I)
    L["rx_khong"] = (re.compile(L["en_khong"], re.I)
                     if L.get("en_khong") else None)
    L["truong"] = L.get("chi_truong")
    L["rx_bai"] = re.compile(L["en_bai"], re.I) if L.get("en_bai") else None
    L["rx_cap"] = [(re.compile(p), r) for p, r in L.get("cap_re", ())]

# ten thu muc out -> script dung Excel
BUILD = {
    "01 SICK": "build_sick_excel.py",
    "02 Pepperl+Fuchs": "build_pepperl_excel.py",
    "03 Reer": "build_reer_excel.py",
    "04 IFM": "build_ifm_excel.py",
    "05 Perma": "build_content_excel.py",
    "06 wenglor": "build_wenglor_excel.py",
    "07 Mitsubishi Electric": "build_mitsubishi_excel.py",
    "08 Leuze": "build_leuze_excel.py",
    "09 Pilz": "build_pilz_excel.py",
    "10 Festo": "build_festo_excel.py",
    "11 Schneider": "build_schneider_excel.py",
    "12 Supmea": "build_supmea_excel.py",
    "13 Tree Electric": "build_tree_excel.py",
    "14 Xingguang-Starshine": "build_xingguang_excel.py",
    "15 ESPE Technology": "build_espe_excel.py",
    "16 OTENNLUX Lighting": "build_otennlux_excel.py",
    "17 Balluff": "build_balluff_excel.py",
}


# ---------------------------------------------------------------------------
def tieng_anh(d, kv):
    """Chuoi tieng Anh di cung mot khoa `*_vi` trong cung mot dict.

    Rieng `name_vi` phai gop CA `title_en` LAN `name`: nhieu hang dich ten bai
    tu `title_en`, con `name` chi la ten muc ngan tren menu. Vi du IFM co
    name="AC sensors" nhung title_en="Inductive sensors with combined AC/DC
    output stage" - chi tra `name` thi khong thay chu "inductive" dau.
    """
    g = kv[:-3]
    if g == "name":
        xs = [d.get(k) for k in ("title_en", "name_en", "name")]
        return " ".join(x for x in xs if isinstance(x, str)) or None
    for k in (g + "_en", g):
        if k in d:
            return d[k]
    return None


def _ghep(e, v, chu, kh, truong):
    """Di song song chuoi EN va chuoi VI, ke ca khi LONG NHIEU TANG.

    Bang thong so cua Reer/SICK/OTENNLUX luu dang [[ten, gia tri], ...] chu
    khong phai list chuoi phang. Bo qua tang long thi sot ca bang thong so.
    `truong` la ten truong `*_vi` goc, giu nguyen qua moi tang long, de luat
    nao can gioi han theo truong (vd chi sua `cat_vi`) con biet duong.
    """
    if isinstance(v, str):
        yield (e if isinstance(e, str) else None), chu, kh, truong
    elif isinstance(v, list):
        for i, x in enumerate(v):
            ei = e[i] if isinstance(e, list) and i < len(e) else None
            for t in _ghep(ei, x, v, i, truong):
                yield t


def duyet(o, nhan=""):
    """Sinh ra (en_text, chu_so_huu, khoa, ten_truong) cho moi chuoi VI.

    `en_text` = chuoi tieng Anh tuong ung, hoac None neu khong co doi tac.
    Sua bang `chu_so_huu[khoa] = chuoi_moi`.
    """
    if isinstance(o, dict):
        for kv, v in list(o.items()):
            if kv.endswith("_vi"):
                for t in _ghep(tieng_anh(o, kv), v, o, kv, kv):
                    yield t
            elif isinstance(v, (dict, list)):
                for t in duyet(v, kv):
                    yield t
    elif isinstance(o, list):
        for x in o:
            if isinstance(x, (dict, list)):
                for t in duyet(x, nhan):
                    yield t


def main():
    xem = "--xem" in sys.argv
    tong = {}
    mau = {}
    doi_hang = []

    for ten in sorted(BUILD):
        p = os.path.join(OUT, ten, "content.json")
        if not os.path.exists(p):
            continue
        doc = json.load(io.open(p, encoding="utf-8"))
        n_hang = {}
        for a in doc["articles"]:
            cap_bai = list(duyet(a))
            # Toan bo tieng Anh cua bai - dung cho chan muc BAI (en_bai).
            blob = " ".join(e for e, _c, _k, _t in cap_bai if e)
            for en, chu, kh, truong in cap_bai:
                s0 = chu[kh]
                s = s0
                for L in LUAT:
                    if L["truong"] and truong not in L["truong"]:
                        continue
                    for cu, moi in L["toan_cuc"]:
                        if cu in s:
                            s = s.replace(cu, moi)
                            n_hang[L["ten"]] = n_hang.get(L["ten"], 0) + 1
                    if L["rx_bai"] is not None and not L["rx_bai"].search(blob):
                        continue
                    if en and L["rx_khong"] is not None                             and L["rx_khong"].search(en):
                        continue
                    if en and L["rx_en"].search(en):
                        for cu, moi in L["cap"]:
                            if cu in s:
                                s = s.replace(cu, moi)
                                n_hang[L["ten"]] = n_hang.get(L["ten"], 0) + 1
                        for rx, moi in L["rx_cap"]:
                            s, k = rx.subn(moi, s)
                            if k:
                                n_hang[L["ten"]] = n_hang.get(L["ten"], 0) + 1
                if s != s0:
                    chu[kh] = s
                    k = (ten, kh if isinstance(kh, str) else "list")
                    if len(mau.get(k, [])) < 2:
                        mau.setdefault(k, []).append((s0[:120], s[:120]))

        if n_hang:
            doi_hang.append(ten)
            tong[ten] = n_hang
            if not xem:
                if not os.path.exists(p + ".bak"):
                    shutil.copy(p, p + ".bak")
                with io.open(p, "w", encoding="utf-8") as f:
                    json.dump(doc, f, ensure_ascii=False)

    print("=" * 72)
    for ten in doi_hang:
        print("\n%s" % ten)
        for k, v in sorted(tong[ten].items()):
            print("    %-46s %4d chỗ" % (k, v))
        for (h, kh), ms in sorted(mau.items()):
            if h != ten:
                continue
            for a, b in ms:
                print("      [%s]" % kh)
                print("        cu : %s" % a)
                print("        mới: %s" % b)

    print("\n" + "=" * 72)
    print("%d hãng bị đổi | tổng %d chỗ"
          % (len(doi_hang), sum(sum(v.values()) for v in tong.values())))
    if xem:
        print("(--xem: chưa ghi gì)")
        return
    print("\nChạy tiếp để dựng lại Excel:")
    for ten in doi_hang:
        print("    python %s" % BUILD[ten])


if __name__ == "__main__":
    main()
