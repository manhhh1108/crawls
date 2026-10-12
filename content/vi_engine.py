# -*- coding: utf-8 -*-
"""Dich EN->VI cho noi dung ky thuat, uu tien DUNG hon la muot.

Quy trinh 1 chuoi:
  1. tra tu dien (vi_terms) khop nguyen chuoi  -> dung luon, khong goi may dich
  2. mask cac token KHONG duoc dich: ma model, don vi, tieu chuan, ky hieu,
     giao thuc (IO-Link, PROFINET...), so
  3. mask cac CUM thuat ngu co trong tu dien -> khi tra lai thi dat ban dich dung
  4. phan con lai moi dua qua may dich
  5. go mask, chay bo luat sua loi (postfix), dinh dang so kieu Viet
  6. suspicious() cham diem nghi ngo de xuat file soat tay

Lo loi may dich hay mac ma bo luat nay chan:
  bus -> "xe buyt" | housing -> "nha o" | Art. -> "Nghe thuat" | drive -> "lai xe"
  bar -> "thanh"   | male/female -> "nam/nu" | 1,200 -> "1,200" (phai la 1.200)
"""
import os
import re
import sys

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
import vi_terms                      # noqa: E402
from translate import vi as _mt              # noqa: E402
from translate import vi_many as _mt_many    # noqa: E402

TERMS = vi_terms.terms()

# --------------------------------------------------------------------------
# 1. Token khong bao gio dich
# --------------------------------------------------------------------------
KEEP_WORDS = {
    "io-link", "profibus", "profinet", "ethernet", "ethernet/ip", "ethercat",
    "canopen", "canbus", "modbus", "devicenet", "ssi", "bissc", "biss-c",
    "hiperface", "hiperface dsl", "sinCos", "rs-232", "rs-422", "rs-485",
    "usb", "bluetooth", "wlan", "nfc", "rfid", "opc ua", "mqtt", "tcp/ip",
    "pnp", "npn", "push-pull", "ttl", "htl", "dc", "ac", "led", "laser",
    "sil", "sil2", "sil3", "silcl2", "silcl3", "pl", "ple", "pld", "plc",
    "mttf", "mttfd", "dcavg", "atex", "iecex", "rohs", "reach", "culus",
    "ukca", "eclass", "etim", "unspsc", "ip65", "ip66", "ip67", "ip68", "ip69k",
    "nema", "vistal", "abs", "pmma", "pom", "pvc", "tpe", "pur", "pa", "pbt",
    "sick", "bluepilot", "twineye", "linespot", "clearsens", "pinpoint",
    "smart task", "appspace", "multimode", "teach-turn", "dustalert",
}
# tieu chuan: EN 61000-6-2 / IEC 60529 / EN ISO 13849-1 / DIN 5033 ...
RX_STD = re.compile(
    r"\b(?:EN\s?ISO|ISO/IEC|EN|IEC|ISO|DIN|UL|CSA|ANSI|JIS|GB|VDE)\s?\d[\d\-:\.]*\b")
# ma model/part: co it nhat 1 chu HOA va 1 chu so, vd W16, WTM16F-24161120A00, M12
RX_CODE = re.compile(r"\b(?=[A-Za-z0-9/\-]*[A-Z])(?=[A-Za-z0-9/\-]*\d)[A-Z][A-Za-z0-9]*(?:[\-/][A-Za-z0-9]+)*\b")
# chi mask so "phuc tap" (co dau phan cach / phep nhan). So nguyen don le de
# nguyen cho may dich giu dung trat tu tu ("8 safety inputs" -> "8 ngo vao an toan").
RX_NUM = re.compile(r"\d[\d.,]*(?:\s?(?:x|×)\s?\d[\d.,]*)+|\d{1,3}(?:[.,]\d{3})+(?:\.\d+)?|\d+\.\d+")
RX_SYM = re.compile(r"[✔✓≤≥±Ø°µΩ×→…‰]|10\^\d+|\bm²\b|\bm³\b|[⁰¹²³⁴⁵⁶⁷⁸⁹⁻]+")

# ten rieng cua hang (dong SP, model...) - nap them luc chay bang add_keep_names()
_EXTRA_KEEP = set()


def add_keep_names(names):
    """Them ten rieng (dong SP / model) vao danh sach khong dich."""
    n = 0
    for x in names or ():
        x = (x or "").strip()
        if 2 <= len(x) <= 40 and not re.fullmatch(r"[\d\W]+", x):
            if x.lower() not in _EXTRA_KEEP:
                _EXTRA_KEEP.add(x.lower())
                n += 1
    if n:
        _scanner.rx = None        # buoc build lai regex ghep
    return n

_MASK = "ZQ%dQZ"
RX_MASK = re.compile(r"ZQ\s*(\d+)\s*QZ")

# Quet MOT LUOT bang 1 regex ghep: moi ky tu chi bi an di 1 lan, khong long nhau
# (neu mask nhieu luot thi RX_CODE/RX_NUM se an chinh cac ma mask truoc do).
_PHRASE_VI = {}


def _scanner(light=False):
    """light=True: chi an ma/tieu chuan/ten rieng (dung cho cau van dai - neu an
    qua nhieu thi may dich bo cuoc va tra lai nguyen tieng Anh)."""
    attr = "rx_light" if light else "rx"
    global _PHRASE_VI
    if getattr(_scanner, attr, None) is not None:
        return getattr(_scanner, attr)
    ph = [(k, v) for k, v in vi_terms.phrases() if len(k) >= 7]
    _PHRASE_VI = {k: v for k, v in ph}
    phr = "|".join(re.escape(k) for k, _ in ph)
    keepset = sorted(KEEP_WORDS | _EXTRA_KEEP, key=len, reverse=True)
    keep = "|".join(re.escape(w) for w in keepset)
    parts = [r"(?P<std>%s)" % RX_STD.pattern]
    if phr and not light:
        parts.append(r"(?P<phr>(?<![A-Za-z])(?:%s)(?![A-Za-z]))" % phr)
    if keep:
        parts.append(r"(?P<keep>(?<![A-Za-z0-9])(?:%s)(?![A-Za-z0-9]))" % keep)
    parts.append(r"(?P<sym>%s)" % RX_SYM.pattern)
    parts.append(r"(?P<code>%s)" % RX_CODE.pattern)
    if not light:
        parts.append(r"(?P<num>%s)" % RX_NUM.pattern)
    rx = re.compile("|".join(parts), re.I)
    setattr(_scanner, attr, rx)
    return rx


class Masker(object):
    def __init__(self):
        self.vals = []

    def add(self, text):
        self.vals.append(text)
        return _MASK % (len(self.vals) - 1)

    def mask_all(self, s, light=False):
        def f(m):
            if m.lastgroup == "phr":
                return self.add(_PHRASE_VI.get(m.group(0).lower(), m.group(0)))
            return self.add(m.group(0))
        return _scanner(light).sub(f, s)

    def restore(self, s):
        def f(m):
            i = int(m.group(1))
            return self.vals[i] if i < len(self.vals) else m.group(0)
        return RX_MASK.sub(f, s)


# --------------------------------------------------------------------------
# 2. Dinh dang so kieu Viet: 1,200 -> 1.200 ; 0.35 -> 0,35 (mot luot duy nhat)
# --------------------------------------------------------------------------
RX_LOCNUM = re.compile(
    r"(?<![\d.,])(?:(?P<th>\d{1,3}(?:,\d{3})+)(?P<thdec>\.\d+)?"
    r"|(?<![A-Za-z])(?P<dec>\d+\.\d+))(?![\d.,])")   # bo qua V1.1, IO-Link V1.1...


def localize_numbers(s):
    def f(m):
        if m.group("th"):
            out = m.group("th").replace(",", ".")
            if m.group("thdec"):
                out += "," + m.group("thdec")[1:]
            return out
        return m.group("dec").replace(".", ",")
    return RX_LOCNUM.sub(f, s)


# --------------------------------------------------------------------------
# 3. Sua loi may dich (chay sau cung)
# --------------------------------------------------------------------------
POSTFIX = [
    (re.compile(r"\bxe\s*buýt\b", re.I), "bus"),
    (re.compile(r"\bNhà\s*ở\b"), "Vỏ"),
    (re.compile(r"\bnhà\s*ở\b"), "vỏ"),
    (re.compile(r"\bNghệ\s*thuật\.?\s*(\d)"), r"Điều \1"),
    (re.compile(r"\bNghệ\s*thuật\b"), "Điều"),
    (re.compile(r"\bLái\s*xe\b"), "Truyền động"),
    (re.compile(r"\blái\s*xe\b"), "truyền động"),
    (re.compile(r"(\d)\s*thanh\b"), r"\1 bar"),
    (re.compile(r"\bgiống\s*đực\b", re.I), "đầu nối đực"),
    (re.compile(r"\bgiống\s*cái\b", re.I), "đầu nối cái"),
    (re.compile(r"(?<![\wÀ-ỹ])nam(?![\wÀ-ỹ])"), "ren ngoài"),
    (re.compile(r"(?<![\wÀ-ỹ])nữ(?![\wÀ-ỹ])"), "ren trong"),
    (re.compile(r"\bthép\s*không\s*gỉ\s*thép\b", re.I), "thép không gỉ"),
    (re.compile(r"\bsự\s+phản\s+xạ\s+lại\b", re.I), "phản xạ gương"),
    (re.compile(r"\bChìa\s*khóa\s*học\b", re.I), "Teach-in"),
    (re.compile(r"\bdạy\s*-?\s*trong\b", re.I), "Teach-in"),
    (re.compile(r"\bđộ\s*phân\s*giải\s*độ\s*phân\s*giải\b", re.I), "độ phân giải"),
    # may dich hay "dich" ca don vi -> tra ve ky hieu
    (re.compile(r"(\d)\s*mili\s*giây\b", re.I), r"\1 ms"),
    (re.compile(r"(\d)\s*micro\s*giây\b", re.I), r"\1 µs"),
    (re.compile(r"(\d)\s*nano\s*giây\b", re.I), r"\1 ns"),
    (re.compile(r"(\d)\s*mili\s*mét\b", re.I), r"\1 mm"),
    (re.compile(r"(\d)\s*mili\s*ampe\b", re.I), r"\1 mA"),
    (re.compile(r"(\d)\s*ki-?lô-?gam\b", re.I), r"\1 kg"),
    (re.compile(r"\bV(\d),(\d)"), r"V\1.\2"),   # phien ban: V1,1 -> V1.1
    (re.compile(r"\bDây/pin\b"), "Dây/chân"),
    (re.compile(r"\bdây/pin\b"), "dây/chân"),
    (re.compile(r"\s*\.\.\.\s*"), " ... "),     # SICK dung "a ... b" cho khoang gia tri
    (re.compile(r"\s+([,;:])"), r"\1"),
    (re.compile(r"\s{2,}"), " "),
]


def postfix(s):
    for rx, rep in POSTFIX:
        s = rx.sub(rep, s)
    return s.strip()


# --------------------------------------------------------------------------
# 4. API chinh
# --------------------------------------------------------------------------
RX_PIN = re.compile(r"\b(\d+)\s*-?\s*pin\b", re.I)


def lookup(en):
    """Khop nguyen chuoi trong tu dien (bo qua hoa thuong + dau cau cuoi)."""
    k = (en or "").strip()
    if not k:
        return ""
    v = TERMS.get(k.lower())
    if v:
        return v
    k2 = k.rstrip(":.").strip()
    return TERMS.get(k2.lower()) or ""


def _needs_mt(s):
    """Con chu cai Latin nao ngoai mask khong? Neu khong thi khoi dich."""
    return bool(re.search(r"[A-Za-z]{2,}", RX_MASK.sub(" ", s)))


def _prep(en, kind):
    s = (en or "").strip()
    if not s:
        return ("done", "")
    hit = lookup(s)
    if hit:
        return ("done", localize_numbers(hit) if kind == "value" else hit)
    s = RX_PIN.sub(lambda m: "%s chân" % m.group(1), s)
    ms = Masker()
    t = ms.mask_all(s, light=(kind == "prose"))
    return (("mt" if _needs_mt(t) else "local"), (ms, t))


def _fin(ms, out, kind):
    out = localize_numbers(postfix(ms.restore(out)))
    if kind == "key" and out:
        out = out[:1].upper() + out[1:]
    return out


def translate_many(items, kind="prose"):
    """Dich ca loat - gop request nen nhanh hon goi tung chuoi rat nhieu."""
    preps = [_prep(x, kind) for x in items]
    need = [p[1][1] for p in preps if p[0] == "mt"]
    got = iter(_mt_many(need) if need else [])
    out = []
    for st, payload in preps:
        if st == "done":
            out.append(payload)
        elif st == "local":
            out.append(_fin(payload[0], payload[1], kind))
        else:
            out.append(_fin(payload[0], next(got), kind))
    return out


def translate(en, kind="prose"):
    """kind: 'key' (ten thong so) | 'value' (gia tri) | 'prose' (cau van)."""
    return translate_many([en], kind)[0]


def vi_key(en):
    return translate(en, "key")


def vi_value(en):
    return translate(en, "value")


def vi_prose(en):
    return translate(en, "prose")


# --------------------------------------------------------------------------
# 5. Cham diem nghi ngo -> file soat tay
# --------------------------------------------------------------------------
BAD_WORDS = [
    "xe buýt", "nhà ở", "nghệ thuật", "lái xe", "giống đực", "giống cái",
    "chìa khóa học", "quán bar", "thanh bar", "cây gậy", "bộ phận sinh dục",
]
RX_EN_RUN = re.compile(r"\b(?:the|and|with|for|of|is|are|can|be|which|that|from)\b", re.I)
RX_LEFT_MASK = re.compile(r"ZQ\d+QZ|ZQ\s*\d+\s*QZ")
VI_CHARS = re.compile(r"[ăâđêôơưĂÂĐÊÔƠƯáàảãạắằẳẵặấầẩẫậéèẻẽẹếềểễệíìỉĩịóòỏõọốồổỗộớờởỡợúùủũụứừửữựýỳỷỹỵ]")


def suspicious(en, vn):
    """Tra ve ly do can soat tay, hoac '' neu nhin on."""
    v = (vn or "").strip()
    e = (en or "").strip()
    if not v:
        return "trong"
    if RX_LEFT_MASK.search(v):
        return "sot ma mask"
    low = v.lower()
    for b in BAD_WORDS:
        if b in low:
            return "tu nghi sai: " + b
    words = re.findall(r"[A-Za-z]{3,}", v)
    if len(words) >= 4 and RX_EN_RUN.search(v) and not VI_CHARS.search(v):
        return "con nguyen tieng Anh"
    if len(e) > 40 and v == e and len(_real_words(e)) >= 2:
        return "chua dich"      # giong het ban EN va van con tu that su can dich
    return ""


# don vi / ky hieu: xuat hien nguyen van la binh thuong, khong tinh la "chua dich"
UNITS = set("""mm cm dm m km um µm nm inch ft kg g mg t ms us µs ns s min h hz khz mhz ghz
v mv kv a ma ua w kw mw kva var nm ncm bar mbar psi pa kpa mpa lx klx db mt ut gauss
byte bytes bit bits kbit mbit gbit baud rpm fps ppr cpr px mp dpi gb mb kb tb ohm
ip nema sil pl type class cat pro plus max min typ avg nom ref art
and or of for the with to in on at by""".split())


def _real_words(s):
    """Tu tieng Anh THUC SU can dich (bo don vi, acronym, ma, ten rieng)."""
    out = []
    for w in re.findall(r"[A-Za-z][A-Za-z'\-]{2,}", RX_MASK.sub(" ", Masker().mask_all(s))):
        lw = w.lower()
        if lw in UNITS or lw in KEEP_WORDS or lw in _EXTRA_KEEP:
            continue
        if w.isupper():                      # acronym: PROM, AIDA, ATEX...
            continue
        out.append(w)
    return out


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    tests = [
        ("Bus termination", "key"),
        ("Housing", "key"),
        ("IP67, housing side (IEC 60529)", "value"),
        ("Information according to Art. 3 of the Data Act", "key"),
        ("1,200 mm (mode 1, 3, 4, 5)", "value"),
        ("≤ 6,000 min⁻¹", "value"),
        ("Blind hollow shaft", "value"),
        ("30 g, 10 Hz ... 2,000 Hz (EN 60068-2-6)", "value"),
        ("Stainless steel", "value"),
        ("The W16 sensors deliver reliable detection thereby reducing unplanned "
         "machine downtimes.", "prose"),
    ]
    for t, k in tests:
        print("%-8s %-62s -> %s" % (k, t[:62], translate(t, k)))
