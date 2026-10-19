# -*- coding: utf-8 -*-
"""Dich EN->VI bang LLM (Gemini) - thay cho translate.googleapis.com.

Khac biet chinh so voi may dich thuong: dua duoc NGU CANH + TU DIEN vao prompt,
nen nhung cau chua tung gap cung dich dung nganh thay vi sai bua.

Dung:
    export GEMINI_API_KEY=...        (hoac dat trong file .env o goc repo)
    python llm_translate.py --models            # xem cac model dung duoc
    python llm_translate.py "Housing material"  # dich thu

API dung REST qua curl (giong cach shared/fetcher lam, khong them thu vien).
"""
import io
import json
import os
import re
import sys
import urllib.error
import urllib.request

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)

# Hang muc free tier tinh RIENG cho tung model VA tung nha cung cap, nen khi
# model hien tai het quota (429) hoac qua tai (503) thi tu nhay sang muc sau.
# Thu tu: Gemini truoc (tieng Viet tot nhat), Groq lam du phong.
ENDPOINTS = [
    ("gemini", "gemini-3.8-flash"),
    ("gemini", "gemini-3.1-flash-lite"),
    ("gemini", "gemini-flash-lite-latest"),
    ("gemini", "gemini-3.5-flash"),
    ("gemini", "gemini-3.6-flash"),
    ("gemini", "gemini-3-flash-preview"),
    ("groq", "GROQ_MODEL_1"),          # dien bang model that luc chay
    ("groq", "GROQ_MODEL_2"),
]
GEMINI_BASE = "https://generativelanguage.googleapis.com/v1beta"
GROQ_BASE = "https://api.groq.com/openai/v1"
_cur = 0
MAX_RETRY = 3
# Han muc free cua Gemini tinh theo PHUT: dinh 429 thi nghi 60s roi dung lai
# duoc, khong nen bo han model tot nhat.
COOLDOWN = 60
MAX_ROUNDS = 60
_cooldown = [0.0] * 32          # thoi diem moi muc duoc dung lai


def _keys(prefix):
    """Danh sach API key: PREFIX, PREFIX_2, PREFIX_3... hoac 1 bien ngan cach ','.

    Moi key co han muc free rieng -> het key nay thi tu nhay sang key sau.
    """
    _load_dotenv()
    out = []
    for name in [prefix] + ["%s_%d" % (prefix, i) for i in range(2, 9)]:
        v = (os.environ.get(name) or "").strip()
        for part in v.split(","):
            part = part.strip()
            if part and part not in out:
                out.append(part)
    return out


def _env(name):
    _load_dotenv()
    return os.environ.get(name)


def groq_models():
    """Model chat cua Groq, uu tien ban lon (dich tot hon)."""
    ks = _keys("GROQ_API_KEY")
    if not ks:
        return []
    k = ks[0]
    d = _http(GROQ_BASE + "/models", None, 60, {"Authorization": "Bearer " + k})
    ids = [m["id"] for m in d.get("data", []) if "whisper" not in m["id"]
           and "tts" not in m["id"] and "guard" not in m["id"]]
    ids.sort(key=lambda x: (0 if "120b" in x or "70b" in x or "qwen" in x.lower()
                            else 1, x))
    return ids


def _resolve_endpoints():
    """Trai ENDPOINTS thanh (nha cung cap, model, key) cho TUNG key co trong .env.

    Han muc free tinh rieng cho tung key, nen 2 key Groq = gap doi han muc.
    """
    global ENDPOINTS
    if not any(len(e) == 2 and str(e[1]).startswith("GROQ_MODEL")
               for e in ENDPOINTS):
        return
    try:
        real = groq_models()[:2]
    except Exception:
        real = []
    gk = _keys("GEMINI_API_KEY") or _keys("GOOGLE_API_KEY") or [""]
    qk = _keys("GROQ_API_KEY") or [""]
    out = []
    for p, m in [e[:2] for e in ENDPOINTS if not str(e[1]).startswith("GROQ_MODEL")]:
        for k in gk:
            out.append((p, m, k))
    for k in qk:
        for m in real:
            out.append(("groq", m, k))
    ENDPOINTS = out
    global _cooldown
    if len(_cooldown) < len(ENDPOINTS):
        _cooldown += [0.0] * (len(ENDPOINTS) - len(_cooldown))


# ---------------------------------------------------------------------------
# API key
# ---------------------------------------------------------------------------
def _load_dotenv():
    p = os.path.join(REPO, ".env")
    if not os.path.exists(p):
        return
    for line in io.open(p, encoding="utf-8"):
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        k, v = line.split("=", 1)
        os.environ.setdefault(k.strip(), v.strip().strip('"').strip("'"))


def api_key():
    _load_dotenv()
    k = os.environ.get("GEMINI_API_KEY") or os.environ.get("GOOGLE_API_KEY")
    if not k:
        raise RuntimeError(
            "Chua co GEMINI_API_KEY. Tao file .env o goc repo voi dong:\n"
            "    GEMINI_API_KEY=<key lay o aistudio.google.com>")
    return k


def available():
    """Co key hay khong - de translate.py tu chon backend."""
    try:
        api_key()
        return True
    except Exception:
        return False


# ---------------------------------------------------------------------------
# Goi API
# ---------------------------------------------------------------------------
def _http(url, data=None, timeout=90, extra_headers=None):
    """Goi HTTP bang urllib (KHONG dung subprocess curl - curl goi tu nhieu luong
    tren Windows/Git Bash bi deadlock o duong ong, da lam treo lan chay 14/10)."""
    req = urllib.request.Request(
        url, data=json.dumps(data).encode("utf-8") if data is not None else None,
        headers=dict({"Content-Type": "application/json",
                      # Groq chan User-Agent mac dinh cua urllib o tang Cloudflare
                      # (HTTP 403 error code 1010) -> phai dat UA trinh duyet
                      "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                                    "AppleWebKit/537.36 (KHTML, like Gecko) "
                                    "Chrome/131.0.0.0 Safari/537.36",
                      "Accept": "application/json"},
                     **(extra_headers or {})))
    try:
        with urllib.request.urlopen(req, timeout=timeout) as r:
            return json.loads(r.read().decode("utf-8", "replace"))
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", "replace")
        try:
            return json.loads(body)
        except Exception:
            raise RuntimeError("HTTP %s: %s" % (e.code, body[:300]))


def _call_gemini(model, system, prompt, n, key=None):
    d = _http("%s/models/%s:generateContent?key=%s" % (GEMINI_BASE, model, key or api_key()),
              {"systemInstruction": {"parts": [{"text": system}]},
               "contents": [{"role": "user", "parts": [{"text": prompt}]}],
               "generationConfig": {"temperature": 0,
                                    "responseMimeType": "application/json",
                                    "responseSchema": {"type": "ARRAY",
                                                       "items": {"type": "STRING"}}}})
    if "error" in d:
        raise RuntimeError("Gemini loi: %s" % json.dumps(d["error"])[:300])
    return d["candidates"][0]["content"]["parts"][0]["text"]


def _call_groq(model, system, prompt, n, key=None):
    """Groq dung chuan OpenAI. Khong co responseSchema -> ep JSON object
    {"items":[...]} vi json_object mode cua ho khong nhan mang o cap cao nhat."""
    k = key or (_keys("GROQ_API_KEY") or [""])[0]
    if not k:
        raise RuntimeError("chua co GROQ_API_KEY")
    d = _http(GROQ_BASE + "/chat/completions",
              {"model": model, "temperature": 0, "max_tokens": 16000,
               "response_format": {"type": "json_object"},
               "messages": [
                   {"role": "system", "content": system},
                   {"role": "user", "content": prompt +
                    '\n\nTra ve JSON dang {"items": [<%d chuoi da dich>]}' % n}]},
              90, {"Authorization": "Bearer " + k})
    if "error" in d:
        raise RuntimeError("Groq loi: %s" % json.dumps(d["error"])[:300])
    return d["choices"][0]["message"]["content"]


CALLERS = {"gemini": _call_gemini, "groq": _call_groq}


def _key_no(e):
    """Key thu may cua nha cung cap do - chi de in log cho de theo doi."""
    if len(e) < 3:
        return 1
    ks = _keys("GEMINI_API_KEY") if e[0] == "gemini" else _keys("GROQ_API_KEY")
    return (ks.index(e[2]) + 1) if e[2] in ks else 1


def _parse(txt, n):
    got = json.loads(txt)
    if isinstance(got, dict):
        for key in ("items", "translations", "result", "data"):
            if isinstance(got.get(key), list):
                got = got[key]
                break
    if not isinstance(got, list) or len(got) != n:
        raise RuntimeError("lech so dong: can %d, nhan %s"
                           % (n, len(got) if isinstance(got, list) else type(got).__name__))
    return [str(x).strip() for x in got]


def list_models():
    d = _http("%s/models?key=%s" % (BASE, api_key()), None, 60)
    if "error" in d:
        raise RuntimeError(json.dumps(d["error"])[:300])
    return [m["name"].replace("models/", "")
            for m in d.get("models", [])
            if "generateContent" in (m.get("supportedGenerationMethods") or [])]


# ---------------------------------------------------------------------------
# Prompt
# ---------------------------------------------------------------------------
SYSTEM = """Bạn là biên dịch viên kỹ thuật Anh–Việt chuyên ngành tự động hóa công nghiệp (cảm biến, encoder, thiết bị an toàn máy, đọc mã vạch, RFID, cáp và đầu nối).
Văn bản nguồn lấy từ catalog sản phẩm của các hãng như SICK, Pepperl+Fuchs, IFM, Festo; bản dịch dùng để đăng lên website bán hàng tại Việt Nam.

QUY TẮC BẮT BUỘC:
1. Viết tiếng Việt CÓ DẤU đầy đủ, đúng chính tả. Không bao giờ trả về tiếng Việt không dấu.
2. GIỮ NGUYÊN, không dịch, không đổi hoa thường:
   - mã sản phẩm / model (W16, DFS60B-S4PA10000, IME2S, microScan3, deTec)
   - đơn vị đo (mm, mT, V DC, Hz, IP67), con số, ký hiệu
   - tên tiêu chuẩn (EN ISO 13849-1, IEC 60529, ATEX, SIL2, PL e)
   - tên giao thức / công nghệ (IO-Link, PROFINET, EtherNet/IP, CANopen, RS-485)
   - tên riêng của hãng và dòng sản phẩm
   - các token dạng ⟦...⟧ hoặc {{...}} nếu có — copy y nguyên
3. Dịch theo nghĩa NGÀNH CÔNG NGHIỆP, không dịch nghĩa đen:
   housing = vỏ (KHÔNG phải "nhà ở") | bus = bus truyền thông (KHÔNG phải "xe buýt")
   drive = truyền động (KHÔNG phải "ổ đĩa"/"lái xe") | Art. = mã hàng (KHÔNG phải "nghệ thuật")
   male/female = đực/cái của đầu nối (KHÔNG phải "nam/nữ")
   light curtain = màn chắn sáng an toàn | protection class IP.. = cấp bảo vệ IP..
   (IP là chống bụi/nước, KHÔNG phải "cấp cách điện")
4. Văn phong kỹ thuật, ngắn gọn, tự nhiên như tài liệu kỹ thuật tiếng Việt.
   Không thêm lời giải thích, không thêm dấu câu thừa.
5. Số thập phân theo kiểu Việt Nam: 1,200 → 1.200 ; 0.5 → 0,5
6. Nếu một chuỗi vốn là mã hoặc tên riêng thuần túy thì trả về y nguyên chuỗi đó."""

KIND_HINT = {
    "key": "Cac chuoi duoi day la TEN THONG SO ky thuat (nhan cot trong bang). "
           "Dich ngan gon, dung thuat ngu chuan, khong thanh cau.",
    "value": "Cac chuoi duoi day la GIA TRI thong so. Nhieu gia tri la ma/ky hieu "
             "-> giu nguyen. Chi dich phan la tu tieng Anh thuong.",
    "prose": "Cac chuoi duoi day la CAU VAN mo ta san pham. Dich thanh cau tieng "
             "Viet hoan chinh, muot, dung thuat ngu nganh.",
}


def _glossary_for(items, terms, limit=80):
    """Chi nhet nhung tu dien THUC SU xuat hien trong lo nay (tiet kiem token)."""
    blob = " ".join(items).lower()
    hit = [(en, vi) for en, vi in terms.items()
           if len(en) >= 3 and en in blob]
    hit.sort(key=lambda kv: -len(kv[0]))
    return hit[:limit]


def build_prompt(items, kind, terms):
    gl = _glossary_for(items, terms)
    parts = [KIND_HINT.get(kind, KIND_HINT["prose"])]
    if gl:
        parts.append("\nTU DIEN THUAT NGU BAT BUOC DUNG (EN = VI):")
        parts += ["  %s = %s" % (en, vi) for en, vi in gl]
    parts.append("\nDich %d chuoi sau sang tieng Viet. Tra ve DUNG mot mang JSON "
                 "gom %d chuoi, cung thu tu, khong kem gi khac."
                 % (len(items), len(items)))
    parts.append(json.dumps(items, ensure_ascii=False, indent=1))
    return "\n".join(parts)


# ---------------------------------------------------------------------------
# Dich
# ---------------------------------------------------------------------------
# So chuoi toi da 1 request. De GIONG NHAU cho moi nha cung cap: giua chung co
# the phai chuyen tu Gemini sang Groq, neu lo dang to (250) gui sang Groq thi
# treo im lang (da gap 14/10, ket 50 phut). 60 chuoi/request chay ~4 giay,
# ca bo 14.448 chuoi het ~241 request - van thoai mai trong moi han muc free.
CHUNK = {"gemini": 60, "groq": 60}


def chunk_size():
    _resolve_endpoints()
    return CHUNK.get(ENDPOINTS[_cur][0], 60)


def translate_batch(items, kind="prose", terms=None):
    """Dich 1 lo chuoi -> list cung do dai. Nem Exception neu that bai."""
    items = [str(x) for x in items]
    if not items:
        return []
    if terms is None:
        sys.path.insert(0, HERE)
        import vi_terms
        terms = vi_terms.terms()

    payload = {
        "systemInstruction": {"parts": [{"text": SYSTEM}]},
        "contents": [{"role": "user",
                      "parts": [{"text": build_prompt(items, kind, terms)}]}],
        "generationConfig": {
            "temperature": 0,
            "responseMimeType": "application/json",
            "responseSchema": {"type": "ARRAY", "items": {"type": "STRING"}},
        },
    }
    global _cur
    import time
    _resolve_endpoints()
    system, prompt = SYSTEM, build_prompt(items, kind, terms)

    last = None
    for rounds in range(MAX_ROUNDS):
        # chon muc dau tien chua bi phat (thu tu ENDPOINTS = thu tu uu tien,
        # Gemini truoc vi tieng Viet tot nhat)
        now = time.time()
        idx = next((i for i in range(len(ENDPOINTS)) if _cooldown[i] <= now), None)
        if idx is None:
            wait = min(max(min(_cooldown) - now, 1), 90)
            print("  [cho %ds - moi nha cung cap deu dang bi gioi han]" % int(wait))
            time.sleep(wait)
            continue
        if idx != _cur:
            e = ENDPOINTS[idx]
            print("  [chuyen -> %s/%s key#%d]"
                  % (e[0], e[1], _key_no(e)))
        _cur = idx
        prov, model = ENDPOINTS[idx][0], ENDPOINTS[idx][1]
        key = ENDPOINTS[idx][2] if len(ENDPOINTS[idx]) > 2 else None
        try:
            return _parse(CALLERS[prov](model, system, prompt, len(items), key),
                          len(items))
        except Exception as e:
            last = e
            msg = str(e)
            quota = "429" in msg or "RESOURCE_EXHAUSTED" in msg or "rate_limit" in msg
            busy = "503" in msg or "UNAVAILABLE" in msg or "timed out" in msg.lower()
            if quota or busy:
                # Phat muc nay mot luc roi quay lai - han muc free cua Gemini
                # tinh theo PHUT, cho 60s la dung tiep duoc, khong phai bo han
                # sang Groq (Groq hay treo im lang, da ket 50 phut hom 14/10).
                _cooldown[idx] = time.time() + (COOLDOWN if quota else COOLDOWN / 2)
                print("  [%s/%s %s -> nghi %ds]"
                      % (prov, model, "het quota" if quota else "qua tai",
                         int(_cooldown[idx] - time.time())))
                continue
            time.sleep(2 * min(rounds + 1, 4))
    raise last or RuntimeError("tat ca nha cung cap deu khong dung duoc")


if __name__ == "__main__":
    sys.stdout.reconfigure(encoding="utf-8")
    if len(sys.argv) > 1 and sys.argv[1] == "--models":
        _resolve_endpoints()
        for p, m in ENDPOINTS:
            print("  %-8s %s" % (p, m))
        sys.exit()
    args = sys.argv[1:] or ["Housing material", "Bus termination", "Drive concept"]
    for en, vi in zip(args, translate_batch(args, "key")):
        print("  %-34s -> %s" % (en, vi))
