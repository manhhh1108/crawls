# -*- coding: utf-8 -*-
"""Tu dien thuat ngu EN->VI cho SICK (dung chung duoc cho cac hang cam bien).

Gom 3 nguon:
  1. TERM_DICT trong sick/sick-translate.js (662 cap, tu du an SICK truoc)
  2. cot suggested_fix trong sick/loi dich san pham.xlsx (cac loi da soat tay)
  3. EXTRA ben duoi - bo sung tay, nham dung cac loi may dich hay mac
     (bus -> "xe buyt", housing -> "nha o", Art. -> "Nghe thuat", drive -> "lai xe"...)

API:  terms()    -> dict {en_lower: vi}
      phrases()  -> list (en_lower, vi) sap xep dai->ngan, de thay cum trong cau
"""
import io
import json
import os
import re

HERE = os.path.dirname(os.path.abspath(__file__))
REPO = os.path.dirname(HERE)
JS = os.path.join(REPO, "sick", "sick-translate.js")
FIXXLSX = os.path.join(REPO, "sick", "loi dich san pham.xlsx")
USERXLSX = os.path.join(REPO, "data", "Tu dien dich EN-VI.xlsx")
CACHE = os.path.join(HERE, "data", "glossary_sick.json")

# ---------------------------------------------------------------------------
# EXTRA - bo sung tay, nhom theo loai cho de soat.
# ---------------------------------------------------------------------------
EXTRA = {
    # ---- nhung tu may dich SAI NGHIA (uu tien cao nhat) ----
    "Bus": "Bus",
    "Bus termination": "Kết thúc bus (terminator)",
    "Bus address": "Địa chỉ bus",
    "Bus cycle time": "Thời gian chu kỳ bus",
    "Fieldbus": "Bus trường",
    "Bus connection": "Kết nối bus",
    "Housing": "Vỏ",
    "Housing side": "Phía vỏ",
    "Housing material": "Vật liệu vỏ",
    "Housing material detail": "Chi tiết vật liệu vỏ",
    "Housing design": "Kiểu vỏ",
    "Housing colour": "Màu vỏ",
    "Housing color": "Màu vỏ",
    "Housing paint finish": "Lớp sơn vỏ",
    "Drive": "Truyền động",
    "Drive concept": "Nguyên lý truyền động",
    "Shaft": "Trục",
    "Shaft diameter": "Đường kính trục",
    "Shaft material": "Vật liệu trục",
    "Shaft type": "Kiểu trục",
    "Blind hollow shaft": "Trục rỗng bịt đầu",
    "Through hollow shaft": "Trục rỗng suốt",
    "Solid shaft": "Trục đặc",
    "Face mount flange": "Mặt bích lắp mặt đầu",
    "Servo flange": "Mặt bích servo",
    "Flange": "Mặt bích",
    "Flange material": "Vật liệu mặt bích",
    "Art.": "Điều",
    "Art. No.": "Mã đặt hàng",
    "Part number": "Mã đặt hàng",
    "Order number": "Mã đặt hàng",
    "Type code": "Mã loại",
    "Spring": "Lò xo",
    "Spring-loaded": "Có lò xo",
    "Light source": "Nguồn sáng",
    "Light spot size": "Kích thước điểm sáng",
    "Shape of light spot": "Hình dạng điểm sáng",
    "Dark switching": "Chế độ tối (dark-on)",
    "Light switching": "Chế độ sáng (light-on)",
    "Light/dark switching": "Chế độ sáng/tối (light-on/dark-on)",
    "Window mode": "Chế độ cửa sổ",
    "Teach-in": "Teach-in (học mẫu)",
    "Pin assignment": "Sơ đồ chân",
    "Male connector": "Đầu nối đực",
    "Female connector": "Đầu nối cái",
    "Jacket color": "Màu vỏ cáp",
    "Jacket colour": "Màu vỏ cáp",
    "Jacket material": "Vật liệu vỏ cáp",
    "Core structure": "Cấu trúc lõi",
    "Core cross-section": "Tiết diện lõi",
    "Number of cores": "Số lõi",
    "Resolution": "Độ phân giải",
    "Shutter": "Màn trập",
    "Shutter technology": "Công nghệ màn trập",
    "Global shutter": "Màn trập toàn cục",
    "Rolling shutter": "Màn trập cuốn",
    "Frame rate": "Tốc độ khung hình",
    "Field of view": "Trường nhìn",

    # ---- ten thong so hay gap tren SICK ----
    "Sensor principle": "Nguyên lý cảm biến",
    "Functional principle": "Nguyên lý hoạt động",
    "Functional principle detail": "Chi tiết nguyên lý hoạt động",
    "Measurement principle": "Nguyên lý đo",
    "Sensing range": "Khoảng cách phát hiện",
    "Sensing range min.": "Khoảng cách phát hiện nhỏ nhất",
    "Sensing range max.": "Khoảng cách phát hiện lớn nhất",
    "Safe sensing range": "Khoảng cách phát hiện an toàn",
    "Operating range": "Phạm vi làm việc",
    "Scanning range": "Phạm vi quét",
    "Working range": "Phạm vi làm việc",
    "Detection range": "Phạm vi phát hiện",
    "Response time": "Thời gian đáp ứng",
    "Response sensitivity": "Độ nhạy đáp ứng",
    "Response sensitivity, typ.": "Độ nhạy đáp ứng, điển hình",
    "Hysteresis": "Độ trễ (hysteresis)",
    "Hysteresis, typ.": "Độ trễ (hysteresis), điển hình",
    "Repeatability": "Độ lặp lại",
    "Reproducibility": "Độ tái lập",
    "Linearity deviation": "Sai lệch tuyến tính",
    "Measurement accuracy": "Độ chính xác đo",
    "Switching frequency": "Tần số chuyển mạch",
    "Switching mode": "Chế độ chuyển mạch",
    "Switching output": "Ngõ ra chuyển mạch",
    "Output current": "Dòng điện ngõ ra",
    "Output signal": "Tín hiệu ngõ ra",
    "Analog output": "Ngõ ra analog",
    "Digital output": "Ngõ ra số",
    "Digital input": "Ngõ vào số",
    "Number of digital outputs": "Số ngõ ra số",
    "Type of digital output": "Kiểu ngõ ra số",
    "Supply voltage": "Điện áp cấp",
    "Power consumption": "Công suất tiêu thụ",
    "Current consumption": "Dòng tiêu thụ",
    "Enclosure rating": "Cấp bảo vệ vỏ",
    "Protection class": "Cấp cách điện",
    "Ambient operating temperature": "Nhiệt độ môi trường làm việc",
    "Ambient operation temp. max": "Nhiệt độ môi trường làm việc tối đa",
    "Ambient operation temp. min": "Nhiệt độ môi trường làm việc tối thiểu",
    "Storage temperature": "Nhiệt độ bảo quản",
    "Operating temperature range": "Dải nhiệt độ làm việc",
    "Storage temperature range": "Dải nhiệt độ bảo quản",
    "Relative humidity": "Độ ẩm tương đối",
    "Permissible relative humidity": "Độ ẩm tương đối cho phép",
    "Shock load": "Khả năng chịu sốc",
    "Vibration load": "Khả năng chịu rung",
    "Weight": "Khối lượng",
    "Dimensions": "Kích thước",
    "Mounting system type": "Kiểu hệ thống lắp đặt",
    "Mounting method": "Cách lắp đặt",
    "Mounting": "Lắp đặt",
    "Actuation option": "Tùy chọn tác động",
    "Actuating element": "Bộ phận tác động",
    "Code alignment": "Hướng đọc mã",
    "Fiber arrangement": "Cách bố trí sợi quang",
    "Fiber optic cable": "Cáp quang",
    "Magnetic sensitivity": "Độ nhạy từ",
    "Magnetic field strength": "Cường độ từ trường",
    "Connection type": "Kiểu kết nối",
    "Connection type Detail": "Chi tiết kiểu kết nối",
    "Cable material": "Vật liệu cáp",
    "Cable material detail": "Chi tiết vật liệu cáp",
    "Cable length": "Chiều dài cáp",
    "Communication interface": "Giao diện truyền thông",
    "Interface": "Giao diện",
    "Data transmission rate": "Tốc độ truyền dữ liệu",
    "Initialization time": "Thời gian khởi tạo",
    "Delay time": "Thời gian trễ",
    "Time functions": "Chức năng định thời",
    "Setting Method": "Cách cài đặt",
    "Special features": "Đặc điểm nổi bật",
    "Certificate": "Chứng nhận",
    "Certificates": "Chứng nhận",
    "Classifications": "Phân loại",
    "Safety-related parameters": "Thông số liên quan đến an toàn",
    "Performance": "Hiệu năng",
    "Electronics": "Phần điện",
    "Mechanics": "Phần cơ khí",
    "Ambient data": "Điều kiện môi trường",
    "Features": "Đặc điểm",
    "Reverse polarity protection": "Bảo vệ ngược cực",
    "Short-circuit protection": "Bảo vệ ngắn mạch",
    "Startup torque": "Mô-men khởi động",
    "Operating torque": "Mô-men vận hành",
    "Operating speed": "Tốc độ vận hành",
    "Angular acceleration": "Gia tốc góc",
    "Bearing lifetime": "Tuổi thọ vòng bi",
    "Steps per revolution": "Số xung trên vòng",
    "Number of revolutions": "Số vòng quay",
    "Permissible movement": "Dịch chuyển cho phép",
    "Permissible static movement": "Dịch chuyển tĩnh cho phép",
    "Electromagnetic compatibility": "Tương thích điện từ (EMC)",
    "Electromagnetic compatibility (EMC)": "Tương thích điện từ (EMC)",
    "Scanning rate": "Tốc độ quét",
    "Aperture angle": "Góc mở",
    "Protective field range": "Tầm vùng bảo vệ",
    "Number of beams": "Số tia",
    "Beam separation": "Khoảng cách giữa các tia",
    "Resolution for hand detection": "Độ phân giải phát hiện bàn tay",
    "Safety integrity level": "Mức toàn vẹn an toàn (SIL)",
    "Performance level": "Mức hiệu năng (PL)",
    "Smallest detectable object": "Vật thể nhỏ nhất phát hiện được",
    "Smallest detectable object (MDO) typ.": "Vật thể nhỏ nhất phát hiện được (MDO), điển hình",
    "Object distance (max.)": "Khoảng cách vật thể (tối đa)",
    "Focus position": "Vị trí tiêu cự",
    "Bestseller": "Sản phẩm bán chạy",
    "Product segment": "Phân khúc sản phẩm",
    "Items included": "Phụ kiện kèm theo",
    "Fields of application": "Lĩnh vực ứng dụng",
    "Advantages": "Ưu điểm",
    "Benefits": "Lợi ích",
    "At a glance": "Tổng quan nhanh",

    # ---- gia tri hay gap ----
    "Heavy-Duty": "Môi trường khắc nghiệt",
    "Rubber-mounted": "Có lắp đệm cao su",
    "Location-dependent": "Phụ thuộc vị trí",
    "Application-specific": "Tùy theo ứng dụng cụ thể",
    "Halogen-free": "Không chứa halogen",
    "Flame-retardant": "Chống cháy lan",
    "Silicone-free": "Không chứa silicone",
    "Maintenance-free": "Không cần bảo trì",
    "Oil-resistant": "Chịu dầu",
    "Drag-chain capable": "Dùng được trong xích cáp",
    "Shielded": "Có vỏ chống nhiễu",
    "Unshielded": "Không có vỏ chống nhiễu",
    "Flush": "Lắp chìm",
    "Non-flush": "Lắp nổi",
    "Stainless steel": "Thép không gỉ",
    "Aluminum": "Nhôm",
    "Aluminium": "Nhôm",
    "Brass": "Đồng thau",
    "Zinc die-cast": "Kẽm đúc áp lực",
    "Plastic": "Nhựa",
    "Steel": "Thép",
    "Cable": "Cáp",
    "Cable with male connector": "Cáp có đầu nối đực",
    "Yes": "Có",
    "No": "Không",
    "Optional": "Tùy chọn",
    "On request": "Theo yêu cầu",
    "Factory setting": "Cài đặt mặc định",
    "deactivated": "tắt",
    "activated": "bật",
    "Background suppression": "Triệt nền (BGS)",
    "Foreground suppression": "Triệt tiền cảnh (FGS)",
    "Through-beam": "Thu phát độc lập",
    "Retro-reflective": "Phản xạ gương",
    "Diffuse": "Khuếch tán",
    "at room temperature": "ở nhiệt độ phòng",
    "radial": "hướng kính",
    "axial": "hướng trục",

    # ---- bo sung sau khi soat lot dich dau tien ----
    "TFT display": "Màn hình TFT",
    "LCD display": "Màn hình LCD",
    "OLED display": "Màn hình OLED",
    "Display": "Màn hình",
    "socket": "ổ cắm",
    "RJ45 socket": "cổng RJ45",
    "spring terminals": "cầu đấu lò xo",
    "screw terminals": "cầu đấu vít",
    "Front connector with spring terminals": "Đầu nối mặt trước, cầu đấu lò xo",
    "safety gateway": "gateway an toàn",
    "Gateway": "Gateway",
    "main module": "mô-đun chính",
    "Relay module": "Mô-đun rơ-le",
    "I/O module": "Mô-đun I/O",
    "Inputs/outputs": "Ngõ vào/ngõ ra",
    "safety inputs": "ngõ vào an toàn",
    "safety outputs": "ngõ ra an toàn",
    "test outputs": "ngõ ra kiểm tra",
    "safety capable inputs": "ngõ vào hỗ trợ an toàn",
    "output expansion module for OSSDs": "mô-đun mở rộng ngõ ra OSSD",
    "modular safety controller": "bộ điều khiển an toàn dạng mô-đun",
    "Application": "Ứng dụng",
    "Applications": "Ứng dụng",
    "Product": "Sản phẩm",
    "Products": "Sản phẩm",
    "New products": "Sản phẩm mới",
    "Configuration method": "Cách cấu hình",
    "Configuration software": "Phần mềm cấu hình",
    "Configuration and diagnostics interface": "Giao diện cấu hình và chẩn đoán",
    "Fieldbus, industrial network": "Fieldbus, mạng công nghiệp",
    "Third party certificate": "Chứng nhận bên thứ ba",
    "EAC certificate / DoC": "Chứng nhận EAC / DoC",
    "China RoHS": "RoHS Trung Quốc",
    "UK declaration of conformity": "Tuyên bố hợp chuẩn UK",
    "EU declaration of conformity": "Tuyên bố hợp chuẩn EU",
    "via software": "qua phần mềm",
    "Module": "Mô-đun",
    "Modules": "Mô-đun",
    "Number of modules": "Số mô-đun",
    "Cycle time": "Thời gian chu kỳ",
    "Function blocks": "Khối chức năng",
    "Logic": "Logic",
    "stationary": "cố định",
    "mobile applications": "ứng dụng di động",
    "micro USB": "micro USB",
    "system plugs": "phích cắm hệ thống",
    "system plug": "phích cắm hệ thống",
    "Bar Light": "Đèn thanh (bar light)",
    "Ring Light": "Đèn vòng (ring light)",
    "Dome Light": "Đèn vòm (dome light)",
    "External illumination": "Chiếu sáng ngoài",
    "illumination color": "màu ánh sáng",
    "Illuminations": "Đèn chiếu sáng",
    "focal length": "tiêu cự",
    "focal lengths": "tiêu cự",
    "lens": "ống kính",
    "lenses": "ống kính",
    "part no.": "mã SP",
    "megapixel": "megapixel",
    "Quick and easy mounting": "Lắp đặt nhanh và dễ dàng",
    "mini USB": "mini USB",
    "USB-C": "USB-C",
    # "track and trace" tung bi may dich thanh "theo doi va theo doi"
    "track and trace": "theo dõi và truy xuất",
    "track and trace system": "hệ thống theo dõi và truy xuất",
    "track and trace systems": "hệ thống theo dõi và truy xuất",
    "Track and Trace": "theo dõi và truy xuất",
    "traceability": "truy xuất nguồn gốc",
}


def _from_js():
    try:
        t = io.open(JS, encoding="utf-8").read()
    except Exception:
        return {}
    i = t.find("const TERM_DICT = {")
    j = t.find("\n};", i)
    out = {}
    for k, v in re.findall(r'^\s*"?([^":\n]+?)"?\s*:\s*"([^"]*)"\s*,?\s*$', t[i:j], re.M):
        k, v = k.strip().strip('"'), v.strip()
        if k and v:
            out[k] = v
    return out


def _from_user_xlsx():
    """data/Tu dien dich EN-VI.xlsx - file nguoi dung tu dien, uu tien CAO NHAT.

    Cot A = tieng Anh, cot B = tieng Viet. Bo qua dong trong / dong tieu de.
    """
    try:
        from openpyxl import load_workbook
        ws = load_workbook(USERXLSX, read_only=True, data_only=True).active
    except Exception:
        return {}
    out = {}
    for r in ws.iter_rows(values_only=True):
        if not r or len(r) < 2:
            continue
        en = str(r[0] or "").strip()
        vi = str(r[1] or "").strip()
        if not en or not vi or en == "Tiếng Anh":
            continue
        out[en] = vi
    return out


def _from_fixes():
    try:
        from openpyxl import load_workbook
        ws = load_workbook(FIXXLSX, read_only=True).active
    except Exception:
        return {}
    out = {}
    for r in ws.iter_rows(values_only=True):
        if not r or len(r) < 8 or r[0] == "family":
            continue
        en, fix = (r[6] or "").strip(), (r[7] or "").strip()
        # bo cac dong chi doi dau thap phan (da co rule rieng xu ly)
        if en and fix and not re.match(r"^[<>=~\s\d.,]+(mT|mm|cm|m|kg|g)?$", en):
            out[en] = fix
    return out


_TERMS = None


def terms():
    global _TERMS
    if _TERMS is None:
        merged = {}
        # nguon sau de len nguon truoc -> file nguoi dung tu dien uu tien cao nhat
        for src in (_from_js(), _from_fixes(), EXTRA, _from_user_xlsx()):
            for k, v in src.items():
                merged[k.strip().lower()] = v
        _TERMS = merged
        try:
            os.makedirs(os.path.dirname(CACHE), exist_ok=True)
            with io.open(CACHE, "w", encoding="utf-8") as f:
                json.dump(merged, f, ensure_ascii=False, indent=0, sort_keys=True)
        except Exception:
            pass
    return _TERMS


_PHRASES = None


def phrases(minlen=7):
    """Cum tu de thay truc tiep trong cau (dai -> ngan de khong cat nham)."""
    global _PHRASES
    if _PHRASES is None:
        ps = [(k, v) for k, v in terms().items() if len(k) >= minlen and " " in k]
        ps.sort(key=lambda kv: -len(kv[0]))
        _PHRASES = ps
    return _PHRASES


if __name__ == "__main__":
    import sys
    sys.stdout.reconfigure(encoding="utf-8")
    t = terms()
    print("js:", len(_from_js()), "| fixes:", len(_from_fixes()),
          "| extra:", len(EXTRA), "| file nguoi dung:", len(_from_user_xlsx()))
    print("tong:", len(t), "->", CACHE)
