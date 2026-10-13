# -*- coding: utf-8 -*-
"""Ten dong SP SICK dich sang tieng Viet (chi dung o COT B - ban dang web).

Nguyen tac:
  DICH   - ten mo ta chung loai hang hoa (phu kien, cap, nguon, thiet bi do...)
           khach Viet tra cuu bang tieng Viet.
  GIU    - ten he thong / dong SP rieng cua SICK (Lector, Flexi, CLV, Safe
           EFI-pro, ... System). Doi ten se mat tu khoa tra cuu chinh xac, va
           Meta title da co phan loai bang tieng Viet dung truoc roi.

Pham vi anh huong: The tag, Keyword chinh, Meta title, Meta description, mo ta
(cot B). KHONG doi URL/slug va KHONG doi cot C (tieng Anh) - de giu SEO cu.

Sua xong chi can chay lai `python build_sick_excel.py` (khong phai crawl lai).
"""

NAME_VI = {
    # ---- Phu kien ----
    "Actuators for electromechanical safety switches":
        "Cơ cấu tác động cho công tắc an toàn cơ điện",
    "Actuators for magnetic safety switches":
        "Cơ cấu tác động cho công tắc an toàn từ tính",
    "Actuators for RFID safety switches":
        "Cơ cấu tác động cho công tắc an toàn RFID",
    "RFID antennas": "Anten RFID",
    "WLAN antennas": "Anten WLAN",
    "GPS antennas": "Anten GPS",
    "GSM, UMTS and LTE antennas": "Anten GSM, UMTS và LTE",
    "Front screens and front screen protection": "Kính chắn trước và tấm bảo vệ kính chắn",
    "Protective caps and hoods": "Nắp và chụp bảo vệ",
    "Purge air equipment": "Thiết bị thổi khí làm sạch",
    "Interface modules": "Module giao tiếp",
    "Power supply modules": "Module cấp nguồn",
    "SSI analog adapters": "Bộ chuyển đổi SSI sang analog",
    "Measuring wheel mechanics": "Cơ cấu bánh xe đo",
    "Flanges and nozzles": "Mặt bích và đầu phun",
    "Acoustic signal transmitters": "Bộ phát tín hiệu âm thanh",
    "Optical signal transmitters": "Bộ phát tín hiệu quang",
    "Wire draw mechanism for rope pull switches": "Cơ cấu kéo dây cho công tắc giật dây",
    "Wire draw mechanism for wire draw encoders": "Cơ cấu kéo dây cho encoder kéo dây",

    # ---- Thiet bi phan tich ----
    "Air quality measuring devices": "Thiết bị đo chất lượng không khí",
    "Digital analyzer solutions": "Giải pháp phân tích số",
    "Dust Measuring Devices": "Thiết bị đo bụi",
    "Visual range measuring devices": "Thiết bị đo tầm nhìn",

    # ---- Cam bien ----
    "Roller Sensor Bar": "Thanh cảm biến con lăn",

    # ---- Do luu luong ----
    "Digital flow solutions": "Giải pháp đo lưu lượng số",
    "Flow measurement instruments": "Thiết bị đo lưu lượng",
    "Flow metering systems": "Hệ thống đo đếm lưu lượng",

    # ---- Cap va ket noi ----
    "Industrial Ethernet cables and fieldbus cables":
        "Cáp Ethernet công nghiệp và cáp bus trường",
    "Other connectors and cables": "Đầu nối và cáp khác",
    "Passive distribution boxes": "Hộp phân phối thụ động",
    "Y- and T-splitters": "Bộ chia Y và T",
    "Power supply cables": "Cáp cấp nguồn",
    "Power supply units": "Bộ nguồn",
    "System plugs and front connectors Flexi":
        "Phích cắm hệ thống và đầu nối mặt trước Flexi",
    "System plugs Flexi Soft": "Phích cắm hệ thống Flexi Soft",
}

# Ten GIU NGUYEN tieng Anh + ban dich goi y, phong khi muon doi y sau nay.
# Chuyen 1 dong tu day len NAME_VI la ap dung duoc ngay.
SUGGESTED = {
    "CSS High Resolution": "CSS độ phân giải cao",
    "CSS/CSX High Speed": "CSS/CSX tốc độ cao",
    "Telematic Data Collector": "Bộ thu thập dữ liệu telematic",
    "Flexi Soft Drive Monitor": "Bộ giám sát truyền động Flexi Soft",
    "Speed Monitor": "Bộ giám sát tốc độ",
    "AGV Dynamic Weather Assist": "Hỗ trợ AGV theo điều kiện thời tiết",
    "Safe Brake Assist": "Hệ thống hỗ trợ phanh an toàn",
    "Safe EFI-pro System": "Hệ thống an toàn EFI-pro",
    "Safe Entry Exit": "Hệ thống kiểm soát ra vào an toàn",
    "Safe Robotics Area Protection": "Bảo vệ vùng làm việc robot an toàn",
    "Automated Load Detect Ident System": "Hệ thống nhận dạng và phát hiện tải tự động",
    "Area Hotspot Detection System": "Hệ thống phát hiện điểm nóng theo vùng",
    "Backup Assistance System": "Hệ thống hỗ trợ lùi xe",
    "Drive Assist System": "Hệ thống hỗ trợ lái",
    "Overheight Detection System": "Hệ thống phát hiện quá chiều cao",
    "Vehicle Hotspot Detection System": "Hệ thống phát hiện điểm nóng trên xe",
    "Axle Classification System": "Hệ thống phân loại trục xe",
    "Free Flow Profiler": "Thiết bị đo biên dạng dòng chảy tự do",
    "Load Volume Measurement System": "Hệ thống đo thể tích hàng hóa",
    "Multi-Lane Profiling System": "Hệ thống đo biên dạng đa làn",
    "Foreign Object Detection System": "Hệ thống phát hiện vật thể lạ",
    "Pallet Classification System": "Hệ thống phân loại pallet",
    "Pallet Integrity Inspection System": "Hệ thống kiểm tra nguyên vẹn pallet",
    "Pallet Serialization System": "Hệ thống gán số sê-ri pallet",
    "Body Position System": "Hệ thống xác định vị trí thân xe",
    "CLV Identification System": "Hệ thống nhận dạng CLV",
    "ICR Identification System": "Hệ thống nhận dạng ICR",
    "Ident Gate System": "Hệ thống cổng nhận dạng",
    "Inspector Logistics System": "Hệ thống Inspector cho logistics",
    "Lector Identification System": "Hệ thống nhận dạng Lector",
    "Master Data Analyzer": "Thiết bị phân tích dữ liệu gốc",
    "Master Data Analyzer Vision": "Thiết bị phân tích dữ liệu gốc bằng hình ảnh",
    "RF Identification System": "Hệ thống nhận dạng bằng sóng vô tuyến",
    "VMV Dimensioning System": "Hệ thống đo kích thước VMV",
}


def name_vi(name):
    return NAME_VI.get((name or "").strip(), name)
