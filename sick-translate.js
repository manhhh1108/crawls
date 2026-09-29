const { readFileSync, writeFileSync, existsSync } = require("node:fs");
const { resolve } = require("node:path");

const TERM_DICT = {
  // --- Nguyen ly & loai cam bien ---
  "Functional principle": "Nguyên lý hoạt động",
  "Functional principle detail": "Chi tiết nguyên lý hoạt động",
  "Photoelectric sensors": "Cảm biến quang",
  "Photoelectric retro-reflective sensor": "Cảm biến quang phản xạ gương",
  "Photoelectric proximity sensor": "Cảm biến quang tiệm cận",
  "Through-beam photoelectric sensor": "Cảm biến quang thu phát độc lập",
  "Diffuse photoelectric sensor": "Cảm biến quang khuếch tán",
  "Background suppression": "Triệt tiêu nền",
  "Foreground suppression": "Triệt tiêu tiền cảnh",
  "Energetic diffuse": "Khuếch tán năng lượng",
  "Inductive proximity sensor": "Cảm biến tiệm cận cảm ứng",
  "Inductive sensor": "Cảm biến cảm ứng",
  "Capacitive proximity sensor": "Cảm biến tiệm cận điện dung",
  "Capacitive sensor": "Cảm biến điện dung",
  "Magnetic sensor": "Cảm biến từ",
  "Magnetic cylinder sensor": "Cảm biến xi lanh từ",
  "Ultrasonic sensor": "Cảm biến siêu âm",
  "Distance sensor": "Cảm biến khoảng cách",
  "Displacement sensor": "Cảm biến dịch chuyển",
  "Level sensor": "Cảm biến mức",
  "Flow sensor": "Cảm biến lưu lượng",
  "Pressure sensor": "Cảm biến áp suất",
  "Temperature sensor": "Cảm biến nhiệt độ",
  "Tilt sensor": "Cảm biến nghiêng",
  "Inclination sensor": "Cảm biến độ nghiêng",
  "Acceleration sensor": "Cảm biến gia tốc",
  "Shock sensor": "Cảm biến va chạm",
  "Vibration sensor": "Cảm biến rung",
  "Color sensor": "Cảm biến màu sắc",
  "Contrast sensor": "Cảm biến tương phản",
  "Luminescence sensor": "Cảm biến phát quang",
  "Gloss sensor": "Cảm biến độ bóng",
  "Fork sensor": "Cảm biến chạc",
  "Slot sensor": "Cảm biến khe",
  "Array sensor": "Cảm biến mảng",
  "Fiber-optic sensor": "Cảm biến sợi quang",
  "Fiber-optic sensors": "Cảm biến sợi quang",
  "Fiber-optic amplifier": "Bộ khuếch đại sợi quang",
  "Fiber-optic cables": "Cáp sợi quang",
  "Fiber optic cable": "Cáp sợi quang",
  "Light curtain": "Rèm quang",
  "Light grid": "Lưới quang",
  "Safety light curtain": "Rèm quang an toàn",
  "Safety light curtains": "Rèm quang an toàn",
  "Safety laser scanner": "Máy quét laser an toàn",
  "Safety laser scanners": "Máy quét laser an toàn",
  "Laser scanner": "Máy quét laser",
  "Bar code scanner": "Máy quét mã vạch",
  "Barcode scanner": "Máy quét mã vạch",
  "2D code reader": "Đầu đọc mã 2D",
  "1D/2D code reader": "Đầu đọc mã 1D/2D",
  "QR code reader": "Đầu đọc mã QR",
  "RFID reader": "Đầu đọc RFID",
  "Vision sensor": "Cảm biến thị giác",
  "Camera": "Camera",
  "3D camera": "Camera 3D",
  "Smart camera": "Camera thông minh",
  "Smart sensor": "Cảm biến thông minh",
  "Encoder": "Bộ mã hóa",
  "Incremental encoder": "Bộ mã hóa gia tăng",
  "Absolute encoder": "Bộ mã hóa tuyệt đối",
  "Linear encoder": "Bộ mã hóa tuyến tính",
  "Multiturn encoder": "Bộ mã hóa nhiều vòng",
  "Singleturn encoder": "Bộ mã hóa một vòng",
  "Motor feedback system": "Hệ thống phản hồi động cơ",
  "Safety controller": "Bộ điều khiển an toàn",
  "Safety controllers": "Bộ điều khiển an toàn",
  "Safety relay": "Rơ-le an toàn",
  "Safety relays": "Rơ-le an toàn",
  "Safety switch": "Công tắc an toàn",
  "Safety switches": "Công tắc an toàn",
  "Safety gate switch": "Công tắc cửa an toàn",
  "Emergency stop": "Dừng khẩn cấp",
  "Enabling switch": "Công tắc cho phép",
  "Rope pull switch": "Công tắc kéo dây",
  "Foot switch": "Công tắc chân",
  "Limit switch": "Công tắc giới hạn",
  "Magnetic safety switch": "Công tắc an toàn từ",
  "Interlock": "Khóa liên động",
  "Guard locking": "Khóa bảo vệ",
  "Transponder": "Bộ thu phát (Transponder)",
  "Solenoid interlock": "Khóa liên động điện từ",

  // --- Mo ta chung ---
  "Description": "Mô tả",
  "Description detail": "Chi tiết mô tả",
  "Brief description": "Mô tả ngắn",
  "Short description": "Mô tả ngắn",
  "Long description": "Mô tả chi tiết",
  "Items supplied": "Các mặt hàng được cung cấp",
  "Specialty": "Tính năng đặc biệt",
  "Sales Kit": "Sales Kit",
  "Sales-Kit": "Sales Kit",

  // --- Hinh dang & ki hieu ---
  "Point-shaped": "Hình điểm",
  "Line-shaped": "Hình đường",
  "Spot-shaped": "Hình điểm tròn",
  "Rectangular-shaped": "Hình chữ nhật",
  "Round-shaped": "Hình tròn",

  // --- ID & Communication (DEC = decimal, HEX = hexadecimal, KHONG dich) ---
  "Vendor ID": "ID nhà cung cấp",
  "Device ID HEX": "ID thiết bị HEX",
  "Device ID DEC": "ID thiết bị DEC",
  "DeviceID DEC": "ID thiết bị DEC",
  "DeviceID HEX": "ID thiết bị HEX",
  "Vendor ID HEX": "ID nhà cung cấp HEX",
  "Vendor ID DEC": "ID nhà cung cấp DEC",
  "Compatible master port type": "Loại cổng chính tương thích",
  "Process data length": "Độ dài dữ liệu xử lý",
  "Process data structure": "Cấu trúc dữ liệu xử lý",
  "Min. cycle time": "Thời gian chu kỳ tối thiểu",
  "Cycle time": "Thời gian chu kỳ",

  // --- Loai thiet bi ---
  "Device type": "Loại thiết bị",
  "Device type detail": "Chi tiết loại thiết bị",
  "Stand-alone": "Độc lập",
  "Master": "Thiết bị chủ",
  "Slave": "Thiết bị phụ",
  "Transmitter": "Bộ phát",
  "Receiver": "Bộ thu",
  "Sender": "Bộ phát",
  "Transceiver": "Bộ thu phát",
  "Controller": "Bộ điều khiển",
  "Amplifier": "Bộ khuếch đại",
  "Evaluation unit": "Đơn vị xử lý",
  "Processing unit": "Đơn vị xử lý",
  "Gateway": "Cổng kết nối",
  "Repeater": "Bộ lặp",
  "Junction box": "Hộp đấu nối",
  "Splitter": "Bộ chia",

  // --- Pham vi do ---
  "Sensing range": "Phạm vi cảm biến",
  "Sensing range min.": "Phạm vi cảm biến tối thiểu",
  "Sensing range max.": "Phạm vi cảm biến tối đa",
  "Nominal sensing range": "Phạm vi cảm biến danh nghĩa",
  "Assured sensing range": "Phạm vi cảm biến đảm bảo",
  "Operating range": "Khoảng cách vận hành",
  "Operating range min.": "Khoảng cách vận hành tối thiểu",
  "Operating range max.": "Khoảng cách vận hành tối đa",
  "Detection range": "Phạm vi phát hiện",
  "Measuring range": "Dải đo",
  "Measuring range min.": "Dải đo tối thiểu",
  "Measuring range max.": "Dải đo tối đa",
  "Scanning range": "Phạm vi quét",
  "Scanning distance": "Khoảng cách quét",
  "Working range": "Phạm vi làm việc",
  "Object distance": "Khoảng cách vật thể",
  "Typical detectable object": "Vật thể phát hiện điển hình",
  "Min. detectable object": "Vật thể phát hiện tối thiểu",
  "Reduction factor": "Hệ số giảm",
  "Correction factor": "Hệ số hiệu chỉnh",
  "Blind zone": "Vùng mù",
  "Dead zone": "Vùng chết",
  "Near range": "Khoảng cách gần",
  "Far range": "Khoảng cách xa",
  "Field of view": "Trường nhìn",
  "Beam angle": "Góc chùm tia",
  "Opening angle": "Góc mở",
  "Aperture angle": "Góc khẩu độ",
  "Scanning angle": "Góc quét",
  "Angular resolution": "Độ phân giải góc",
  "Resolution": "Độ phân giải",
  "Accuracy": "Độ chính xác",
  "Repeatability": "Độ lặp lại",
  "Reproducibility": "Khả năng tái hiện",
  "Linearity": "Độ tuyến tính",
  "Hysteresis": "Độ trễ",
  "Drift": "Độ trôi",
  "Temperature drift": "Độ trôi nhiệt",
  "Span": "Khoảng đo",
  "Of span": "Toàn dải đo",
  "Of the span": "Toàn dải đo",
  "Scale division": "Vạch chia thang đo",

  // --- Nguon sang ---
  "Light source": "Nguồn sáng",
  "LED": "LED",
  "Laser": "Laser",
  "Laser class": "Cấp độ laser",
  "Laser power": "Công suất laser",
  "Infrared": "Hồng ngoại",
  "Infrared light": "Ánh sáng hồng ngoại",
  "Red light": "Ánh sáng đỏ",
  "Visible red light": "Ánh sáng đỏ khả kiến",
  "Visible light": "Ánh sáng khả kiến",
  "Blue light": "Ánh sáng xanh",
  "White light": "Ánh sáng trắng",
  "UV light": "Ánh sáng UV",
  "PinPoint LED": "LED PinPoint",
  "Wavelength": "Bước sóng",
  "Wavelength min.": "Bước sóng tối thiểu",
  "Wavelength max.": "Bước sóng tối đa",
  "Type of light": "Loại ánh sáng",
  "Light spot size": "Kích thước điểm sáng",
  "Spot size": "Kích thước điểm sáng",
  "Beam diameter": "Đường kính chùm tia",
  "Reflector": "Gương phản xạ",
  "Polarisation filter": "Bộ lọc phân cực",
  "Optical filter": "Bộ lọc quang",
  "Pulse modulation": "Điều chế xung",
  "Modulated light": "Ánh sáng điều chế",

  // --- Dien ---
  "Supply voltage": "Điện áp cấp",
  "Supply voltage V": "Điện áp cấp (V)",
  "Supply voltage min.": "Điện áp cấp tối thiểu",
  "Supply voltage max.": "Điện áp cấp tối đa",
  "Voltage": "Điện áp",
  "Voltage type": "Loại điện áp",
  "DC": "DC (Một chiều)",
  "AC": "AC (Xoay chiều)",
  "AC/DC": "AC/DC",
  "Current consumption": "Dòng tiêu thụ",
  "No-load current consumption": "Dòng tiêu thụ không tải",
  "Max. current consumption": "Dòng tiêu thụ tối đa",
  "Power consumption": "Công suất tiêu thụ",
  "Residual ripple": "Độ gợn sóng dư",
  "Ripple": "Độ gợn sóng",
  "Fuse": "Cầu chì",
  "Max. fuse rating": "Cầu chì tối đa",
  "Insulation resistance": "Điện trở cách điện",
  "Dielectric strength": "Độ bền điện môi",
  "EMC": "Tương thích điện từ (EMC)",
  "EMI": "Nhiễu điện từ (EMI)",
  "Electromagnetic compatibility": "Tương thích điện từ",
  "Ground": "Nối đất",
  "Galvanic isolation": "Cách ly điện",
  "Leakage current": "Dòng rò",
  "Voltage drop": "Sụt áp",

  // --- Dau ra ---
  "Switching output": "Đầu ra đóng ngắt",
  "Switching output detail": "Chi tiết đầu ra đóng ngắt",
  "Digital output": "Đầu ra kỹ thuật số",
  "Analog output": "Đầu ra analog",
  "Output type": "Loại đầu ra",
  "Output function": "Chức năng đầu ra",
  "Output current Imax.": "Dòng ra tối đa Imax.",
  "Output current": "Dòng ra",
  "Output voltage": "Điện áp ra",
  "Signal output": "Tín hiệu đầu ra",
  "Current output": "Đầu ra dòng điện",
  "Voltage output": "Đầu ra điện áp",
  "4...20 mA": "4...20 mA",
  "0...10 V": "0...10 V",
  "0...20 mA": "0...20 mA",
  "Relay output": "Đầu ra rơ-le",
  "Transistor output": "Đầu ra transistor",
  "Dual output": "Đầu ra kép",
  "Number of outputs": "Số lượng đầu ra",
  "Number of inputs": "Số lượng đầu vào",
  "Input": "Đầu vào",
  "Output": "Đầu ra",
  "Residual current": "Dòng dư",
  "Voltage at output": "Điện áp tại đầu ra",
  "Short-circuit current": "Dòng ngắn mạch",
  "Max. load": "Tải tối đa",
  "Load capacity": "Khả năng tải",
  "OSSD": "OSSD (Đầu ra an toàn)",
  "OSSD output": "Đầu ra OSSD",
  "Test pulse": "Xung kiểm tra",

  // --- Tan so & thoi gian ---
  "Switching frequency": "Tần số đóng ngắt",
  "Response time": "Thời gian phản hồi",
  "Switch-on delay": "Trễ bật",
  "Switch-off delay": "Trễ tắt",
  "Delay time": "Thời gian trễ",
  "Initialization time": "Thời gian khởi tạo",
  "Start-up time": "Thời gian khởi động",
  "Warm-up time": "Thời gian làm nóng",
  "Pulse/Pause ratio": "Tỷ lệ xung/nghỉ",
  "Sampling rate": "Tốc độ lấy mẫu",
  "Scan rate": "Tốc độ quét",
  "Cycle time": "Thời gian chu kỳ",
  "Update rate": "Tần suất cập nhật",
  "Baud rate": "Tốc độ baud",
  "Data rate": "Tốc độ dữ liệu",
  "Integration time": "Thời gian tích phân",
  "Exposure time": "Thời gian phơi sáng",
  "Processing time": "Thời gian xử lý",
  "Reaction time": "Thời gian phản ứng",
  "Safe state response time": "Thời gian phản hồi trạng thái an toàn",

  // --- Che do ---
  "Light/Dark mode": "Chế độ sáng/tối",
  "Light switching": "Chuyển đổi sáng",
  "Dark switching": "Chuyển đổi tối",
  "Switching mode": "Chế độ đóng ngắt",
  "Operating mode": "Chế độ vận hành",
  "Window mode": "Chế độ cửa sổ",
  "Differential mode": "Chế độ vi sai",
  "Single-beam": "Chùm tia đơn",
  "Multi-beam": "Chùm tia đa",
  "Auto-power control": "Điều khiển công suất tự động",
  "Auto-sensitivity": "Độ nhạy tự động",
  "Manual sensitivity": "Độ nhạy thủ công",
  "Sensitivity": "Độ nhạy",
  "Threshold": "Ngưỡng",
  "Trigger mode": "Chế độ kích hoạt",
  "Free-running mode": "Chế độ chạy tự do",
  "Continuous mode": "Chế độ liên tục",
  "Single shot mode": "Chế độ đơn lần",
  "Power mode": "Chế độ công suất",
  "Speed mode": "Chế độ tốc độ",
  "Precision mode": "Chế độ chính xác",

  // --- Bao ve ---
  "Enclosure rating": "Cấp bảo vệ vỏ",
  "Protection class": "Cấp bảo vệ điện",
  "IP rating": "Cấp bảo vệ IP",
  "IP67": "IP67",
  "IP68": "IP68",
  "IP69K": "IP69K",
  "NEMA": "NEMA",
  "Short-circuit protected": "Bảo vệ ngắn mạch",
  "Reverse polarity protected": "Bảo vệ đảo cực",
  "Overload protection": "Bảo vệ quá tải",
  "Overvoltage protection": "Bảo vệ quá áp",
  "Undervoltage protection": "Bảo vệ thiếu áp",
  "Polarity reversal protection": "Bảo vệ đảo chiều cực",
  "Surge protection": "Bảo vệ xung điện",
  "ESD protection": "Bảo vệ tĩnh điện (ESD)",
  "Shock resistance": "Khả năng chịu va đập",
  "Vibration resistance": "Khả năng chịu rung",
  "Corrosion resistance": "Khả năng chịu ăn mòn",
  "Chemical resistance": "Khả năng chịu hóa chất",
  "Sunlight resistance": "Khả năng chịu ánh nắng",
  "Ambient light immunity": "Khả năng chịu ánh sáng môi trường",
  "Interference immunity": "Khả năng chịu nhiễu",
  "Mutual interference": "Nhiễu lẫn nhau",

  // --- Nhiet do ---
  "Ambient operating temperature": "Nhiệt độ môi trường vận hành",
  "Ambient storage temperature": "Nhiệt độ môi trường lưu trữ",
  "Operating temperature": "Nhiệt độ vận hành",
  "Storage temperature": "Nhiệt độ lưu trữ",
  "Ambient operation temp. max": "Nhiệt độ vận hành tối đa",
  "Ambient operation temp. min": "Nhiệt độ vận hành tối thiểu",
  "Temperature range": "Dải nhiệt độ",
  "Max. operating temperature": "Nhiệt độ vận hành tối đa",
  "Min. operating temperature": "Nhiệt độ vận hành tối thiểu",
  "Temperature coefficient": "Hệ số nhiệt độ",
  "Humidity": "Độ ẩm",
  "Relative humidity": "Độ ẩm tương đối",
  "Max. relative humidity": "Độ ẩm tương đối tối đa",
  "Condensation": "Ngưng tụ",
  "Non-condensing": "Không ngưng tụ",
  "Altitude": "Độ cao",

  // --- Vat lieu & co khi ---
  "Housing material": "Vật liệu vỏ",
  "Housing design": "Thiết kế vỏ",
  "Housing design (light emission)": "Thiết kế vỏ (phát xạ ánh sáng)",
  "Housing material detail": "Chi tiết vật liệu vỏ",
  "Front screen material": "Vật liệu mặt kính",
  "Lens material": "Vật liệu thấu kính",
  "Housing": "Vỏ",
  "Enclosure": "Vỏ bọc",
  "Protection hood": "Nắp bảo vệ",
  "Typ. Ambient light immunity": "Khả năng chống nhiễu ánh sáng môi trường điển hình",
  "Weight": "Trọng lượng",
  "Dimensions": "Kích thước",
  "Width": "Chiều rộng",
  "Height": "Chiều cao",
  "Depth": "Chiều sâu",
  "Length": "Chiều dài",
  "Diameter": "Đường kính",
  "Thread": "Ren",
  "Thread size": "Kích thước ren",
  "Mounting thread": "Ren lắp đặt",
  "Material": "Vật liệu",
  "Plastic": "Nhựa",
  "ABS": "Nhựa ABS",
  "PC": "Nhựa Polycarbonate",
  "PBT": "Nhựa PBT",
  "PA": "Nhựa Polyamide",
  "PVC": "PVC",
  "TPU": "TPU",
  "Metal": "Kim loại",
  "Stainless steel": "Thép không gỉ",
  "Aluminum": "Nhôm",
  "Brass": "Đồng thau",
  "Die-cast zinc": "Kẽm đúc",
  "PMMA": "PMMA",
  "Glass": "Kính",
  "Sapphire glass": "Kính Sapphire",
  "Design detail": "Chi tiết thiết kế",
  "Cylindrical": "Hình trụ",
  "Rectangular": "Hình chữ nhật",
  "Cubic": "Hình khối",
  "Flat": "Dẹt",
  "Compact": "Nhỏ gọn",
  "Threaded": "Ren",
  "Smooth": "Trơn",
  "Mounting": "Lắp đặt",
  "Mounting method": "Phương pháp lắp đặt",
  "Mounting type": "Kiểu lắp đặt",
  "Flush mounting": "Lắp chìm",
  "Non-flush mounting": "Lắp nổi",
  "Flush": "Lắp chìm",
  "Non-flush": "Lắp nổi",
  "Semi-flush": "Lắp nửa chìm",
  "Panel mounting": "Lắp panel",
  "DIN rail": "Ray DIN",
  "DIN rail mounting": "Lắp trên ray DIN",
  "Brackets": "Giá đỡ",
  "Swivel bracket": "Giá đỡ xoay",
  "Fixing": "Cố định",
  "Clamp": "Kẹp",

  // --- Ket noi ---
  "Connection type": "Loại kết nối",
  "Connection type Detail": "Chi tiết loại kết nối",
  "Cable": "Cáp",
  "Cable length": "Chiều dài cáp",
  "Cable material": "Vật liệu cáp",
  "Cable material detail": "Chi tiết vật liệu cáp",
  "Cable cross-section": "Tiết diện cáp",
  "Cable diameter": "Đường kính cáp",
  "Cable color": "Màu cáp",
  "Cable sheathing": "Vỏ bọc cáp",
  "Plug": "Giắc cắm",
  "Socket": "Ổ cắm",
  "Male connector": "Đầu nối đực",
  "Female connector": "Đầu nối cái",
  "Male connector M8": "Đầu nối đực M8",
  "Male connector M12": "Đầu nối đực M12",
  "Female connector M8": "Đầu nối cái M8",
  "Female connector M12": "Đầu nối cái M12",
  "Male connector M5": "Đầu nối đực M5",
  "Female connector M5": "Đầu nối cái M5",
  "Connector": "Đầu nối",
  "Connector type": "Loại đầu nối",
  "Pin assignment": "Sơ đồ chân",
  "Number of pins": "Số chân",
  "Terminal": "Đầu cực",
  "Screw terminal": "Đầu cực vít",
  "Spring terminal": "Đầu cực lò xo",
  "Wire": "Dây",
  "Communication interface": "Giao diện truyền thông",
  "IO-Link": "IO-Link",
  "IO-Link version": "Phiên bản IO-Link",
  "IO-Link interface": "Giao diện IO-Link",
  "RS-232": "RS-232",
  "RS-422": "RS-422",
  "RS-485": "RS-485",
  "Ethernet": "Ethernet",
  "PROFIBUS": "PROFIBUS",
  "PROFINET": "PROFINET",
  "EtherCAT": "EtherCAT",
  "EtherNet/IP": "EtherNet/IP",
  "DeviceNet": "DeviceNet",
  "CANopen": "CANopen",
  "Modbus": "Modbus",
  "Modbus RTU": "Modbus RTU",
  "Modbus TCP": "Modbus TCP",
  "SSI": "SSI",
  "SPI": "SPI",
  "I2C": "I2C",
  "USB": "USB",
  "Bluetooth": "Bluetooth",
  "WLAN": "WLAN",
  "Fieldbus": "Bus trường (Fieldbus)",

  // --- Cai dat ---
  "Setting method": "Phương pháp cài đặt",
  "Teach-in": "Dạy học (Teach-in)",
  "Auto-teach": "Tự động dạy học",
  "1-point teach": "Dạy học 1 điểm",
  "2-point teach": "Dạy học 2 điểm",
  "Dynamic teach": "Dạy học động",
  "Potentiometer": "Chiết áp",
  "Rotary potentiometer": "Chiết áp xoay",
  "Trimpot": "Chiết áp chỉnh tinh",
  "Adjustment": "Điều chỉnh",
  "Calibration": "Hiệu chuẩn",
  "Zero-point adjustment": "Chỉnh điểm không",
  "Span adjustment": "Chỉnh khoảng đo",
  "Offset": "Bù trừ (Offset)",
  "Gain": "Hệ số khuếch đại (Gain)",
  "Programming": "Lập trình",
  "Configuration": "Cấu hình",
  "Parameterization": "Tham số hóa",
  "Push button": "Nút nhấn",
  "DIP switch": "Công tắc DIP",
  "Rotary switch": "Công tắc xoay",
  "SAPS": "SAPS (Hệ thống xử lý tín hiệu tự động)",

  // --- Hien thi ---
  "Indication": "Hiển thị",
  "Status indicator": "Chỉ báo trạng thái",
  "LED indicator": "Đèn LED chỉ báo",
  "Display": "Màn hình hiển thị",
  "7-segment display": "Màn hình 7 đoạn",
  "Alphanumeric display": "Màn hình chữ-số",
  "Bargraph": "Thanh chỉ thị",
  "Output indicator": "Chỉ báo đầu ra",
  "Power indicator": "Chỉ báo nguồn",
  "Stability indicator": "Chỉ báo ổn định",
  "Alignment indicator": "Chỉ báo căn chỉnh",
  "Error indicator": "Chỉ báo lỗi",
  "Ready indicator": "Chỉ báo sẵn sàng",
  "Contamination indicator": "Chỉ báo bẩn",

  // --- Chung nhan ---
  "Application": "Ứng dụng",
  "Approvals": "Chứng nhận",
  "Ex-approvals": "Chứng nhận phòng nổ",
  "Ex approvals": "Chứng nhận phòng nổ",
  "Certifications": "Chứng chỉ",
  "Certificate": "Chứng chỉ",
  "Normative reference": "Tài liệu tham khảo quy chuẩn",
  "LED risk group marking": "Đánh dấu nhóm rủi ro LED",
  "Free group": "Nhóm phi rủi ro",
  "CE marking": "Dấu CE",
  "UL": "UL",
  "CSA": "CSA",
  "cULus": "cULus",
  "ATEX": "ATEX",
  "IECEx": "IECEx",
  "Ex": "Ex (Chống nổ)",
  "Ex zone": "Vùng Ex",
  "FM": "FM",
  "RoHS": "RoHS",
  "REACH": "REACH",
  "ECLASS": "ECLASS",
  "EU Declaration of Conformity": "Chứng nhận phù hợp tiêu chuẩn EU (Châu Âu)",
  "UK Declaration of Conformity": "Chứng nhận Vương quốc Anh (UKCA)",
  "ACMA Declaration of Conformity": "Chứng nhận phù hợp tiêu chuẩn ACMA (Úc)",
  "Moroccan Declaration of Conformity": "Chứng nhận phù hợp tiêu chuẩn Maroc",
  "China RoHS": "Tuân thủ quy định RoHS Trung Quốc",
  "cULus Certificate": "Chứng nhận cULus",
  "Ethernet/IP Certificate": "Chứng nhận giao thức truyền thông Ethernet/IP",
  "Photobiological safety": "An toàn quang sinh học",
  "UNSPSC": "UNSPSC",
  "FCC": "FCC",
  "Food-grade": "Cấp thực phẩm",
  "FDA": "FDA",
  "ECOLAB": "ECOLAB",
  "ETIM": "ETIM",
  "3A": "3A",
  "GL": "GL",
  "DNV": "DNV",
  "IP": "IP",

  // --- Danh muc ---
  "Product category": "Danh mục sản phẩm",
  "Product segment": "Phân khúc sản phẩm",
  "Product family": "Họ sản phẩm",
  "Product line": "Dòng sản phẩm",
  "Series": "Dòng sản phẩm",
  "Series number": "Số seri",
  "Part number": "Mã sản phẩm",
  "Order number": "Số đặt hàng",
  "Article number": "Mã hàng",
  "Type code": "Mã loại",
  "Type": "Loại",
  "Model": "Model",
  "Variant": "Biến thể",
  "Version": "Phiên bản",
  "Revision": "Bản sửa đổi",
  "Generation": "Thế hệ",

  // --- Logic ---
  "Yes": "Có",
  "No": "Không",
  "N/A": "N/A",
  "None": "Không có",
  "Optional": "Tùy chọn",
  "Standard": "Tiêu chuẩn",
  "modified": "đã sửa đổi",
  "Integrated": "Tích hợp",
  "External": "Ngoài",
  "Internal": "Trong",
  "Built-in": "Tích hợp sẵn",
  "Selectable": "Có thể lựa chọn",
  "Configurable": "Có thể cấu hình",
  "Fixed": "Cố định",
  "Adjustable": "Có thể điều chỉnh",
  "Manual": "Thủ công",
  "Automatic": "Tự động",
  "Digital": "Kỹ thuật số",
  "Analog": "Analog",
  "Active": "Chủ động",
  "Passive": "Thụ động",
  "Depending on the optical fiber cable used": "Tùy thuộc vào cáp quang được sử dụng",

  // --- Dau ra switching ---
  "PNP": "PNP",
  "NPN": "NPN",
  "Push-pull": "Push-pull",
  "Antivalent": "Đối kháng",
  "Complementary": "Bổ sung",
  "Normally open": "Thường mở (NO)",
  "Normally closed": "Thường đóng (NC)",
  "NO": "Thường mở (NO)",
  "NC": "Thường đóng (NC)",
  "NO/NC": "NO/NC",
  "PNP/NPN": "PNP/NPN",
  "PNP NO": "PNP Thường mở",
  "PNP NC": "PNP Thường đóng",
  "NPN NO": "NPN Thường mở",
  "NPN NC": "NPN Thường đóng",
  "Switching state": "Trạng thái đóng ngắt",
  "Switching point": "Điểm đóng ngắt",
  "Switching hysteresis": "Trễ đóng ngắt",
  "Switching distance": "Khoảng cách đóng ngắt",

  // --- Ky thuat ---
  "Technical data": "Thông số kỹ thuật",
  "Dimensional drawings": "Bản vẽ kích thước",
  "Operating reserve": "Dự trữ vận hành",
  "Signal reserve": "Dự trữ tín hiệu",
  "Excess gain": "Hệ số dự phòng ánh sáng",
  "Switching differential": "Khoảng trễ đóng ngắt",
  "Noise immunity": "Khả năng chịu nhiễu",
  "Cross-talk immunity": "Khả năng chịu nhiễu xuyên âm",
  "Ambient light": "Ánh sáng môi trường",
  "Background": "Nền",
  "Target": "Đối tượng phát hiện",
  "Object": "Vật thể",
  "Reflectance": "Độ phản xạ",
  "Remission": "Độ phản xạ khuếch tán",
  "Attenuation": "Suy hao",
  "Transmission": "Truyền dẫn",
  "Absorption": "Hấp thụ",
  "Diffuse reflection": "Phản xạ khuếch tán",
  "Specular reflection": "Phản xạ gương",

  // --- An toan ---
  "Safety integrity level (SIL)": "Mức toàn vẹn an toàn (SIL)",
  "SIL": "Mức toàn vẹn an toàn (SIL)",
  "Performance level (PL)": "Mức hiệu suất (PL)",
  "PL": "Mức hiệu suất (PL)",
  "Category": "Loại",
  "MTTFD": "MTTFD",
  "MTTFd": "MTTFd (Thời gian trung bình đến lỗi nguy hiểm)",
  "PFHd": "PFHd (Xác suất lỗi nguy hiểm mỗi giờ)",
  "DCavg": "DCavg (Độ bao phủ chẩn đoán trung bình)",
  "Typ.": "Điển hình",
  "Avg.": "Trung bình",
  "CCF": "CCF (Lỗi nguyên nhân chung)",
  "Mission time": "Thời gian hoạt động",
  "Proof test interval": "Khoảng thời gian kiểm tra xác nhận",
  "Safe state": "Trạng thái an toàn",
  "Fault detection": "Phát hiện lỗi",
  "Self-test": "Tự kiểm tra",
  "Diagnostic coverage": "Độ bao phủ chẩn đoán",
  "Redundancy": "Dự phòng",
  "Dual channel": "Kênh kép",
  "Single channel": "Kênh đơn",
  "OSSD1": "OSSD1",
  "OSSD2": "OSSD2",
  "Muting": "Muting (Tắt tạm thời)",
  "Override": "Override (Ghi đè)",
  "Reset": "Đặt lại",
  "Manual reset": "Đặt lại thủ công",
  "Auto reset": "Đặt lại tự động",
  "Restart interlock": "Khóa khởi động lại",
  "Protective field": "Vùng bảo vệ",
  "Warning field": "Vùng cảnh báo",
  "Detection field": "Vùng phát hiện",
  "Contactor monitoring": "Giám sát công tắc tơ",
  "EDM": "EDM (Giám sát thiết bị ngoài)",
  "External device monitoring": "Giám sát thiết bị ngoài",

  // --- Cac cum tu pho bien trong gia tri ---
  "Visible red": "Đỏ khả kiến",
  "Infrared LED": "LED hồng ngoại",
  "Red LED": "LED đỏ",
  "LED Red": "LED đỏ",
  "Green LED": "LED xanh lá",
  "LED Green": "LED xanh lá",
  "Yellow LED": "LED vàng",
  "LED Yellow": "LED vàng",
  "Orange LED": "LED cam",
  "LED Orange": "LED cam",
  "Blue LED": "LED xanh dương",
  "LED Blue": "LED xanh dương",
  "Bicolor LED": "LED hai màu",
  "LED Bicolor": "LED hai màu",
  "Laser, red": "Laser, đỏ",
  "Laser, infrared": "Laser, hồng ngoại",
  "Laser, green": "Laser, xanh lá",
  "3-pin": "3 chân",
  "4-pin": "4 chân",
  "5-pin": "5 chân",
  "8-pin": "8 chân",
  "12-pin": "12 chân",
  "With cable": "Có cáp",
  "Without cable": "Không có cáp",
  "With connector": "Có đầu nối",
  "Without connector": "Không có đầu nối",
  "Pre-wired": "Đã đấu dây sẵn",
  "Plug-in": "Trình cắm",
  "Flying leads": "Dây đầu tự do",
  "Unterminated": "Không đầu nối",
  "10 mm": "10 mm",
  "Stainless steel 1.4305": "Thép không gỉ 1.4305",
  "Stainless steel 1.4404": "Thép không gỉ 1.4404",
  "Polycarbonate": "Polycarbonate",
  "Polyester": "Polyester",
  "Polyurethane": "Polyurethane",
  "Silicone": "Silicone",
};

// Pre-computed lowercase Map for O(1) dictionary lookup
const TERM_DICT_LOWER = new Map(
  Object.entries(TERM_DICT).map(([k, v]) => [k.toLowerCase().trim(), v])
);

// === DANH SACH TU KHONG DUOC DICH (giu nguyen) ===
// Regex patterns - text PHAI match HOAN TOAN (anchored ^...$) thi moi giu nguyen.
// Neu chi match prefix ma khong anchor cuoi se gay bug:
// "150 mm ... + tieng Anh dai" cung bi keep -> khong dich tieng Anh.
const NO_TRANSLATE_PATTERNS = [
  /^LED$/i,
  /^PNP$/i,
  /^NPN$/i,
  /^RFID$/i,
  /^IO-?Link$/i,
  /^ATEX$/i,
  /^SIL\s?\d+$/i,
  /^PL\s?[a-e]$/i,
  /^IP\s?\d+K?$/i,
  /^M\d+$/i,                       // M8, M12
  // IEC 61508, EN 14119, ISO 13849-1:2015, EN ISO 13849-1, EN IEC 60947
  /^((IEC|EN|ISO|UL)\s)+\d+(-\d+)*(:\d+(-\d+)*)?$/i,
  /^CE$/i,
  /^CSA(\s\d+)?$/i,
  /^EAC$/i,
  /^CCC$/i,
  /^MTTFD$/i,
  /^PROFINET$/i,
  /^EtherNet(\/IP)?$/i,
  /^EtherCAT$/i,
  /^PROFIBUS$/i,
  /^CANopen$/i,
  /^DeviceNet$/i,
  /^Modbus(\s(RTU|TCP))?$/i,
  /^RS-?\d+$/i,                    // RS-232, RS-485
  /^USB(\s\d(\.\d)?)?$/i,
  /^HART$/i,
  /^SSI$/i,
  // So + don vi (whole string only)
  /^[-+]?\d+([.,]\d+)?\s*(mm|cm|m|km|kg|g|mg|µm|mA|A|V|kV|W|kW|Hz|kHz|MHz|GHz|ms|µs|ns|s|h|min|°C|°F|K|nm|lx|cd|dB|Ω|hPa|bar|psi|°|rpm|m\/s|N|Nm|J|VA|VAr|F|µF|nF|pF|°|‰|%)$/i,
  // Range like "-25 ... 55", "-25 ... 55 °C", "150 mm ... 2,000 mm"
  /^[-+]?\d+([.,]\d+)?(\s*[a-zA-Z°]+\d?)?\s*\.{3}\s*[-+]?\d+([.,]\d+)?(\s*[a-zA-Z°]+\d?)?$/i,
  // Product codes like CLV62x, IME30, WTT12-S
  /^[A-Z]{2,}[\d-]+[A-Z]*\d*$/i,
  // Pure numbers
  /^[-+]?\d+([.,]\d+)?$/,
  // Standards/codes (whole string)
  /^ECLASS\s[\d.]+$/i,
  /^ETIM\s[\d.]+$/i,
  /^UNSPSC\s[\d.]+$/i,
  /^NRKH[\w.]*$/i,                 // UL file numbers like NRKH.E300503
  /^EC\d+$/i,                      // ETIM codes
];

// === TU KHONG DICH (exact match, case-insensitive) ===
const NO_TRANSLATE_EXACT = new Set([
  "led", "pnp", "npn", "rfid", "io-link", "iolink",
  "push-pull", "push pull", "rs-232", "rs-485", "rs232", "rs485",
  "canopen", "profinet", "profibus", "ethercat", "ethernet/ip",
  "ethernet", "modbus", "devicenet", "hart", "ssi", "usb",
  "pmma", "abs", "pbt", "pa", "pc", "pur", "pvc", "tpe",
  "ip65", "ip67", "ip68", "ip69", "ip69k",
  "m8", "m12", "m18", "m23", "m30",
  "atex", "iecex", "ul", "csa", "ce", "eac", "ccc",
  "sil", "sil1", "sil2", "sil3", "pl", "cat",
  "mttfd", "mtbf",
  "dc", "ac", "ac/dc",
  "n/a", "n.a.",
  "sick",
  "eclass", "etim", "unspsc",
  "rohs", "reach", "weee",
  "crUus", "cruus",
  "pet", "pom",
]);

/**
 * Kiem tra text co nen giu nguyen (khong dich) hay khong.
 */
function shouldKeepOriginal(text) {
  const trimmed = text.trim();
  if (!trimmed) return true;

  // Exact match
  if (NO_TRANSLATE_EXACT.has(trimmed.toLowerCase())) return true;

  // Pattern match
  for (const pat of NO_TRANSLATE_PATTERNS) {
    if (pat.test(trimmed)) return true;
  }

  // Text chi chua so, don vi, ky hieu (khong co chu cai tieng Anh co nghia)
  if (/^[\d\s.,°×xX\-+/|<>≤≥%()μµ]+$/.test(trimmed)) return true;

  // Text qua ngan (1-2 ky tu) hoac la ma san pham
  if (trimmed.length <= 2) return true;

  return false;
}

/**
 * Dich text qua Google Translate voi bao ve thuat ngu ky thuat.
 * Truoc khi gui Google, thay the cac tu ky thuat bang placeholder,
 * sau khi nhan ket qua thi tra lai tu goc.
 */
// Pre-translate cac viet tat thong dung TRUOC khi gui Google Translate.
// Google hay dich sai cac viet tat ngan: "Typ." -> "Đánh máy" (typing),
// "Avg." -> "Trung bình" (OK), "Approx." -> "Tương đối" (sai context).
// Bang nay thay the trong text TRUOC khi gui Google, neu Google nhan thay tieng Viet
// thi giu nguyen.
const INLINE_TRANSLATIONS = [
  [/\bTyp\.(?=\s|$|[,;:])/g, "Điển hình"],
  [/\bTypical\b/gi, "Điển hình"],
  [/\bAvg\.(?=\s|$|[,;:])/g, "Trung bình"],
  [/\bAverage\b/gi, "Trung bình"],
  [/\bApprox\.(?=\s|$|[,;:])/g, "Khoảng"],
  [/\bApproximately\b/gi, "Khoảng"],
  [/\bMax\.(?=\s|$|[,;:])/g, "Tối đa"],
  [/\bMin\.(?=\s|$|[,;:])/g, "Tối thiểu"],
  [/\bNo\.(?=\s|$|[,;:])/g, "Số"],
  [/\bRef\.(?=\s|$|[,;:])/g, "Tham chiếu"],
  [/\bIncl\.(?=\s|$|[,;:])/g, "Bao gồm"],
  [/\bExcl\.(?=\s|$|[,;:])/g, "Không bao gồm"],
  [/\bResp\.(?=\s|$|[,;:])/g, "Tương ứng"],
];

async function smartTranslate(text, retries = 3) {
  if (!text || !text.trim()) return text;
  if (shouldKeepOriginal(text)) return text;

  // Tim tat ca tu/cum tu ky thuat trong text va thay bang placeholder
  const protectedTerms = [];
  let processedText = text;

  // Buoc 1: Pre-translate viet tat thong dung (Google hay dich sai)
  for (const [pat, repl] of INLINE_TRANSLATIONS) {
    processedText = processedText.replace(pat, repl);
  }

  // Buoc 2: Bao ve cac pattern ky thuat trong chuoi (giu nguyen, khong dich)
  // DEC, HEX la viet tat (decimal, hexadecimal) khong phai thang Dec/December
  const inlinePatterns = [
    /\b(LED|PNP|NPN|RFID|IO-Link|ATEX|SIL\d?|IP\d+K?|M\d+|PMMA|PBT|ABS|DC|AC|PVC|POM|PET|PC|DEC|HEX|RGB|HSV)\b/g,
    /\b(EN|IEC|ISO|UL|ECLASS|ETIM|UNSPSC)\s+\d[\d.:/-]*/gi,
    /\d+[\s.,]?\d*\s*(mm|cm|m|kg|g|mA|A|V|W|Hz|kHz|MHz|ms|µs|ns|°C|°F|nm|lx)\b/gi,
    /[-+]?\d+\s*\.{3}\s*[-+]?\d+/g,  // ranges: -25 ... 55
    /0x[0-9A-Fa-f]+/g,  // hex numbers like 0x80024A
  ];

  for (const pat of inlinePatterns) {
    processedText = processedText.replace(pat, (match) => {
      const idx = protectedTerms.length;
      protectedTerms.push(match);
      return `__PROT${idx}__`;
    });
  }

  // Gui Google Translate voi retry
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const url = `https://translate.googleapis.com/translate_a/single?client=gtx&sl=en&tl=vi&dt=t&q=${encodeURIComponent(processedText)}`;
      const resp = await fetch(url, {
        headers: { "user-agent": "Mozilla/5.0" },
      });
      if (resp.status === 429) {
        // Rate limited - wait and retry
        await new Promise((r) => setTimeout(r, 1000 * attempt));
        continue;
      }
      if (!resp.ok) {
        if (attempt < retries) {
          await new Promise((r) => setTimeout(r, 500 * attempt));
          continue;
        }
        return text;
      }
      const data = await resp.json();
      let translated = (data?.[0] || []).map((p) => p?.[0] || "").join("");

      // Tra lai cac tu da bao ve
      for (let i = 0; i < protectedTerms.length; i++) {
        const variants = [
          `__PROT${i}__`,
          `__prot${i}__`,
          `__ PROT${i}__`,
          `__PROT${i} __`,
          `__ PROT${i} __`,
          `__Prot${i}__`,
        ];
        for (const v of variants) {
          translated = translated.split(v).join(protectedTerms[i]);
        }
        // Fallback regex
        const regex = new RegExp(`_*\\s*PROT\\s*${i}\\s*_*`, "gi");
        translated = translated.replace(regex, protectedTerms[i]);
      }

      return translated || text;
    } catch {
      if (attempt === retries) return text;
      await new Promise((r) => setTimeout(r, 500 * attempt));
    }
  }
  return text;
}

/**
 * Dich batch texts, uu tien tu dien -> smart translate cho phan con lai.
 */
// === Translation disk cache ===
const CACHE_PATH = resolve(process.cwd(), "translation-cache.json");
let _diskCache = null;

function loadTranslationCache() {
  if (_diskCache) return _diskCache;
  try {
    if (existsSync(CACHE_PATH)) {
      _diskCache = JSON.parse(readFileSync(CACHE_PATH, "utf8"));
    } else {
      _diskCache = {};
    }
  } catch {
    _diskCache = {};
  }
  return _diskCache;
}

function saveTranslationCache() {
  try {
    if (_diskCache) writeFileSync(CACHE_PATH, JSON.stringify(_diskCache, null, 2), "utf8");
  } catch { /* ignore */ }
}

async function translateBatchSmart(texts) {
  const unique = [...new Set(texts.filter((t) => t && t.trim()))];
  const result = {};
  const cache = loadTranslationCache();

  const needTranslate = [];
  for (const text of unique) {
    // Check dictionary first (exact match)
    if (TERM_DICT[text]) {
      result[text] = TERM_DICT[text];
      continue;
    }
    // Check case-insensitive dictionary (O(1) via TERM_DICT_LOWER)
    const lowerMatch = TERM_DICT_LOWER.get(text.toLowerCase().trim());
    if (lowerMatch) {
      result[text] = lowerMatch;
      continue;
    }
    // Check if should keep original
    if (shouldKeepOriginal(text)) {
      result[text] = text;
      continue;
    }
    // Check disk cache
    if (cache[text]) {
      result[text] = cache[text];
      continue;
    }
    needTranslate.push(text);
  }

  // Smart translate in batches
  const concurrency = 5;
  for (let i = 0; i < needTranslate.length; i += concurrency) {
    const batch = needTranslate.slice(i, i + concurrency);
    const promises = batch.map(async (text) => {
      const translated = await smartTranslate(text);
      return [text, translated];
    });
    const done = await Promise.all(promises);
    for (const [src, dst] of done) {
      result[src] = dst;
      cache[src] = dst;
    }
    // Save cache periodically
    if ((i / concurrency) % 10 === 9) saveTranslationCache();
    if (i + concurrency < needTranslate.length) {
      await new Promise((r) => setTimeout(r, 300));
    }
  }

  // Save cache at end of batch
  if (needTranslate.length > 0) saveTranslationCache();

  return result;
}

module.exports = { TERM_DICT, translateBatchSmart, shouldKeepOriginal };
