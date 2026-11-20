# SICK — đối chiếu dòng sản phẩm: file thống kê cũ vs web hiện tại (2026-10-12)

Nguồn mới: `catalog/b_sick.py` → sheet **01 SICK** trong `data/Danh muc san pham 16 hang - FULL.xlsx`.

| | Số dòng SP |
|---|---|
| `SICK_Web_Products_Statistics.xlsx` (cũ) | 589 |
| Web sick.com/sg/en hiện tại (nhánh `Products`) | 607 |
| Không còn trên web (đã chuyển sang nhánh `Archive`) | 40 |
| Mới, chưa có trong file cũ | 58 |

## 40 dòng SP KHÔNG còn trên sick.com/sg/en

Các dòng này đã bị SICK chuyển vào nhánh `Archive` (ngừng bán / đổi tên). Bài nội dung đã viết cho chúng cần gắn cảnh báo hoặc trỏ sang dòng thay thế — VD `RMS2000` (`g577390`) nay được thay bằng `RMS` (`g628337`).

| Mã danh mục | Tên dòng SP | Số SP trong file cũ |
|---|---|---|
| `g385751` | APS | 1 |
| `g187251` | C4000 Advanced | 187 |
| `g304854` | C4000 Advanced ATEX II 3G/3D | 6 |
| `g304853` | C4000 Advanced Ex | 12 |
| `g187260` | C4000 Entry/Exit | 15 |
| `g339951` | C4000 Entry/Exit ATEX II 3G/3D | 1 |
| `g187259` | C4000 Fusion | 55 |
| `g339952` | C4000 Fusion ATEX II 3G/3D | 7 |
| `g543465` | C4000 Fusion Ex | 12 |
| `g187263` | C4000 Palletizer | 63 |
| `g552041` | C4000 Palletizer ATEX II 3G/3D | 1 |
| `g365751` | DBV50 | 98 |
| `g383757` | DFV60 | 15 |
| `g383551` | DKV60 | 49 |
| `g168298` | DME4000 | 0 |
| `g403052` | DUV60 | 108 |
| `g187988` | ELG | 88 |
| `g585667` | End-of-Arm-Safeguard | 2 |
| `g187986` | HLG | 2 |
| `g187283` | M2000 Cascadable | 14 |
| `g187282` | M2000 RES/EDM | 22 |
| `g187285` | M2000 RES/EDM A/P | 3 |
| `g187280` | M2000 Standard | 22 |
| `g187284` | M2000 Standard A/P | 3 |
| `g244219` | M4000 Advanced | 88 |
| `g305783` | M4000 Advanced A/P | 20 |
| `g187265` | M4000 Advanced Curtain | 134 |
| `g187279` | M4000 Area | 33 |
| `g244218` | M4000 Standard | 247 |
| `g311876` | M4000 Standard A/P | 43 |
| `g555591` | RMS1000 | 3 |
| `g577390` | RMS2000 | 5 |
| `g412755` | Safe AGV Easy | 6 |
| `g560275` | Tag-LOC System | 5 |
| `g606109` | UE402 | 1 |
| `g606110` | UE403 | 1 |
| `g355653` | Upgrade kit FGS to deTec4 Core | 22 |
| `g409651` | Upgrade kit MSL to M4000 | 17 |
| `g211351` | W250-2 | 33 |
| `g211356` | W280-2 | 31 |

## 58 dòng SP MỚI (chưa có trong file thống kê cũ)

| Mã danh mục | Đường dẫn danh mục |
|---|---|
| `g184875` | Detection sensors > Photoelectric sensors > W100 |
| `g203757` | Detection sensors > Photoelectric sensors > W250 |
| `g204129` | Detection sensors > Photoelectric sensors > W280 |
| `g616808` | Digital services and software > Application software > AI-Assist |
| `g576968` | Digital services and software > Application software > CODE-LOC |
| `g541272` | Digital services and software > Application software > LiDAR-LOC |
| `g609717` | Digital services and software > Application software > Track and Trace Vision Software |
| `g546067` | Digital services and software > Engineering Tools > Function Block Factory |
| `g622676` | Digital services and software > Engineering Tools > Sentio Creator |
| `g547567` | Digital services and software > Engineering Tools > SICK AppEngine |
| `g446551` | Digital services and software > Engineering Tools > SICK AppManager |
| `g387652` | Digital services and software > Engineering Tools > SICK AppStudio |
| `g547657` | Digital services and software > Engineering Tools > SICK dStudio |
| `g609542` | Digital services and software > Engineering Tools > SICK Stream Software |
| `g584721` | Digital services and software > Engineering Tools > SOPAS ET |
| `g363451` | Digital services and software > Insight Tools > Baggage Analytics |
| `g626751` | Digital services and software > Insight Tools > Condition Monitoring |
| `g623071` | Digital services and software > Insight Tools > Energy Monitoring |
| `g555630` | Digital services and software > Insight Tools > FieldAnalytics |
| `g584843` | Digital services and software > Insight Tools > Incoming Goods Suite |
| `g624833` | Digital services and software > Insight Tools > Location Management |
| `g448352` | Digital services and software > Insight Tools > Logistics Diagnostic Analytics |
| `g309559` | Digital services and software > Insight Tools > Package Analytics |
| `g559811` | Digital services and software > Insight Tools > Safety Laser Scanner Visualization |
| `g574738` | Digital services and software > Insight Tools > SICK Safety Assistant App |
| `g448351` | Digital services and software > Insight Tools > Tire Analytics |
| `g576768` | Digital services and software > Insight Tools > Track and Trace Analytics Lite |
| `g555711` | Digital services and software > Integration software > Hermes Standard Solution |
| `g584460` | Digital services and software > Integration software > SICK ConnectX |
| `g628337` | LiDAR and radar sensors > Radar sensors > RMS |
| `g625244` | Machine vision and identification > Image-based code readers > sensingCam SEC200 |
| `g623175` | Motion control sensors > Motor feedback systems > ENS/ENM35 |
| `g628391` | Network and connection technology > Connectors and cables > Power cables |
| `g584237` | Service > Consulting > Accident investigation |
| `g584233` | Service > Consulting > Concept and feasibility |
| `g584232` | Service > Consulting > Safety and risk assessment |
| `g584243` | Service > Engineering and integration > Acceptance services |
| `g584261` | Service > Engineering and integration > Extended warranty |
| `g584260` | Service > Engineering and integration > Initial verification |
| `g584248` | Service > Engineering and integration > Installation and commissioning |
| `g584240` | Service > Engineering and integration > Solution Engineering |
| `g584328` | Service > Maintenance and inspection > Inspection of protective devices |
| `g584332` | Service > Maintenance and inspection > Machine safety inspection |
| `g584340` | Service > Maintenance and inspection > Maintenance |
| `g584346` | Service > Maintenance and inspection > Monitoring Services |
| `g584335` | Service > Maintenance and inspection > Performance check |
| `g584344` | Service > Maintenance and inspection > Service agreements |
| `g584327` | Service > Maintenance and inspection > Smart Parts |
| `g584334` | Service > Maintenance and inspection > Software and firmware maintenance |
| `g584331` | Service > Maintenance and inspection > Stop time measurement |
| `g584223` | Service > Modernization > Upgrade and retrofit |
| `g584321` | Service > Technical support > Preparation for reverification |
| `g584272` | Service > Technical support > Repairs |
| `g584271` | Service > Technical support > Troubleshooting |
| `g616798` | Training > Training and education > Product, system, and software trainings |
| `g616799` | Training > Training and education > Safety competence trainings |
| `g616796` | Training > Training and education > Specialized and methodology trainings |
| `g616801` | Training > Training and education > Technology trainings |
