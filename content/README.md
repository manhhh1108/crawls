# content/ — crawl nội dung bài đăng web NTREND (mỗi dòng SP = 1 bài)

Mỗi bài = 1 dòng sản phẩm. Nội dung theo template `data/ntrend template.xlsx` +
review trong `data/Crawl nội dung Ntrend.docx`, dựng 2 bản: **cột B tiếng Việt
(bản đăng web)**, **cột C tiếng Anh (chỉ để soát chính tả bản dịch)**.

## Cấu trúc

```
content/
├─ translate.py            EN→VI qua translate.googleapis.com
│                          vi()      1 chuỗi
│                          vi_many() gộp nhiều chuỗi / 1 request + 6 luồng (~20× nhanh hơn)
│                          cache .cache/translate_vi.json
├─ vi_terms.py             TỪ ĐIỂN thuật ngữ EN→VI (879 cặp) - xem mục "Dịch" bên dưới
├─ vi_engine.py            máy dịch có kiểm soát: tra từ điển → mask → dịch → sửa hậu kỳ
├─ brandimg.py             ghép ảnh SP: nền trắng + logo NTREND + hotline (tỉ lệ theo Canva)
├─ assets/                 ntrend-logo.png, phone-icon.png
│
├─ c_perma.py              Perma: crawl + dịch        → out/05 Perma/content.json
├─ build_content_excel.py  Perma: content.json        → data/Noi dung Perma/
│
├─ c_sick_fetch.py         SICK pha 1: API sick.com   → out/01 SICK/raw.json
├─ sick_tdo.py             SICK pha 1b: bóc bảng TECHNICAL DATA OVERVIEW
│                          trong PDF "Product overview" → out/01 SICK/tdo.json
├─ sick_name_translate.py  dịch TÊN dòng SP (mã → giữ, cụm tiếng Anh → dịch)
│                          → out/01 SICK/names_vi.json + data/SICK - ten dong SP can dich.xlsx
├─ ntrend_cat.py           đọc danh mục Ntrend từ data/Danh mục SP web NTREND.xlsx (sheet 01 SICK)
├─ c_sick.py               SICK pha 2: dịch + ghép ảnh→ out/01 SICK/content.json
├─ build_sick_excel.py     SICK pha 3: content.json   → data/Noi dung SICK/
└─ out/<hãng>/             content.json + src_img/ (ảnh gốc) + images/ (ảnh đã ghép)
```

## Chạy

```bash
cd content
python c_perma.py && python build_content_excel.py          # Perma

# SICK - chạy theo đúng thứ tự này
python c_sick_fetch.py        # 1  API sick.com      -> raw.json
python sick_tdo.py            # 1b PDF Product overview -> tdo.json  (lâu nhất, ~25')
python c_sick_fetch.py        # 1c chạy lại: biết dòng nào thiếu TDO -> lấy thêm
                              #    cột Description của từng variant
python sick_name_translate.py # 1d dịch tên dòng SP   -> names_vi.json
python c_sick.py              # 2  dịch nội dung      -> content.json
python build_sick_excel.py    # 3  xuất Excel         -> data/Noi dung SICK/
```

Sửa vặt mà **không phải crawl lại**:

| Muốn sửa | Sửa ở đâu | Chạy lại |
|---|---|---|
| Danh mục Ntrend của 1 dòng SP | `data/Danh mục SP web NTREND.xlsx`, sheet `01 SICK`, cột D | `c_sick.py` + `build_sick_excel.py` |
| Tên tiếng Việt của dòng SP | `data/SICK - ten dong SP can dich.xlsx` cột C | `sick_name_translate.py --from-xlsx` + `build_sick_excel.py` |
| Một thuật ngữ dịch sai | `data/Tu dien dich EN-VI.xlsx` | `c_sick.py` + `build_sick_excel.py` |

`c_sick.py 20` = chỉ làm 20 dòng đầu (xem thử nhanh).

## SICK (hãng 01)

Đầu vào: `data/SICK_Web_Products_Statistics.xlsx`, sheet **"Thong ke SP SICK Web SP"**,
các dòng có `x` ở cột G = 618 dòng sản phẩm cần viết bài. Cột *"Danh mục SP Tiếng
Việt"* được điền thẳng vào ô **Danh mục sản phẩm** của bài.

Dữ liệu lấy từ API của sick.com (web là SPA, HTML rỗng):

| API | Cho ra |
|---|---|
| `fact-finder/records/detail/enSG-sick/<gcode>` | mô tả, slogan, Features, Benefits, lĩnh vực ứng dụng, ảnh, thông số toàn dòng |
| `fact-finder/search?...ProductFamilyCode:<g>&idsOnly=true` | danh sách id biến thể (rất nhẹ) |
| `api.sick.com/occ/v2/sg/sick-products?productCodes=…` | tên model + part number (~0,9 KB/SP) |

### Thông số kỹ thuật lấy ở đâu (sửa theo review 13/10)

Trước đây quét các **facet lọc ở tab Variants** — ra toàn metadata catalog
(`Certificate`, `New products`, `Product segment`, `Task`…), đăng web vô nghĩa.

Nay theo đúng review: lấy bảng **TECHNICAL DATA OVERVIEW** trong tài liệu
*Product overview* (PDF) của hãng — `sick_tdo.py` tải PDF, bóc bảng, xoá file.
Bảng này **không có trong API**: trường `TechData` ở cấp dòng SP luôn rỗng.

| Tình trạng | Số dòng | Bảng thông số trong bài |
|---|---|---|
| PDF có mục TECHNICAL DATA OVERVIEW | 403 | "Thông số chi tiết" = đúng bảng đó + "Bảng mã sản phẩm" riêng |
| Không có PDF, hoặc PDF không có mục đó | 155 | "Thông số chi tiết (bảng mã ở tab Variants)" = Model \| Part No. \| Mô tả |

Dòng SP hãng **không công bố Features/Benefits** (83 bài) thì "Mô tả dài" lấy y
nguyên "Mô tả ngắn" — trước đây ghép từ metadata catalog nên ra list vô nghĩa.

Bảng mã sản phẩm in hết, không cắt.

## Dịch — vì sao không dùng thẳng Google

Google dịch thẳng sai rất nhiều thuật ngữ ngành (`Housing` → "Nhà ở",
`Bus termination` → "Chấm dứt xe buýt", `Art. 3` → "Nghệ thuật 3"). `vi_engine`
chạy 3 lớp:

1. **Tra từ điển trước** (`vi_terms.py`, 879 cặp) — khớp nguyên chuỗi thì dùng
   luôn, không gọi máy dịch. Nguồn: `sick/sick-translate.js` (662 cặp của dự án
   SICK cũ) + cột `suggested_fix` trong `sick/loi dich san pham.xlsx` (lỗi đã
   soát tay) + `EXTRA` bổ sung tay.
2. **Mask** những thứ cấm dịch trước khi đưa qua máy: mã model/part
   (`W16`, `FLX3-CPUN100`), tên riêng của SICK (`Flexi Loop`, `Safety Designer`),
   tiêu chuẩn (`EN ISO 13849-1`, `IEC 60529`), giao thức (`IO-Link`, `PROFINET`),
   ký hiệu (`≤ ± Ø ✔ °C`), số có dấu phân cách.
   Câu văn dài thì mask nhẹ hơn (nếu mask dày quá máy dịch bỏ cuộc, trả lại nguyên EN).
3. **Sửa hậu kỳ**: bộ regex chặn lỗi kinh điển + đổi số sang kiểu Việt
   (`1,200` → `1.200`, `0.35` → `0,35`).

Sau cùng `suspicious()` chấm điểm nghi ngờ từng câu → `00 CAN SOAT DICH.xlsx`
trong thư mục kết quả, chỉ liệt kê câu cần soát tay.

**Muốn sửa 1 thuật ngữ dịch sai:** thêm cặp vào `EXTRA` trong `vi_terms.py`, rồi
chạy lại `python c_sick.py && python build_sick_excel.py` (crawl đã cache nên nhanh).

## Lưu ý đã khảo sát trước (2026-09-30)

| Hãng | Tình trạng nội dung trên web |
|---|---|
| Festo, IFM, Schneider, Pepperl, Pilz, Leuze, Balluff, Reer, wenglor, Mitsubishi, Perma | Đầy đủ — crawl bình thường |
| Supmea EN, ESPE, OTENNLUX EN | Đủ ảnh + text |
| **Xingguang** | Ảnh SP có, nhưng thông số là **ảnh chụp bảng** → cần OCR hoặc nhập tay |
| **Tree Electric** | **Không có trang SP** — model nằm trong catalogue PDF, phải trích từ PDF |

- Ảnh trên nhiều web là lazy-load (`data-src`, `lazy=`…) — đừng chỉ lấy `src`.
- Code SICK đời cũ (crawl theo từng SKU) vẫn giữ trong `sick/` để tham khảo.
