# Danh mục sản phẩm 16 hãng — bản crawl lại đầy đủ (2026-09-29)

**File kết quả:** `data/Danh muc san pham 16 hang - FULL.xlsx`
**Code crawl:** thư mục `catalog/` ở gốc repo (`b_<hãng>.py`, chạy `python build_excel.py` để dựng lại file Excel; hạ tầng HTTP/browser dùng chung nằm ở `shared/`)

## Quy tắc lấy dữ liệu
Lấy cây danh mục **xuống tới cấp thấp thứ 2 từ dưới lên** — tức là cấp ngay trên mã sản phẩm cụ thể.
Ví dụ Festo:

```
Actuators and drives > Pneumatic cylinders > Piston rod cylinders > Round cylinders > Round cylinder DSNU-S   <-- LẤY ĐẾN ĐÂY
                                                                                     └─ DSNU-12-70-P-A       <-- mã SP, KHÔNG lấy
```

Cột **"La cap ngay tren ma SP"** đánh dấu `x` ở đúng những dòng là cấp cuối đó.

## Cấu trúc file
- `00 TONG HOP` — số danh mục theo từng cấp cho từng hãng.
- Mỗi hãng 1 sheet: `STT` (đánh số phân cấp 1 / 1.1 / 1.1.1…), tên có thụt dòng, `Cấp`,
  các cột `Cấp 1..Cấp 6` (đường dẫn phẳng để lọc/pivot), cờ cấp cuối, link, mã danh mục, nguồn.

## So với file cũ
| Hãng | Cũ | Mới |
|---|---|---|
| Festo | 232 | **1.527** |
| IFM | 163 | **920** |
| Mitsubishi Electric | 164 | **476** |
| wenglor | 73 | **181** |
| Pilz | 46 | **290** |
| Reer | 99 | **294** |
| Pepperl+Fuchs | 307 | **381** |
| Leuze | 183 | **223** |
| Balluff | 178 | **199** |
| Perma | 41 | **95** |
| Xingguang / Starshine | 23 | **26** |
| Schneider | 1.232 | 1.218 |
| Supmea / Tree / ESPE / OTENNLUX | 157 | 152 |
| **Tổng** | **2.898** | **5.982** |

Riêng **Supmea** và **OTENNLUX** còn crawl thêm **mã sản phẩm cụ thể** (type `product`,
104 + 156 dòng, chữ nghiêng xám trong Excel) nằm dưới cấp đánh dấu `x`; các dòng này
không tính vào bảng đếm danh mục ở trên.

Schneider / Supmea / ESPE / OTENNLUX / Tree Electric gần như không đổi — cấp cuối của các hãng này
trong file cũ **đã đúng** là cấp ngay trên mã SP.

## Nguồn dữ liệu từng hãng
- **Festo** — API `/search/categories/{pimId}` (cây danh mục) + `/products` (dòng sản phẩm). Cấp cuối = product family (VD `Round cylinder DSNU-S`).
- **IFM** — REST `menu/navbar` (cấp 1–3) + render trang danh mục để lấy cấp 4/5 (product range).
- **Mitsubishi** — `megadropdown_products.json` (cấp 1–3) + danh sách nhóm SP trên trang series (cấp 4).
- **Schneider** — sitemap category / subcategory / range của `se.com/ca/en` (bản EN đầy đủ nhất). Cấp 3 = *range* (VD `Harmony XB4`).
- **Pepperl+Fuchs** — navigation (cấp 1–2) + thẻ danh mục trên trang (render bằng Chrome). 162 mục cấp 3 của file cũ chưa thấy lại trên web được giữ nguyên, đánh dấu `file cu` ở cột **Nguon**.
- **Pilz** — `www.pilz.com` chặn bot bằng Cloudflare (curl, Chromium headless và Chrome thật đều không qua). Dùng **mirror Trung Quốc `www.pilz.com.cn`, bản tiếng Anh `/en-CN/`** — cùng catalogue, crawl qua eShop nên có tới 4 cấp.
- **Reer / Leuze / Balluff / Perma / wenglor** — sitemap + breadcrumb (JSON-LD hoặc microdata).
- **Supmea / Xingguang / OTENNLUX** — crawl từ **bản EN** (`en.supmea.com`, `en.xgcd.cn`, `en.otennlux.com`), tên tiếng Anh. Lưu ý bản EN là catalogue riêng của hãng nên hơi khác bản CN: Supmea EN không có nhóm "Display instruments" nhưng thêm "Valve" và "System products" (Recorder, Process indicator…); OTENNLUX CN có 5 dòng SP chưa đưa lên bản EN.
- **Mã SP Supmea** lấy từ trang từng dòng SP (model đứng đầu tên: FMC240, FMX400, SUP-…); 3 dòng SP chưa có sản phẩm nào trên web (Digital pressure gauge, TSS/SS, Online water quality analyzer). **Mã SP OTENNLUX** lấy từ grid trên trang danh mục (mã `_pNN`; trang cha cũng có SP riêng không thuộc subcategory nào — đã lấy cả).
- **Tree Electric** — web chỉ có tiếng Trung; tên EN dịch theo file cũ, dạng `English (中文)`.

## Lưu ý
- **wenglor:** ví dụ "P1PC Series…" trong file .docx là trang *marketing* (`/s/…Portfolio`), không phải một cấp trong catalogue. Trong catalogue thật, mã `P1PC011`, `PNBC101`… nằm trực tiếp dưới `Laser Distance Sensors Triangulation` — đó chính là cấp ngay trên mã SP và file đã lấy đến đó.
- **Tree Electric:** web chỉ có 9 dòng sản phẩm, model nằm trong catalogue PDF nên không có cấp con.
- Balluff: 5 URL trong sitemap trả 404 (SP đã ngừng), đã bỏ qua.
