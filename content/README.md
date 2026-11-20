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

## Ảnh đại diện

`brandimg.compose()` ghép ảnh SP vào khung NTREND 800×800 (nền trắng + logo
NTREND + hotline). Ảnh gốc nhỏ thì nó tự hạ khung xuống thay vì phóng to cho vỡ
(`MAX_UPSCALE`).

| Bước | Script | Việc |
|---|---|---|
| sau khi crawl | `fix_images_hd.py` | tìm bản HD của ảnh đã lấy |
| rồi | `recompose_images.py` | ghép lại từ bản nét nhất; **ảnh gốc < 200px thì bỏ hẳn** (`MIN_SRC`) vì phóng lên là vỡ nát |
| lúc build | `logo_fallback.py` | bài nào vẫn chưa có ảnh thì **lấy logo hãng** |

Bước cuối đã gắn sẵn vào đầu mọi `build_*_excel.py`, không phải chạy tay. Bài
dùng logo được đánh dấu `img_logo` trong `content.json` để phân biệt với ảnh
thật, và `recompose_images.py` bỏ qua các bài đó.

`logo_fallback.bases()` **phải né cả những tên ảnh thật đã có sẵn**. Crawler đặt
tên ảnh theo thứ tự *cào được*, còn `bases()` duyệt theo thứ tự *bài* — hai thứ
tự khác nhau thì một bài không có ảnh có thể trùng tên với bài có ảnh thật, và
logo **ghi đè lên ảnh sản phẩm**. Đã xảy ra: hai mục Pilz cùng tên "Education
Systems", mục có ảnh thật bị thay bằng logo Pilz. Cùng lý do, `--undo` xoá theo
đúng tên đã ghi trong bài, không tính lại tên.

**Hãng mới** cần một file logo trong `assets/<key>-logo.png`, lấy bằng:

```bash
python get_brand_logo.py <key> <trang chu hang>
python get_brand_logo.py ifm https://www.ifm.com/us/en --inplace   # logo chữ trắng
```

Phải dùng trình duyệt vì phần lớn logo là SVG. Chụp xong **mở ra xem**: script
chỉ lấy phần tử đầu tiên trông giống logo, có khi vớ phải icon. Thêm `<key>` vào
`BRANDS` trong `logo_fallback.py`.

## Nơi ghi file giao hàng

Hãng nào soát xong thì chuyển sang `data/File đã crawls/Noi dung <hãng>/`. Mọi
script đều hỏi `out_dir.brand_dir()` để biết ghi vào đâu: thư mục "đã check" có
sẵn thì ghi thẳng vào đó, không thì ghi vào `data/Noi dung <hãng>/`.

Đừng ghi cứng đường dẫn `data/Noi dung <hãng>` nữa: làm vậy sinh ra bản thứ hai
rồi hai bản lệch nhau — đã xảy ra với 5 hãng (IFM, Leuze, Pepperl+Fuchs, Perma,
wenglor), bản đã check thiếu ảnh logo mới thêm, bản mới thì thiếu sạch ảnh thật.

## OTENNLUX — bảng dọc và bố cục mô tả

**Bảng của hãng này hầu hết là bảng DỌC 2 cột** (`Item No. | OBS1-USB-01`,
`Brand | Otennlux`, …): 114/125 bảng. Cột trái là *tên thông số*, không phải mã
SP. `code_col()` trả về `-1` khi không tìm thấy cột mã, `kv_table()` nhận ra
dạng này; đừng để `code_col()` trả 0 làm mặc định — làm vậy thì cột nhãn bị coi
là cột mã, kéo theo hai lỗi: cột nhãn không được dịch (bảng nửa Anh nửa Việt) và
các nhãn chui vào thẻ tag thành `ItemNo., Brand, Materail, Voltage, …`. Mã SP
thật trong bảng dọc nằm ở *giá trị* của dòng có nhãn `Item No.`/`Model`.

**Mô tả dài phải giữ bố cục của trang hãng**: dòng tiêu đề `Details:` →
`Chi tiết:`, dấu `*` đầu mỗi gạch đầu dòng, và một dòng trống trước mỗi tiêu đề
(`bo_cuc()` trong builder chèn). Dịch thì gỡ dấu `*` ra, dịch xong gắn lại, và
ép dòng tiêu đề vẫn kết thúc bằng dấu hai chấm — máy dịch hay nuốt nó.

**Dòng đầu mô tả thường chỉ nhắc lại tên sản phẩm** (có khi thêm bớt vài chữ),
`nhac_lai_ten()` so theo tập từ nên bắt được cả bản thêm bớt. Tên đã nằm ở tiêu
đề bài rồi, để lại là thừa.

## Festo — hai chỗ dễ sập

**Chrome phải có cửa sổ.** festo.com dùng Akamai: headless bị trả `Access
Denied` ở trang chủ và 403 ở trang dòng SP. `Browser` tự thử lại tối đa 4 lần
khi Chrome mở lên rồi tắt ngay — lỗi chập chờn, không sửa thì cả lệnh chết từ
dòng 0.

**Cào theo lô, mỗi lô một tiến trình.** Đã gặp `MemoryError` ngay trong đường
ống nối sang Playwright khi máy còn ~1,4 GB RAM; lỗi đó làm đường ống chết hẳn,
tiến trình treo vĩnh viễn và thử lại bên trong không cứu được. Nên:

    python c_festo.py --crawl-only --batch 250   # cào 250 dòng tải thật rồi thoát
    python c_festo.py                            # pha dịch, chạy một mình
    python build_festo_excel.py

Dòng đã có cache đọc lại gần như không tốn gì nên lần chạy sau tiếp đúng chỗ.
Ở chế độ `--crawl-only` nó không gom dữ liệu vào bộ nhớ (pha đó ném đi ở cuối).

**Trang danh mục lọt vào urlmap.** 10 dòng trỏ vào `/c/...-id_pimNNN` thay vì
`/p/...`; trang `/c/` không có `_ROUTER_DATA` nên bị báo "nội dung không hợp lệ".
`festo_cat_expand.py` gọi API `/search/categories/<pimId>/products` để lấy dòng
SP con, thêm vào map và gỡ dòng danh mục đi. Chạy sau `cat_map.py festo`, trước
`c_festo.py`; chạy lại được nhiều lần. So trùng phải **bỏ phần `?categoryCode=`**
— một dòng SP nằm ở nhiều danh mục thì mỗi nơi cho một URL khác nhau ở đoạn đó.
Lần chạy 2026-10-28 cả 24 dòng con đều đã có trong map, nên chỉ gỡ 10 dòng
danh mục, không mất sản phẩm nào.

**Dòng SP hãng đã gỡ khỏi web** nằm trong `GO_KHOI_WEB` của `c_festo.py` (mở ra
thì chuyển về `/error/404`). Không ghi ra đó thì mỗi lần chạy đều báo lỗi và lần
sau lại mất công dò lại.

## Supmea — bố cục trang không cố định

**Không cắt theo "văn trước, bảng sau".** Gần một nửa số trang để bảng
`Specification` **lên trước** rồi mới tới `Introduction`, và có trang phần
Introduction chỉ toàn ảnh chụp (Coriolis mass flowmeter). `sections()` cắt theo
*nhãn khối*, mà nhãn khối thì mỗi trang một thẻ: có trang `<h3>`, có trang
`<ul><li><p>Specification</p></li></ul>`. Chỉ coi là nhãn khi **cả dòng** khớp
danh sách từ — hãng bọc luôn danh sách đặc điểm trong cùng thẻ `<ul>`, nhận nhầm
là mất hẳn đoạn đó.

**Bảng thông số có khi 3 cột.** 14 trang thêm một cột *nhóm* bên trái
(`System | working power | (220±22)VAC`). Vẫn là bảng thông số, chỉ khác là tên
thông số gồm hai cấp. Phân biệt với bảng tra cứu 3 cột thật
(`Degree no | Signal type | Range`) bằng dấu vết của ô bị gộp: bảng thông số có
cột nhóm lặp lại hoặc có dòng bị `colspan` kéo ngang. Bỏ qua chỗ này thì 14 bài
ra "không có bảng thông số" trong khi dữ liệu nằm nguyên ở khối bảng tra cứu
phía dưới.

**Ba dòng trỏ vào trang danh mục rỗng** (`BO_TRANG_DANH_MUC` trong
`c_supmea.py`) — giống hệt 7 dòng của OTENNLUX: trang có thật, có `h1`, nhưng
không có thẻ sản phẩm nào. Để lại thì ra ba bài rỗng.

**Vài trang hãng đã gỡ** (404 khi tải mới) nhưng nội dung vẫn còn trong
`.cache/http` từ hồi dựng danh mục, nên bài vẫn đầy đủ. Đừng xoá cache của hãng
này trước khi giao hàng.

## Schneider — danh sách mã nằm sau API, phân trang kiểu Endeca

`se.com` **không** chặn curl, nhưng trang dòng SP là một ứng dụng Javascript:
danh sách mã sản phẩm không có trong HTML. Mở tab *Products* bằng Chrome thật và
ghi lại request thì ra hai API JSON công khai, không cần khoá:

```
/ranges/<id>/products?brand=se&country-code=ca&language-code=en&Nrpp=1000
/products-card/main?brand=se&country-code=ca&language-code=en&ids=A,B,C
```

**Phân trang là `No` / `Nrpp` (Endeca), viết hoa đúng như vậy.** Tất cả tên quen
thuộc — `page`, `pageNumber`, `offset`, `limit`, `pageSize`, `recsPerPage`,
`first-rec-num` — đều bị bỏ qua lặng lẽ: API vẫn trả đúng 12 mã đầu và
`firstRecNum: 1`, không báo lỗi gì. Thử lọc theo `node-id` chỉ gom được 48/132
mã. `Nrpp=1000` lấy được cả danh sách trong một lần gọi.

`/products-card/main` chỉ nhận khoảng **80 mã một lần** (URL dài hơn thì trả về
rỗng, không báo lỗi) → gọi theo lô 40.

**Phần Overview thì ngược lại**: nằm sẵn trong HTML, nhưng trong một khối
Javascript của builder.io chứ không phải trong thẻ HTML — `main:{id:"...",
name:"...", description:"...", longDescription:"...", image:{url:"..."}}` và
`presentation:{presentation1:"<p>..."}`. Phải cắt lấy đúng khối
`main:{id:"<mã dòng>"` trước khi đọc: trang có hàng chục chỗ khác cũng dùng khoá
`description`.

**Nội dung tab Overview nằm trong trang HAI BẢN, cả hai đều thiếu một nửa**
(review khách, 2026-11-02 — `presentation1` chỉ có ở trang kiểu cũ, 128/250):

* bản **cây builder.io** bị rút gọn: chuỗi nào dùng hai lần bị thay bằng biến
  (`description:v`, `title:u`) và biến đó **không được định nghĩa trong HTML**
* bản **mảng dữ liệu SvelteKit** thì đủ chữ nhưng xếp theo *vị trí*, không có
  tên trường

Nên `overview_of()` quét **chuỗi Javascript ở cả hai bản** rồi bỏ trùng theo nội
dung. Phải đọc đúng kiểu chuỗi Javascript (`\"` và `\u003C`); **đừng** thay thô
`s.replace('\\"', '"')` cho cả trang — làm vậy thì dấu nháy *bên trong* chuỗi
thành dấu nháy thật, không còn biết chuỗi kết thúc ở đâu, và đúng những đoạn có
`<strong class="accent">` (tức gần hết phần Overview) bị mất.

Thứ tự lấy mô tả dài khách chốt: **thân khối "Features" → thân khối
"Discover …" → đoạn dài đầu tiên trong vùng Overview → blurb cạnh ảnh sản
phẩm** (`longDescription` rồi `description`). Phải tính vùng Overview bắt đầu từ
tiêu đề đầu tiên khớp "Discover/Features/Benefits of/Facts about" — nếu không,
đoạn dài đầu trang lại là lời mời *đăng ký nhận tin* của khung trang. Vài trang
xếp khối "Delivering value…" **trước** đoạn giới thiệu nên không thể chỉ lấy
đoạn dài đầu tiên.

**Ảnh: `rendition_<n>` trong URL chính là số pixel.** Dữ liệu sản phẩm trả về
`rendition_369_jpg`, có chỗ còn `rendition_113_png` — 113 px phóng lên khung 800
thì vỡ nhoè nên `brandimg` tự hạ khung xuống 400 (đúng lỗi "ảnh mờ và nhỏ").
**`rendition_1500_jpg` có thật** và nét: thử `1500 → 369 → URL gốc`, cỡ nào
không có thì máy chủ trả ảnh mặc định ~1,3 KB nên lọc theo dung lượng. Sau khi
đổi: 675/675 ảnh gốc là 1500×1500, 759/759 khung giao đủ 800×800.

**Khối Features phải đọc theo CÂY COMPONENT, không theo từng đoạn văn**
(review khách 04/11). Trang se.com là một chuỗi component
`{name:…,tag:"…",options:{…}}`; tiêu đề mỗi mục là component
`heading-section-tittle-h1-h2-h3`, nên **khối Features = mọi component nằm giữa
tiêu đề "Features" và tiêu đề kế tiếp**. Hai trường chứa chữ:

| tag | trường | ghi chú |
|---|---|---|
| `content-text-with-left-title` | `rightColumn` | khối văn một mảng |
| `content-accordion-container` | `itemDescription` | **mỗi panel một cái** |

Đọc theo đoạn văn (cách cũ, `overview_of()`) hỏng hai kiểu: accordion của
**TeSys island** có 6 panel mà chỉ lấy được panel đầu (khách đếm đúng 6), còn
đoạn SEO nằm ngay *sau* khối Features thì dễ bị vơ nhầm vào vì không biết khối
kết thúc ở đâu. Trang kiểu cũ không có cây component này → `khoi_features()`
trả rỗng và chuỗi ưu tiên cũ chạy tiếp.

Ba cái chặn khác cũng cắt cụt mô tả dài, đã nới:

* `MAX_HL` 12 → **60**. Khối Features của **TeSys LR9** dài 25 dòng (ba dòng SP
  LR9D01-110S / LR9D / LR9F), để 12 là mất hơn nửa.
* ngưỡng **12 ký tự** → ít nhất **hai ký tự chữ/số**: tiêu đề dòng SP `LR9D`,
  `LR9F` chỉ 4 ký tự nên trước đây bị bỏ.
* **bỏ trùng trên cả khối** → chỉ bỏ dòng trùng **liền kề**: dòng "Three-phase
  and single-phase applications" lặp ở cả ba dòng SP là có thật.

Văn bản hãng đã có sẵn dấu `•` ở các ý con, nên `build_schneider_excel.gach()`
thay dấu đó bằng `-` thay vì cộng thêm, kẻo ra `- • 5 current ranges…`.

**Ảnh: nhiều trang ghi `image:{url:a,…}` — `a` là BIẾN.** Cây builder.io bị rút
gọn nên regex chuỗi không thấy gì, và bài rơi xuống logo hãng dù trang có ảnh
(review 04/11). Thẻ `og:image` ở đầu trang luôn ghi đủ URL đó: 82/84 bài từng
dùng logo thực ra có ảnh thật.

**Tên dòng SP có thẻ HTML tô màu bên trong.** `name` và `description` trong khối
`main` không phải chữ trơn: `Power<span class="accent">PacT</span>` và
`<span class="highlight-naming">…</span>` (285/759 bài). Bỏ thẻ bằng cách thay
bằng *khoảng trắng* thì ra "Power PacT" — sai tên; phải nối liền, nên có
`plain()` riêng cạnh `txt()`. Để lọt thì tên bài, `name_vi` **và cả tên file
ảnh** (slug hoá từ tên) đều mang nguyên đoạn HTML.

**`presentation1` viết HTML kiểu cũ, thẻ `<LI>` không đóng**
(`<UL><LI>a <LI>b <LI>c</LI></UL>`). Cắt ở `</li>` thôi thì ba ý dính thành một
dòng dài — phải cắt ở cả thẻ mở.

**Nặng đĩa**: mỗi trang dòng SP ~1,5 MB, 759 trang → cache HTTP phồng thêm vài
trăm MB. Đừng xoá rồi chạy lại nếu không cần.

## Pilz — phân trang và khối bộ lọc

Mỗi dòng trong danh mục Ntrend là **một mục eshop** (`/c/<mã>`), tức một nhóm
mã SP. Trang mục cho sẵn đúng ba cột của template: `product__list--name`
(Model), `product__code` (Mã đặt hàng), `product-keywords` (Mô tả).

Trang chỉ hiện **20 mã một lần**, phân trang bằng `?q=%3Aproductrelevance&page=N`.
Tham số `pageSize` bị bỏ qua. 31/180 mục có hơn 20 mã, mục lớn nhất 244 mã.
Thỉnh thoảng một trang trả về `HTTP 000` (curl đứt kết nối) — **không được dừng
vòng lặp** ở đó, các trang sau vẫn lấy được; `c_pilz.py` ghi lại trang hỏng và
in ra "THIEU n TRANG".

**Phải gộp cả hai bản của trang mục.** Trang mở trần (không có `?q=`) sắp xếp
khác với các trang phân trang: ở mục "Contact expansions" hai bản chỉ trùng
10/20 mã. Chỉ lấy một bên thì bên kia sót mã — đã thử lấy riêng bản `?q=` và
*mất* mã (PSENmlock mini 20→16, một mục về 0). Nên `c_pilz.py` lấy **trang trần
+ từng trang `?q=`** rồi bỏ trùng theo mã: 2.344 → **2.395 mã**. Trang `?q=`
của `page=0` có thể rỗng trong khi trang trần vẫn có mã, nên gặp rỗng ở trang
đầu thì `continue`, đừng `break`.

**Lệch với con số hãng đăng không có nghĩa là cào sót.** Còn 7/180 mục ít hơn số
hãng ghi, vì **hãng tự đăng trùng mã**: mục "Accessories PSSuniversal" đếm 27
sản phẩm nhưng có 2 dòng trùng y hệt cả mã lẫn tên (313111, 313116) — bỏ trùng
còn 25; mục "Pre-assembled cable" 264 dòng thô chỉ gồm 231 mã riêng. Một mục
(`PNOZelog > Monitoring of E-STOP…`) nay hãng đã gỡ hết sản phẩm, trả về 0 dù
danh mục ghi 26.

Trang mục **không có đoạn văn giới thiệu nào** (`meta description` rỗng, không có
khối mô tả danh mục). Mô tả dài lấy từ **khối bộ lọc bên trái**: mỗi bộ lọc là
một thông số của cả dòng kèm các giá trị có thật trong mục đó — giống cách đã
làm cho Festo (facets). Template Pilz ghi "(Loại các thông số trùng tên dòng)"
nên bỏ bộ lọc nào chỉ nói lại chính tên mục.

Ảnh trên `pilz.com.cn` **phải tải bằng curl**: `urllib` bị trả về `HTTP 567`.

**Mô tả ngắn phải đủ cả hai câu của mẫu** (review khách, 2026-11-02):
"[Loại SP] [Tên dòng SP] là sản phẩm trong công nghiệp thuộc hãng [Thương hiệu],
… [Loại SP] [Tên dòng SP] giúp đáp ứng …". Trang Pilz không có đoạn giới thiệu
nào nên mô tả ngắn hoàn toàn là mẫu này — cắt còn câu đầu (`cau_dau`) là mất
nửa thông tin. Tên dòng SP: 111/180 mục đã có sẵn mã dòng trong tên; 55 mục lấy
mã từ **mục cha** ("PNOZsigma - safety relays (77)" → `PNOZsigma`). Chỉ nhận cha
trông như mã dòng — cha kiểu "Cables and plug-in connectors" chỉ là một tên loại
nữa, ghép vào thành câu nửa Anh nửa Việt mà không thêm thông tin. 14 mục còn lại
hãng thật sự không có tên dòng.

**Đừng hạ thường cả tên.** `en.lower()` biến "PSSuniversal - base modules" thành
"pssuniversal…", còn luật cũ của cột tiếng Việt (`nv.isupper()`) hạ "PSSuniversal"
thành "pSSuniversal" ngay ở Meta description. Hai hàm `thuong()` / `type_l` chỉ
hạ thường từ viết hoa chữ đầu, giữ nguyên từ có chữ hoa bên trong hoặc có số.

## Khung ảnh luôn 800×800

`brandimg` từng tự **thu nhỏ khung** (xuống tận 400) khi ảnh gốc quá bé, để
không phóng quá 1,4 lần cho khỏi vỡ. Khách không chịu: *"có sản phẩm thì lại bị
nhỏ quá 400x400"* (review Pilz 04/11) — bài đăng lên web cần ảnh đều nhau. Nay
mọi khung đều 800×800, ảnh gốc bé thì phóng bằng Lanczos.

Đây là giới hạn của **nguồn**, không phải crawl sót: ảnh Pilz chỉ 188–300 px và
trang hãng không đăng bản nào to hơn (quét 784 trang đã cache chỉ thấy đúng một
cỡ `300Wx300H`, không có `srcset`).

## Pilz — mục "0 Products found" vẫn có ảnh banner

Ảnh bài lấy từ ảnh thu nhỏ của mã SP đầu tiên trong mục. Mục không có mã nào
thì không có thẻ sản phẩm, nhưng đầu trang vẫn có ảnh banner
(`<img class="category-image">`) — 10/18 mục đang dùng logo thực ra có ảnh này
(review 04/11).

**Ảnh banner nằm dưới `/mam/` thì curl không tải được.** Mọi request tới
`/mam/…` (và tới `/en-CN/products/…`) trả về một trang 639–657 byte đặt cookie
`__wc_jsct` bằng Javascript rồi reload — ảnh dưới `/medias/` thì tải bình
thường. `co_anh()` kiểm tra file tải về có phải ảnh thật không rồi **xoá** file
hỏng, để lần sau còn thử lại và không lấy rác làm ảnh gốc.

Mấy ảnh này lấy bằng **`pilz_banner_img.py`** — script chạy tay riêng, mở
`shared.browser.Browser` (Chromium thật) cho Javascript đặt cookie xong rồi tải
ảnh trong cùng phiên đó. Không gộp vào `c_pilz.py` vì mỗi lần cào lại sẽ phải
bật Chromium cho đúng ~10 ảnh, trong khi 162 ảnh còn lại curl tải thừa. Script
giữ nguyên tên file ảnh cũ nên ảnh thật **ghi đè** lên logo, không để lại file
rác trong `images/`. Chạy sau `c_pilz.py`, rồi chạy lại `build_pilz_excel.py`.

Kết quả (05/11): 10 mục đã có ảnh banner thật, **8 mục còn lại vẫn dùng logo**
vì trang hãng không có thẻ `category-image` nào. Hai mục *PES - Actuator Board
Contactor* và *Education Systems* dùng chung một ảnh vì hãng đăng cùng một URL
banner.

## Xingguang / Starshine — cả trang là một đoạn văn

`en.xgcd.cn` do một nền tảng website Trung Quốc sinh ra, lớp thẻ đều nhau. Phần
`Product Description` **không có bảng nào** — cả trang chỉ là một đoạn ngăn cách
bằng `<br>`, gồm ba kiểu dòng trộn lẫn:

| dòng | nghĩa | đi về đâu |
|---|---|---|
| `Output speed: 7.15rpm-155.56rpm` | thông số | bảng Thông số kỹ thuật |
| `Characteristic:` | nhãn khối | bỏ |
| `1. ...` / `● ...` | đặc điểm | Mô tả dài + Mô tả ngắn |

Riêng `Input:` để **giá trị trống** rồi liệt kê từng kiểu đầu vào ở các dòng sau
— phải gom các dòng không phải gạch đầu dòng ngay sau nó lại làm giá trị.

**Mô tả phải giữ đúng thứ tự trên trang** (review khách 02/11). `doc_desc()`
từng trả về hai danh sách rời (văn xuôi, đặc điểm) rồi người gọi nối
`prose + feats` — thế là đoạn kết "Therefore, it can be used as…" của trang SWB
bị đẩy lên **trước** các gạch đầu dòng, đọc ra ngược hẳn. Nay hàm trả về một
`body` xếp theo thứ tự trang, mỗi dòng là `[có_phải_gạch_đầu_dòng, chữ]`, và
`long_bul` trong `content.json` giữ cờ đó để bài in ra `- …` đúng chỗ. Các khối
còn lại được lợi theo: đoạn giới thiệu của SNFEG3000/SMQ340/SVF về lại đầu bài,
và ba tiêu đề phụ của SKC87 ("Leak proof - sealed"…) nằm đúng trên nhóm gạch
đầu dòng của nó thay vì dồn lên trên.

Trang nào **không đánh số** `1. 2.` thì cứ mỗi đoạn thành một gạch đầu dòng cho
dễ đọc. Cả 16 trang hiện tại đều có đánh số nên nhánh này chưa chạy.

Cuối khối mô tả vài trang còn **dán lại danh sách từ khoá** ("Line series
reducer; reducer" ở trang SHMF) — không phải câu, không có dấu chấm, các vế
ngăn bằng dấu chấm phẩy. Chỉ loại ở **dòng cuối cùng**, nếu không thì cắt nhầm
tiêu đề phụ giữa khối.

**Bảng thông số của hãng chỉ có ở dạng ảnh.** Trang nào cũng có đúng một khối
`reset_style` không chữ, chỉ chứa ảnh bảng (~500 px) — khách xin tải về kèm bài
(review 02/11). `grab_spec_img()` tải **nguyên bản**, không ghép khung NTREND:
ghép vào khung 800 thì bảng bị thu nhỏ lại, chữ trong ô khó đọc. File đặt tên
`<slug>-thong-so.jpg` cạnh ảnh SP, vừa ghi tên vào ô "Ảnh bảng thông số" vừa
nhúng thẳng xuống cuối bài để soát cho nhanh. CDN `thefastimg.com` **bỏ qua**
tham số `?x-oss-process=image/resize` nên 500 px là bản gốc, không có bản to hơn.

Ô "Bảng mã sản phẩm" trong template **để trống**, và đúng là hãng không đăng mã
đặt hàng nào: mỗi trang là một dòng SP, mã dòng nằm ngay trong tên bài (SVF,
NCJ, SHMF…) nên lấy làm thẻ tag.

## Tree Electric — chín dòng SP nằm trên một trang

`?index=N` chỉ là để Javascript hiện khối nào, **không đổi nội dung tải về**:
một lần tải `products.html` là có cả 9 khối. Ghép khối với dòng danh mục bằng
**tên tiếng Trung**, không bằng số `index` — số `list_pdfN.pdf` không khớp số
khối (khối 4 dùng `list_pdf5.pdf`, khối 5 dùng `list_pdf10.pdf`).

Mục "工业交换机系列" (Industrial Switches) bị `<!-- -->` bọc ở **thanh bên**
nhưng khối nội dung vẫn còn trong trang, và danh mục Ntrend có dòng đó — nên
không được bỏ phần nằm trong comment.

Ảnh duy nhất trong mỗi khối là **ảnh bìa catalogue**: một tờ A4 900×1143
(riêng `list11.png` là 1240×1754). Nửa trên là khối ảnh (ảnh ngành + ảnh sản
phẩm trên nền xanh), nửa dưới là tên mục bằng chữ Trung, logo hãng và rất nhiều
khoảng trắng. Khách khoanh đúng **nửa trên** và yêu cầu dùng nó làm ảnh bài
(review 02/11: "hiện tại đang là ảnh logo"), nên `cat_trang()` tìm dải khoảng
trắng đầu tiên sau 1/6 chiều cao rồi cắt tại đó — cả 9 tờ đều cắt ở 57-60% chiều
cao — xong mới ghép khung NTREND. Link ảnh bìa vẫn ghi xuống cuối bài ở dòng
"Ảnh bìa catalogue (tham khảo)".

Hai chỗ nữa review bắt được, cùng sửa: **mô tả ngắn phải đủ cả hai câu** của mẫu
`DESCDESC` (trang hãng không có chữ nào khác bù vào nửa câu bị cắt), và ô **Mô tả
thông số** kết bằng *"Dưới đây là **datasheet** chi tiết của sản phẩm:"* chứ
không phải "thông tin chi tiết" — thứ nằm ngay dưới ô đó là link catalogue PDF,
hãng này không có bảng thông số nào.

## Đừng chạy hai crawler cùng lúc

`translate.py` dùng **một file cache duy nhất**, nạp lúc import và ghi đè cả file
mỗi lần `_flush()`. Chạy song song hai hãng thì tiến trình nào flush sau sẽ xoá
sạch bản dịch của tiến trình kia.

## Dịch — vì sao không dùng thẳng Google

Google dịch thẳng sai rất nhiều thuật ngữ ngành (`Housing` → "Nhà ở",
`Bus termination` → "Chấm dứt xe buýt", `Art. 3` → "Nghệ thuật 3"). `vi_engine`
chạy 3 lớp:

1. **Tra từ điển trước** (`vi_terms.py`, 930 cặp) — khớp nguyên chuỗi thì dùng
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

**Muốn sửa 1 thuật ngữ dịch sai:** thêm cặp vào `data/Tu dien dich EN-VI.xlsx`
(ưu tiên cao nhất, đè lên mọi nguồn khác) hoặc vào `EXTRA` trong `vi_terms.py`.

### Áp từ điển vào dữ liệu ĐÃ crawl — `apply_tudien.py`

Cào lại **không** đủ để áp từ mới: cache dịch đánh theo sha1 của chuỗi tiếng
Anh, nên phần lớn câu sẽ trả về y nguyên bản cũ — chưa kể cào lại 17 hãng mất
nhiều giờ. `apply_tudien.py` sửa thẳng trên `content.json` (có `.bak`), sau đó
chỉ cần chạy lại `build_<hãng>_excel.py`.

Mỗi luật có **hai tầng**, và tầng hai là bắt buộc:

* `toan_cuc` — cụm chắc chắn đúng nghĩa, thay ở mọi trường `*_vi`.
  Ví dụ "Cảm biến chạc" chắc chắn là *fork sensor*.
* `cap` + `en_bai` — từ dễ nhầm, chỉ thay khi **câu tiếng Anh gốc** (và, nếu cần,
  cả bài) có đúng từ khoá đó.

Không có tầng hai thì hỏng ngay — cả ba trường hợp này đều là thật:

| chuỗi | tiếng Anh gốc | nếu thay bừa |
|---|---|---|
| "Mặt bích chạc" (Festo) | *clevis flange* | → "Mặt bích kìm", sai |
| "Chạc gắn mũ bảo hộ" (P+F) | *helmet mount attachement forks* | → "Kìm gắn…", sai |
| "thiết kế chữ U" (P+F) | *U design* | đã dịch đúng, đừng động vào |

Nên `chạc` trần chỉ đổi khi **cả bài** nói về *fork sensor*, và `chữ U` trần
thì không đổi gì cả.

Một lưu ý nữa: khách chỉ sửa dòng *moderate corrosion stress*, nhưng cả họ CRC
của Festo nằm **chung một ô** ("0 - …; 1 - …; 2 - …"), để lẫn lộn hai cách gọi
thì ô đó đọc rất kỳ, nên đổi đồng bộ cả họ sang "mức độ ăn mòn".

Một từ tiếng Anh thường đã bị dịch ra **nhiều kiểu khác nhau**, nên sửa một kiểu
là chưa xong (review 20/11). *fork* ra cả bốn: "cảm biến chạc", "cảm biến
dạng chữ U", "cảm biến dạng fork" và "cảm biến quang **dạng khe**". Cách
làm đúng là chạy `dd_quet` kiểu đối chiếu: duyệt mọi cặp (EN, VI), lọc câu nào
tiếng Anh **có** từ khoá mà tiếng Việt **chưa có** bản dịch mong muốn, rồi xem
tận mắt các kiểu dịch đang dùng.

Nhưng cẩn thận: kiểu đối chiếu này **báo động giả rất nhiều**. `NO` khớp
"**No**. of pins"; `flush` khớp "non-flush" (mà "lắp nổi" là đúng); `Closed` khớp
mọi chữ "closed" giữa câu. Phải xem mẫu rồi mới viết luật, đừng sửa hàng loạt
theo con số.

Ba cặp dễ lẫn đã gây sự thật, để đây cho lần sau:

| tiếng Việt | có thể là | phân biệt bằng |
|---|---|---|
| "dạng khe" | *fork* hoặc **slot** | chặn ngược chữ `slot` |
| "lưới quang" | *light curtain* hoặc **light grid** | chặn ngược `light grid` |
| "chạc" | *fork sensor* hoặc **clevis** | chặn theo cả bài |

"Rèm quang" thì ngược lại: nó **luôn** là light curtain, nên không được chặn
ngược `light grid` — nhiều dòng của Pepperl+Fuchs nói **cả hai** mặt hàng trong
một câu ("for safety light grids SLP, for safety light curtains SLC").

**Nhãn tiếng Anh trong `items[].attrs_vi` thì ĐừNG đổi.** Có 18.354 ô như vậy
(Leuze, wenglor, Reer) — nhìn thì tưởng sót dịch, nhưng đó là **khoá nội bộ** để
khớp với `cols_en`; bản hiển ra Excel lấy từ `cols_vi` và đã tiếng Việt đầy
đủ. Đổi khoá là vỡ bảng thông số.

Tương tự, cột **"Tiếng Anh (soát chính tả)"** trong file giao giữ nguyên tiếng
Anh là đúng — đừng đếm nó là lỗi dịch sót.

Hai cái bẫy khi duyệt `content.json` để sửa hàng loạt — cả hai đều đã làm
sót thật trong đợt 17/11:

1. **Trường `*_vi` có thể lồng NHIỀU TẦNG.** Bảng thông số của Reer
   (`items[].attrs_vi`), SICK và OTENNLUX lưu dạng `[[tên, giá trị], ...]`.
   Duyệt kiểu "list các chuỗi" là bỏ qua **toàn bộ** bảng thông số.
2. **Nguồn tiếng Anh của `name_vi` thường là `title_en`, không phải `name`.**
   IFM có `name` = "AC sensors" nhưng `title_en` = "Inductive sensors with
   combined AC/DC output stage"; chỉ tra `name` thì không thấy chữ *inductive*
   đâu và bài đó bị bỏ qua.

Từ điển cung cấp hai kiểu khớp, cần nắm để đừng đặt nhầm kỳ vọng:
**cụm từ** (có dấu cách, từ 7 ký tự) thì thay được **giữa câu**; **từ đơn**
(`Closed`, `cell`, `fork`) chỉ dùng khi **nguyên ô đúng bằng từ đó** — xem
`vi_terms.phrases()`.

## Lưu ý đã khảo sát trước (2026-09-30)

| Hãng | Tình trạng nội dung trên web |
|---|---|
| Festo, IFM, Schneider, Pepperl, Pilz, Leuze, Balluff, Reer, wenglor, Mitsubishi, Perma | Đầy đủ — crawl bình thường |
| Supmea EN, ESPE, OTENNLUX EN | Đủ ảnh + text |
| **Xingguang** | Ảnh SP có, thông số là **ảnh chụp bảng** → đã tải kèm bài (`<slug>-thong-so.jpg`), muốn ra chữ thì còn phải OCR hoặc nhập tay |
| **Tree Electric** | **Không có trang SP** — model nằm trong catalogue PDF, phải trích từ PDF |

- Ảnh trên nhiều web là lazy-load (`data-src`, `lazy=`…) — đừng chỉ lấy `src`.
- Code SICK đời cũ (crawl theo từng SKU) vẫn giữ trong `sick/` để tham khảo.
