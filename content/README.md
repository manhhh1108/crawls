# content/ — crawl nội dung bài đăng web NTREND (mỗi dòng SP = 1 bài)

Mỗi bài = 1 dòng sản phẩm (node đánh dấu `x` trong cây danh mục `catalog/out/*.json`).
Nội dung theo template `data/ntrend template.xlsx`, dựng 2 bản EN + VI.

## Cấu trúc

```
content/
├─ translate.py            EN→VI qua translate.googleapis.com, cache .cache/translate_vi.json
├─ brandimg.py             ghép ảnh SP: nền trắng + logo NTREND + hotline (theo mẫu)
├─ assets/ntrend-logo.png  logo gốc
├─ c_<hãng>.py             crawler nội dung từng hãng → out/<hãng>/content.json
├─ build_content_excel.py  content.json → data/Noi dung san pham - <hãng>.xlsx
└─ out/<hãng>/             content.json + src_img/ (ảnh gốc) + images/ (ảnh đã ghép)
```

## Chạy (ví dụ Perma)

```bash
cd content && python c_perma.py && python build_content_excel.py
```

## Lưu ý đã khảo sát trước (2026-09-30)

| Hãng | Tình trạng nội dung trên web |
|---|---|
| Festo, IFM, Schneider, Pepperl, Pilz, Leuze, Balluff, Reer, wenglor, Mitsubishi, Perma | Đầy đủ — crawl bình thường |
| Supmea EN, ESPE, OTENNLUX EN | Đủ ảnh + text |
| **Xingguang** | Ảnh SP có, nhưng thông số là **ảnh chụp bảng** → cần OCR hoặc nhập tay |
| **Tree Electric** | **Không có trang SP** — model nằm trong catalogue PDF, phải trích từ PDF |

- Ảnh trên nhiều web là lazy-load (`data-src`, `lazy=`…) — đừng chỉ lấy `src`.
- SICK (hãng 01) đã crawl xong nội dung từ trước — code tham khảo trong `sick/`.
