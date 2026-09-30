# content/ — crawl nội dung sản phẩm 16 hãng (chưa bắt đầu)

Giai đoạn tiếp theo: lấy nội dung chi tiết (mô tả, thông số, ảnh, tài liệu) cho từng
sản phẩm/danh mục, đầu vào là cây danh mục ở `catalog/out/*.json` (mỗi node lá có URL).

## Khung dự kiến

```
content/
├─ c_<hãng>.py    crawler nội dung từng hãng (dùng shared.fetcher / shared.browser)
└─ out/<hãng>/    JSON + ảnh tải về
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
