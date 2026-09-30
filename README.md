# crawls — thu thập dữ liệu sản phẩm 16 hãng + SICK

## Cấu trúc thư mục

```
crawls/
├─ shared/      Hạ tầng dùng chung cho mọi crawler
│    fetcher.py   HTTP qua curl (bypass TLS fingerprint) + cache đĩa
│    browser.py   Playwright: Browser (headed, cho web chặn bot) & Renderer (headless)
│
├─ sick/        Dự án crawl SICK (hãng 01) — nội dung + ảnh + dịch, đã xong
│
├─ catalog/     Crawl CÂY DANH MỤC 16 hãng (02–17) — đã xong
│    b_<hãng>.py     mỗi hãng 1 crawler, chạy độc lập: python b_festo.py
│    common.py       save/load JSON + helper đọc breadcrumb
│    out/            kết quả JSON từng hãng
│    build_excel.py  gộp out/*.json → data/Danh muc san pham 16 hang - FULL.xlsx
│    logs/           log các lần chạy
│
├─ content/     (SẮP TỚI) Crawl NỘI DUNG sản phẩm 16 hãng — xem content/README.md
│
├─ data/        Kết quả + tài liệu yêu cầu
│    Danh muc san pham 16 hang - FULL.xlsx   ← file danh mục mới nhất
│    Danh muc san pham 16 hang.xlsx          ← file cũ (đối chiếu)
│    README - crawl danh muc.md              ← tài liệu chi tiết crawl danh mục
│    Crawl nội dung Ntrend.docx              ← yêu cầu gốc
│
└─ .cache/      Cache HTTP/browser (gitignore) — xoá được, crawl lại sẽ chậm
```

## Chạy

```bash
# crawl lại 1 hãng (ví dụ)
cd catalog && python b_supmea.py

# dựng lại file Excel danh mục từ out/*.json
cd catalog && python build_excel.py
```

Yêu cầu: Python 3.10+, `openpyxl`, `playwright` (cho Festo/IFM/Pilz/Pepperl), `curl` trong PATH.

## Ghi chú nhanh theo hãng

- **Festo / IFM / Pilz** — web chặn bot, phải dùng `shared.browser.Browser` (Chrome thật, headed).
- **Pilz** — dùng mirror `www.pilz.com.cn/en-CN` (domain chính chặn hoàn toàn).
- **Supmea / Xingguang / OTENNLUX** — crawl bản EN (`en.supmea.com`, `en.xgcd.cn`, `en.otennlux.com`).
- **Tree Electric** — web chỉ có tiếng Trung, không có trang SP (model trong PDF); tên EN dịch tay dạng `English (中文)`.

Chi tiết đầy đủ: `data/README - crawl danh muc.md`.
