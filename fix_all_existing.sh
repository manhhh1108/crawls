#!/bin/bash
# Fix tat ca cac loi dich/format trong file xlsx da crawl - PHIEN BAN AN TOAN.
#
# CACH DUNG:
#   bash fix_all_existing.sh
#
# 3 step doc lap:
#   STEP 1: Replace cac translation sai (VI side - cot B,C)
#   STEP 2: Strip "Lưu trữ" / "Archive" prefix CHI cho file la Archive product
#           (targeted, khong dung global replace -> KHONG anh huong file khac)
#   STEP 3: Verify

set -e
cd "$(dirname "$0")"

echo "=========================================="
echo "STEP 1/3: Replace translation sai (VI side)"
echo "(Cac pattern duoi day chi xuat hien khi Google dich sai,"
echo " an toan replace globally trong cot B,C)"
echo "=========================================="
python fix_translations.py --vi-only \
  --replace "Title mô tả ngắn=>Mô tả ngắn của sản phẩm" \
  --replace "Sự miêu tả=>Mô tả" \
  --replace "Đánh máy.=>Điển hình" \
  --replace "Đánh máy=>Điển hình" \
  --replace "ID thiết bị THÁNG 12=>ID thiết bị DEC" \
  --replace "Phê duyệt cũ=>Chứng nhận phòng nổ" \
  --replace "Ex-approvals=>Chứng nhận phòng nổ" \
  --replace "Đặc sản=>Tính năng đặc biệt" \
  --replace "Bộ bán hàng=>Sales Kit" \
  --replace "Point-shaped=>Hình điểm"

echo ""
echo "=========================================="
echo "STEP 2/3: Strip 'Lưu trữ'/'Archive' prefix"
echo "(CHI fix file la Archive product - detect tu row 8 col B/D)"
echo "(Sua DUNG 6 cell cu the, KHONG dung global replace)"
echo "=========================================="
python fix_archive_prefix.py

echo ""
echo "=========================================="
echo "STEP 3/3: DONE! Verify by:"
echo "  python check_columns.py"
echo "  python find_bad_translations.py"
echo "=========================================="
echo ""
echo "[Optional] Cac script khac neu can:"
echo "  python add_columns.py    # Them row Danh muc + Thuong hieu (cho file thieu)"
echo "  python fix_merges.py     # Fix merge bi loi (cho file qua add_columns truoc do)"
