#!/bin/bash
# Re-crawl all 47k+ products with new format
# Run this in foreground (or with nohup for resilience)
# Log: crawl.log; image errors: image_log.csv

cd "$(dirname "$0")"
node crawl-checklist.js \
  --checklist SICK_Web_Products_Statistics.xlsx \
  --out-dir SICK_Products \
  --delay 500 \
  > crawl.log 2>&1
