#!/usr/bin/env node
const { resolve } = require("node:path");
const {
  parseArgs,
  fetchJson,
  fetchCategoryProductIds,
} = require("./crawl-sick-product");

// Reuse readChecklistSeries from crawl-checklist logic
const { execFileSync } = require("node:child_process");
const { writeFileSync } = require("node:fs");

function readChecklistSeries(checklistPath) {
  const pythonCode = `
import json, sys
from openpyxl import load_workbook

wb = load_workbook(sys.argv[1], data_only=True)
ws = wb.active

rows = []
for row_idx, row in enumerate(ws.iter_rows(values_only=True), 1):
    if row_idx <= 2:
        continue
    cells = list(row)
    if len(cells) < 2:
        continue
    stt_raw = cells[0]
    series_raw = cells[1]
    count_raw = cells[3] if len(cells) > 3 else None

    if stt_raw is None or series_raw is None:
        continue

    stt = str(stt_raw).strip()
    series = str(series_raw).strip()

    if not stt or not series:
        continue

    stt_clean = stt.replace(',', '.')
    try:
        stt_num = float(stt_clean)
    except ValueError:
        continue

    # x.0 rows are category headers
    if stt_num == int(stt_num):
        # This is a category header - extract category name
        rows.append({
            "row": row_idx,
            "stt": stt,
            "series": series,
            "expected": 0,
            "isCategory": True
        })
        continue

    expected = 0
    if count_raw is not None:
        try:
            expected = int(float(str(count_raw)))
        except (ValueError, TypeError):
            expected = 0

    rows.append({
        "row": row_idx,
        "stt": stt,
        "series": series,
        "expected": expected,
        "isCategory": False
    })

print(json.dumps(rows, ensure_ascii=False))
`;

  const pythonCmd = process.platform === "win32" ? "python" : "python3";
  const tmpScript = resolve(checklistPath + ".count_tmp.py");
  writeFileSync(tmpScript, pythonCode, "utf8");
  try {
    const out = execFileSync(pythonCmd, [tmpScript, checklistPath], {
      encoding: "utf8",
      timeout: 30000,
    });
    try { require("node:fs").unlinkSync(tmpScript); } catch { }
    const lastLine = out.trim().split("\n").pop();
    return JSON.parse(lastLine);
  } catch (err) {
    try { require("node:fs").unlinkSync(tmpScript); } catch { }
    throw new Error(`Khong doc duoc checklist: ${err.message}`);
  }
}

async function findFamilyCode(seriesName) {
  const searchUrl = `https://www.sick.com/api/fact-finder/search/enSG-sick?query=${encodeURIComponent(seriesName)}&filter=DefType:ProductVariant&page=1&showPermutedSearchParams=false`;
  let data;
  try {
    data = await fetchJson(searchUrl);
  } catch {
    return { familyName: null, familyCode: null, method: "ERROR" };
  }

  const hits = data?.hits ?? [];
  if (hits.length === 0) {
    return { familyName: null, familyCode: null, method: "NO_HITS" };
  }

  const families = new Map();
  for (const h of hits) {
    const mv = h?.masterValues ?? {};
    const fn = mv.ProductFamilyName || "";
    const fc = mv.ProductFamilyCode || "";
    if (fn && fc) {
      if (!families.has(fn)) families.set(fn, { code: fc, count: 0 });
      families.get(fn).count += 1;
    }
  }

  const nameLower = seriesName.trim().toLowerCase();

  for (const [fn, info] of families) {
    if (fn.toLowerCase() === nameLower)
      return { familyName: fn, familyCode: info.code, method: "EXACT" };
  }
  for (const [fn, info] of families) {
    if (nameLower.includes(fn.toLowerCase()) || fn.toLowerCase().includes(nameLower))
      return { familyName: fn, familyCode: info.code, method: "PARTIAL" };
  }
  if (families.size > 0) {
    let best = null, bestCount = 0;
    for (const [fn, info] of families) {
      if (info.count > bestCount) { best = fn; bestCount = info.count; }
    }
    if (best) return { familyName: best, familyCode: families.get(best).code, method: "GUESS" };
  }
  return { familyName: null, familyCode: null, method: "NO_FAMILY" };
}

async function countProductsInCategory(familyCode) {
  const code = String(familyCode || "").replace(/[^a-z0-9]/gi, "").toLowerCase();
  if (!code) return 0;

  // Just get page 1 to read totalHits
  const url = `https://www.sick.com/api/fact-finder/search/enSG-sick?query=*&filter=DefType:ProductVariant&filter=CategoryCode:${code}&page=1&showPermutedSearchParams=false`;
  try {
    const data = await fetchJson(url);
    return data?.totalHits ?? (data?.hits?.length ?? 0);
  } catch {
    return 0;
  }
}

async function main() {
  const checklistPath = resolve(process.argv[2] || "AUMI_SP_SICK_Checklist.xlsx");
  process.stdout.write(`Doc checklist: ${checklistPath}\n\n`);

  const allRows = readChecklistSeries(checklistPath);
  const categories = allRows.filter(r => r.isCategory);
  const series = allRows.filter(r => !r.isCategory);

  process.stdout.write(`Tim thay ${categories.length} nhom danh muc, ${series.length} dong san pham.\n`);
  process.stdout.write(`Dang truy van SICK API de dem so san pham thuc te...\n\n`);

  let currentCategory = "";
  let categoryActual = 0;
  let categoryExpected = 0;
  let categoryCount = 0;
  let grandTotal = 0;
  let grandExpected = 0;

  const results = [];

  for (const s of allRows) {
    if (s.isCategory) {
      // Print previous category summary
      if (currentCategory) {
        results.push({ type: "cat_summary", category: currentCategory, actual: categoryActual, expected: categoryExpected, seriesCount: categoryCount });
      }
      currentCategory = s.series;
      categoryActual = 0;
      categoryExpected = 0;
      categoryCount = 0;
      results.push({ type: "category", name: s.series, stt: s.stt });
      continue;
    }

    // Find family code
    const { familyName, familyCode, method } = await findFamilyCode(s.series);
    let actualCount = 0;
    if (familyCode) {
      actualCount = await countProductsInCategory(familyCode);
    }

    results.push({
      type: "series",
      stt: s.stt,
      series: s.series,
      familyName,
      familyCode,
      method,
      expected: s.expected,
      actual: actualCount,
    });

    categoryActual += actualCount;
    categoryExpected += s.expected;
    categoryCount += 1;
    grandTotal += actualCount;
    grandExpected += s.expected;

    const status = familyCode ? `${familyCode}` : `KHONG TIM THAY`;
    const matchInfo = familyName && familyName.toLowerCase() !== s.series.toLowerCase() ? ` [${familyName}]` : "";
    const diff = actualCount !== s.expected ? ` (CL: ${s.expected})` : "";
    process.stdout.write(`  ${s.stt.padEnd(8)} ${s.series.padEnd(30)} -> ${String(actualCount).padStart(5)} SP  ${status}${matchInfo}${diff}\n`);

    await new Promise(r => setTimeout(r, 80));
  }

  // Last category summary
  if (currentCategory) {
    results.push({ type: "cat_summary", category: currentCategory, actual: categoryActual, expected: categoryExpected, seriesCount: categoryCount });
  }

  // Print summary table
  process.stdout.write(`\n${"=".repeat(80)}\n`);
  process.stdout.write(`THONG KE TONG HOP\n`);
  process.stdout.write(`${"=".repeat(80)}\n\n`);

  // Category summaries
  process.stdout.write(`${"Nhom danh muc".padEnd(45)} ${"Dong".padStart(5)} ${"Thuc te".padStart(8)} ${"Checklist".padStart(10)}\n`);
  process.stdout.write(`${"-".repeat(70)}\n`);
  for (const r of results) {
    if (r.type === "cat_summary") {
      process.stdout.write(`${r.category.padEnd(45)} ${String(r.seriesCount).padStart(5)} ${String(r.actual).padStart(8)} ${String(r.expected).padStart(10)}\n`);
    }
  }
  process.stdout.write(`${"-".repeat(70)}\n`);
  process.stdout.write(`${"TONG CONG".padEnd(45)} ${String(series.length).padStart(5)} ${String(grandTotal).padStart(8)} ${String(grandExpected).padStart(10)}\n`);

  // Save to JSON
  const outputPath = resolve("product-count-report.json");
  writeFileSync(outputPath, JSON.stringify({ results, grandTotal, grandExpected, seriesCount: series.length }, null, 2), "utf8");
  process.stdout.write(`\nDa luu chi tiet: ${outputPath}\n`);
}

main().catch(err => {
  process.stderr.write(`Loi: ${err.message}\n`);
  process.exit(1);
});
