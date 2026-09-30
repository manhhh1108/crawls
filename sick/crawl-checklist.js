#!/usr/bin/env node

const { execFileSync } = require("node:child_process");
const { writeFileSync, mkdirSync, existsSync, readdirSync, unlinkSync, readFileSync } = require("node:fs");
const { resolve, join } = require("node:path");
const { tmpdir } = require("node:os");
const { randomBytes } = require("node:crypto");

const {
  parseArgs,
  fetchJson,
  fetchCategoryProductIds,
  buildApiUrlFromProductId,
  buildProductUrlFromProductId,
  buildExcelTemplateData,
  writeXlsxWithPython,
  downloadAndWatermarkImages,
  toSafeFilename,
  buildImageBaseName,
  sanitizeText,
} = require("./crawl-sick-product");


// ============================================================
//  Doc checklist Excel -> danh sach dong san pham
// ============================================================
function readChecklistSeries(checklistPath) {
  const pythonCode = `
import json, sys, re
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
    # Optional category code in column C (idx 2)
    code_raw = cells[2] if len(cells) > 2 else None
    count_raw = cells[3] if len(cells) > 3 else None
    # Product category in column E (idx 4) and F (idx 5)
    cat_vi_raw = cells[4] if len(cells) > 4 else None
    cat_en_raw = cells[5] if len(cells) > 5 else None

    if stt_raw is None or series_raw is None:
        continue

    stt = str(stt_raw).strip()
    series = str(series_raw).strip()

    if not stt or not series:
        continue

    # Determine STT level by counting dots
    # 1       -> top category (skip)
    # 1.1     -> subcategory (skip)
    # 1.1.1   -> leaf family (CRAWL)
    stt_clean = stt.replace(',', '.')
    dot_count = stt_clean.count('.')

    # Only crawl leaf-level rows (3-level STTs).
    # If checklist has only 2 levels (no 3-level rows), treat 2-level as leaf.
    # Decide later via post-processing.
    level = dot_count + 1  # "1" -> 1, "1.1" -> 2, "1.1.1" -> 3

    # Try to parse as float to filter out malformed STTs
    try:
        # Replace last dot to allow "1.1.1" -> "1.1" as a sanity check
        float(stt_clean.split('.')[0] + '.' + (stt_clean.split('.')[1] if '.' in stt_clean else '0'))
    except ValueError:
        continue

    # Parse expected count
    expected = 0
    if count_raw is not None:
        try:
            expected = int(float(str(count_raw)))
        except (ValueError, TypeError):
            expected = 0

    # Parse category code if available
    cat_code = ""
    if code_raw is not None:
        cat_code = str(code_raw).strip()
        if not re.match(r'^g\\d+$', cat_code, re.IGNORECASE):
            cat_code = ""

    rows.append({
        "row": row_idx,
        "stt": stt,
        "level": level,
        "series": series,
        "expected": expected,
        "categoryCode": cat_code,
        "productCategoryVi": str(cat_vi_raw or "").strip(),
        "productCategoryEn": str(cat_en_raw or "").strip()
    })

# Decide leaf level: max level present in the file
if rows:
    max_level = max(r["level"] for r in rows)
    if max_level >= 3:
        rows = [r for r in rows if r["level"] >= 3]
    elif max_level == 2:
        rows = [r for r in rows if r["level"] >= 2]

print(json.dumps(rows, ensure_ascii=False))
`;

  const pythonCmd = process.platform === "win32" ? "python" : "python3";
  const tmpScript = join(tmpdir(), `sick_checklist_${randomBytes(6).toString("hex")}.py`);
  writeFileSync(tmpScript, pythonCode, "utf8");
  try {
    const out = execFileSync(pythonCmd, [tmpScript, checklistPath], {
      encoding: "utf8",
      timeout: 30000,
    });
    const lastLine = out.trim().split("\n").pop();
    return JSON.parse(lastLine);
  } catch (err) {
    throw new Error(`Khong doc duoc checklist: ${err.message}`);
  } finally {
    try { unlinkSync(tmpScript); } catch { }
  }
}

// ============================================================
//  Tim ProductFamilyCode tu ten dong san pham
// ============================================================
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

  // Collect all families from hits
  const families = new Map();
  for (const h of hits) {
    const mv = h?.masterValues ?? {};
    const fn = mv.ProductFamilyName || "";
    const fc = mv.ProductFamilyCode || "";
    if (fn && fc) {
      if (!families.has(fn)) {
        families.set(fn, { code: fc, count: 0 });
      }
      families.get(fn).count += 1;
    }
  }

  const nameLower = seriesName.trim().toLowerCase();

  // Exact match (case-insensitive)
  for (const [fn, info] of families) {
    if (fn.toLowerCase() === nameLower) {
      return { familyName: fn, familyCode: info.code, method: "EXACT" };
    }
  }

  // Partial match (one contains the other)
  for (const [fn, info] of families) {
    if (nameLower.includes(fn.toLowerCase()) || fn.toLowerCase().includes(nameLower)) {
      return { familyName: fn, familyCode: info.code, method: "PARTIAL" };
    }
  }

  // Best guess: most frequent
  if (families.size > 0) {
    let best = null;
    let bestCount = 0;
    for (const [fn, info] of families) {
      if (info.count > bestCount) {
        best = fn;
        bestCount = info.count;
      }
    }
    if (best) {
      return { familyName: best, familyCode: families.get(best).code, method: "GUESS" };
    }
  }

  return { familyName: null, familyCode: null, method: "NO_FAMILY" };
}

// ============================================================
//  Lay tat ca product IDs tu family code
//  Dung 3 phuong an song song roi UNION ket qua de bao phu het:
//    1. CategoryCode filter - chi product variants tag truc tiep
//    2. ProductFamilyCode filter - cac product co cung family code
//    3. Parent filter (de quy) - descend vao sub-categories/master products
// ============================================================
async function fetchByCategoryCode(familyCode) {
  const ids = await fetchCategoryProductIds(familyCode);
  return new Set(ids);
}

async function fetchByProductFamilyCode(familyCode) {
  const code = String(familyCode || "").replace(/[^a-z0-9]/gi, "").toLowerCase();
  if (!code) return new Set();
  const results = new Set();
  let page = 1, pageCount = 1;
  while (page <= pageCount) {
    const url = `https://www.sick.com/api/fact-finder/search/enSG-sick?query=*&filter=DefType:ProductVariant&filter=ProductFamilyCode:${code}&page=${page}&hitsPerPage=200&showPermutedSearchParams=false`;
    try {
      const data = await fetchJson(url);
      for (const hit of data?.hits ?? []) {
        const id = String(hit?.id ?? hit?.DefArticleNo ?? hit?.defArticleNo ?? "").trim();
        if (/^p\d+$/i.test(id)) results.add(id.toLowerCase());
      }
      pageCount = data?.paging?.pageCount ?? pageCount;
    } catch { break; }
    page += 1;
  }
  return results;
}

async function fetchByParentRecursive(parentCode, depth = 0, visited = new Set()) {
  const code = String(parentCode || "").trim().toLowerCase();
  if (!code || visited.has(code) || depth > 4) return new Set();
  visited.add(code);

  const results = new Set();
  let page = 1, pageCount = 1;
  while (page <= pageCount) {
    const url = `https://www.sick.com/api/fact-finder/search/enSG-sick?query=*&filter=Parent:${code}&page=${page}&hitsPerPage=200&showPermutedSearchParams=false`;
    try {
      const data = await fetchJson(url);
      for (const hit of data?.hits ?? []) {
        const mv = hit?.masterValues ?? {};
        const defType = mv.DefType || "";
        const id = String(hit?.id ?? mv.DefArticleNo ?? "").trim();

        if (defType === "ProductVariant" && /^p\d+$/i.test(id)) {
          results.add(id.toLowerCase());
        } else if (defType === "MasterProduct" || defType === "Category") {
          // De quy vao sub-category / master product
          const childCode = mv.DefArticleNo || id;
          if (childCode) {
            const childIds = await fetchByParentRecursive(childCode, depth + 1, visited);
            for (const cid of childIds) results.add(cid);
          }
        }
      }
      pageCount = data?.paging?.pageCount ?? pageCount;
    } catch { break; }
    page += 1;
  }
  return results;
}

async function fetchAllProductIds(familyCode) {
  // Chay 3 method song song, union ket qua
  const [byCat, byFam, byParent] = await Promise.all([
    fetchByCategoryCode(familyCode).catch(() => new Set()),
    fetchByProductFamilyCode(familyCode).catch(() => new Set()),
    fetchByParentRecursive(familyCode).catch(() => new Set()),
  ]);

  const union = new Set([...byCat, ...byFam, ...byParent]);
  if (union.size > byCat.size) {
    process.stdout.write(
      `  CategoryCode=${byCat.size}, FamilyCode=${byFam.size}, Parent(rec)=${byParent.size}, UNION=${union.size}\n`
    );
  }
  return [...union];
}

// ============================================================
//  Dem so file xlsx da co trong folder (de resume)
// ============================================================
function countExistingFiles(dir) {
  if (!existsSync(dir)) return { xlsxCount: 0, existingIds: new Set() };
  const files = readdirSync(dir);
  const xlsxFiles = files.filter((f) => f.endsWith(".xlsx"));
  const existingIds = new Set(xlsxFiles.map((f) => f.replace(".xlsx", "")));
  return { xlsxCount: xlsxFiles.length, existingIds };
}

// ============================================================
//  Image log: ghi cac product co loi anh ra image_log.csv
// ============================================================
const IMAGE_LOG_PATH = resolve("image_log.csv");
let _imageLogInited = false;

function appendImageLog(status, seriesName, productCode, productId, imageUrl, notes) {
  // Init header neu lan dau
  if (!_imageLogInited) {
    if (!existsSync(IMAGE_LOG_PATH)) {
      writeFileSync(IMAGE_LOG_PATH, "timestamp,status,series,product_code,product_id,image_url,notes\n", "utf8");
    }
    _imageLogInited = true;
  }
  // Escape CSV: bao boc trong " va escape "
  const esc = (v) => {
    const s = String(v || "");
    if (s.includes(",") || s.includes('"') || s.includes("\n")) {
      return '"' + s.replace(/"/g, '""') + '"';
    }
    return s;
  };
  const line = [
    new Date().toISOString(),
    status,
    esc(seriesName),
    esc(productCode),
    productId,
    esc(imageUrl),
    esc(notes),
  ].join(",") + "\n";
  try {
    require("node:fs").appendFileSync(IMAGE_LOG_PATH, line, "utf8");
  } catch { /* silent */ }
}

// ============================================================
//  Cache productId -> productCode (de resume nhanh, bo qua API call)
// ============================================================
function loadIdMap(seriesDir) {
  const path = resolve(seriesDir, ".id_map.json");
  if (!existsSync(path)) return {};
  try {
    return JSON.parse(readFileSync(path, "utf8"));
  } catch {
    return {};
  }
}

function saveIdMap(seriesDir, map) {
  const path = resolve(seriesDir, ".id_map.json");
  try {
    writeFileSync(path, JSON.stringify(map, null, 2), "utf8");
  } catch { }
}

// ============================================================
//  Crawl 1 dong san pham
// ============================================================
async function crawlOneSeries({ seriesName, familyCode, outputDir, skipImages, resume, delay, productCategoryVi = "", productCategoryEn = "" }) {
  const seriesFolderName = toSafeFilename(seriesName) || seriesName.replace(/[^a-zA-Z0-9_-]+/g, "_");
  const seriesDir = resolve(outputDir, seriesFolderName);
  const imagesDir = resolve(seriesDir, "images");

  if (!existsSync(seriesDir)) mkdirSync(seriesDir, { recursive: true });
  if (!skipImages && !existsSync(imagesDir)) mkdirSync(imagesDir, { recursive: true });

  // Lay danh sach san pham
  process.stdout.write(`  Dang lay danh sach san pham tu category ${familyCode}...\n`);
  const productIds = await fetchAllProductIds(familyCode);
  process.stdout.write(`  Tim thay ${productIds.length} san pham.\n`);

  if (productIds.length === 0) return { success: 0, failed: 0, skipped: 0, total: 0 };

  // Check resume: load existing file list and id->code cache
  const { existingIds } = countExistingFiles(seriesDir);
  const idMap = loadIdMap(seriesDir);  // productId -> safeName

  let success = 0;
  let failed = 0;
  let skipped = 0;

  for (const [idx, productId] of productIds.entries()) {
    const progress = `[${idx + 1}/${productIds.length}]`;

    // FAST RESUME: if we have cached productId -> safeName mapping AND that file exists,
    // skip without API call. Saves time on resume.
    if (resume && idMap[productId] && existingIds.has(idMap[productId])) {
      process.stdout.write(`  ${progress} ${productId} skip (cached)\n`);
      skipped += 1;
      continue;
    }

    try {
      const productUrl = buildProductUrlFromProductId(productId);
      const apiUrl = buildApiUrlFromProductId(productId);

      process.stdout.write(`  ${progress} ${productId} ...`);

      const payload = await fetchJson(apiUrl);
      const templateData = await buildExcelTemplateData({
        productUrl, apiUrl, payload,
        productCategoryVi, productCategoryEn,
      });

      const productCode = templateData.productHeader || productId;
      const partNumber = templateData.partNumber || "";
      const safeName = productCode.replace(/[^a-zA-Z0-9_-]+/g, "_");

      // Update cache so next resume can skip without API call
      idMap[productId] = safeName;

      // Resume: skip if file already exists by product code
      if (resume && existingIds.has(safeName)) {
        process.stdout.write(` skip (da co)\n`);
        skipped += 1;
        continue;
      }

      const outputPath = resolve(seriesDir, `${safeName}.xlsx`);
      writeXlsxWithPython(outputPath, templateData);
      existingIds.add(safeName);

      // Download images + log image status
      if (!skipImages) {
        const imageUrls = templateData.productImageUrls || [];
        if (imageUrls.length === 0) {
          // SICK API tra ve khong co anh nao
          appendImageLog("NO_IMAGE", seriesName, safeName, productId, "", "SICK API khong co image URL");
        } else {
          const imageBaseName = buildImageBaseName(productCode, partNumber);
          const imageFileName = toSafeFilename(imageBaseName);
          const expectedPath = resolve(imagesDir, `${imageFileName}.jpg`);
          const before = existsSync(expectedPath);
          const saved = downloadAndWatermarkImages(imageUrls, imagesDir, imageFileName);
          const after = existsSync(expectedPath);
          if (!after) {
            appendImageLog("DOWNLOAD_FAIL", seriesName, safeName, productId, imageUrls[0] || "", "downloadAndWatermark khong tao file");
          }
        }
      }

      process.stdout.write(` -> ${safeName}.xlsx\n`);
      success += 1;
    } catch (err) {
      process.stdout.write(` LOI: ${err.message}\n`);
      failed += 1;
    }

    // Save id-map periodically (every 20 products) and at end
    if (idx % 20 === 19) saveIdMap(seriesDir, idMap);

    // Delay between products to avoid rate limiting
    if (delay > 0 && idx < productIds.length - 1) {
      await new Promise((r) => setTimeout(r, delay));
    }
  }

  saveIdMap(seriesDir, idMap);
  return { success, failed, skipped, total: productIds.length };
}

// ============================================================
//  MAIN
// ============================================================
async function main() {
  const args = parseArgs(process.argv.slice(2));

  if (args.help) {
    process.stdout.write(`
Crawl san pham SICK theo checklist Excel, to chuc theo dong san pham.

Su dung:
  node crawl-checklist.js --checklist <file.xlsx> [options]

Options:
  --checklist <file>      File Excel checklist (bat buoc)
  --out-dir <folder>      Thu muc output (mac dinh: ./sick-output)
  --series "G2,W16"       Chi crawl cac dong chi dinh (cach nhau boi dau phay)
  --start-row <n>         Bat dau tu dong Excel so n
  --end-row <n>           Ket thuc o dong Excel so n
  --dry-run               Chi hien thi danh sach, khong crawl
  --no-images             Bo qua anh san pham
  --resume                Bo qua san pham da crawl (skip file da co)
  --delay <ms>            Delay giua moi san pham (mac dinh: 200ms)

Vi du:
  node crawl-checklist.js --checklist AUMI_SP_SICK_Checklist.xlsx --dry-run
  node crawl-checklist.js --checklist AUMI_SP_SICK_Checklist.xlsx --series "GLL70" --out-dir ./test-output
  node crawl-checklist.js --checklist AUMI_SP_SICK_Checklist.xlsx --resume
`);
    return;
  }

  if (!args.checklist) {
    process.stderr.write("Loi: Can truyen --checklist <file.xlsx>\n");
    process.exit(1);
  }

  const checklistPath = resolve(String(args.checklist));
  const outputDir = resolve(args["out-dir"] ?? "sick-output");
  const dryRun = !!args["dry-run"];
  const skipImages = !!args["no-images"];
  const resume = !!args.resume;
  const delay = parseInt(args.delay ?? "200", 10);
  const seriesFilter = args.series
    ? String(args.series).split(",").map((s) => s.trim().toLowerCase())
    : null;
  const startRow = args["start-row"] ? parseInt(args["start-row"], 10) : 0;
  const endRow = args["end-row"] ? parseInt(args["end-row"], 10) : Infinity;

  // Doc checklist
  process.stdout.write(`Doc checklist: ${checklistPath}\n`);
  let seriesList = readChecklistSeries(checklistPath);
  process.stdout.write(`Tim thay ${seriesList.length} dong san pham.\n\n`);

  // Filter
  if (seriesFilter) {
    seriesList = seriesList.filter((s) =>
      seriesFilter.some((f) => s.series.toLowerCase().includes(f) || f.includes(s.series.toLowerCase()))
    );
    process.stdout.write(`Sau filter --series: ${seriesList.length} dong.\n\n`);
  }
  if (startRow > 0 || endRow < Infinity) {
    seriesList = seriesList.filter((s) => s.row >= startRow && s.row <= endRow);
    process.stdout.write(`Sau filter row ${startRow}-${endRow}: ${seriesList.length} dong.\n\n`);
  }

  if (seriesList.length === 0) {
    process.stdout.write("Khong co dong san pham nao can xu ly.\n");
    return;
  }

  // Resolve family codes
  process.stdout.write("Dang xac dinh family code cho tung dong san pham...\n");
  const resolvedSeries = [];
  for (const s of seriesList) {
    let familyName = s.series;
    let familyCode = s.categoryCode || "";
    let method = familyCode ? "FROM_CHECKLIST" : null;

    // Fallback: search by name if checklist has no code
    if (!familyCode) {
      const found = await findFamilyCode(s.series);
      familyName = found.familyName || s.series;
      familyCode = found.familyCode;
      method = found.method;
      await new Promise((r) => setTimeout(r, 100));
    }

    resolvedSeries.push({ ...s, familyName, familyCode, method });

    const status = familyCode ? `${familyCode} (${method})` : `KHONG TIM THAY (${method})`;
    const matchInfo = familyName && familyName !== s.series ? ` [API: ${familyName}]` : "";
    process.stdout.write(`  ${s.stt} ${s.series.padEnd(35)} -> ${status}${matchInfo}\n`);
  }

  const totalExpected = resolvedSeries.reduce((sum, s) => sum + s.expected, 0);
  const resolved = resolvedSeries.filter((s) => s.familyCode);
  const unresolved = resolvedSeries.filter((s) => !s.familyCode);

  process.stdout.write(`\n=== TONG KET ===\n`);
  process.stdout.write(`Tong: ${resolvedSeries.length} dong | Co family code: ${resolved.length} | Khong tim thay: ${unresolved.length}\n`);
  process.stdout.write(`Tong SP du kien: ~${totalExpected}\n`);
  process.stdout.write(`Thu muc output: ${outputDir}\n\n`);

  if (unresolved.length > 0) {
    process.stdout.write(`CANH BAO - Cac dong khong tim thay:\n`);
    for (const s of unresolved) {
      process.stdout.write(`  Row ${s.row}: ${s.series}\n`);
    }
    process.stdout.write("\n");
  }

  if (dryRun) {
    process.stdout.write("(Dry run - khong crawl. Bo --dry-run de chay that.)\n");
    return;
  }

  // Crawl!
  if (!existsSync(outputDir)) mkdirSync(outputDir, { recursive: true });

  let grandSuccess = 0;
  let grandFailed = 0;
  let grandSkipped = 0;

  for (const [idx, s] of resolved.entries()) {
    process.stdout.write(`\n${"=".repeat(60)}\n`);
    process.stdout.write(`[${idx + 1}/${resolved.length}] Dong: ${s.series} (expected: ${s.expected} SP)\n`);
    process.stdout.write(`Family: ${s.familyName} | Code: ${s.familyCode} | Method: ${s.method}\n`);
    process.stdout.write(`${"=".repeat(60)}\n`);

    try {
      const result = await crawlOneSeries({
        seriesName: s.series,
        familyCode: s.familyCode,
        outputDir,
        skipImages,
        resume,
        delay,
        productCategoryVi: s.productCategoryVi || "",
        productCategoryEn: s.productCategoryEn || "",
      });

      grandSuccess += result.success;
      grandFailed += result.failed;
      grandSkipped += result.skipped;

      process.stdout.write(`  => Ket qua: ${result.success} OK, ${result.failed} loi, ${result.skipped} skip / ${result.total} tong\n`);
    } catch (err) {
      process.stderr.write(`  LOI DONG ${s.series}: ${err.message}\n`);
    }
  }

  process.stdout.write(`\n${"=".repeat(60)}\n`);
  process.stdout.write(`HOAN TAT TAT CA\n`);
  process.stdout.write(`Thanh cong: ${grandSuccess} | Loi: ${grandFailed} | Skip: ${grandSkipped}\n`);
  process.stdout.write(`Thu muc: ${outputDir}\n`);
  process.stdout.write(`${"=".repeat(60)}\n`);
}

main().catch((err) => {
  process.stderr.write(`\nLoi: ${err.message}\n`);
  process.exit(1);
});
