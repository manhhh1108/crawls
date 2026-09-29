#!/usr/bin/env node

/**
 * Retry tai anh cho cac product co xlsx nhung chua co .jpg.
 *
 * Logic:
 *   - Quet tat ca folder trong SICK_Products
 *   - Tim xlsx khong co .jpg matching trong images/
 *   - Voi moi xlsx do:
 *     - Tim productId tu .id_map.json
 *     - Goi API SICK lay image URLs
 *     - Neu co URL -> download + watermark -> luu vao images/
 *     - Neu khong co URL -> log "no image on source"
 *
 * Usage:
 *   node retry_images.js                    # retry tat ca
 *   node retry_images.js --family AHS_AHM36 # 1 dong (test)
 *   node retry_images.js --delay 500        # delay giua moi product (ms)
 */

const fs = require("node:fs");
const path = require("node:path");
const {
  fetchJson,
  buildApiUrlFromProductId,
  buildExcelTemplateData,
  downloadAndWatermarkImages,
  toSafeFilename,
  buildImageBaseName,
  parseArgs,
} = require("./crawl-sick-product");

const OUTPUT_DIR = "SICK_Products";

function sleep(ms) { return new Promise(r => setTimeout(r, ms)); }

function findXlsxWithoutImage(rootDir) {
  const result = [];  // [{folder, xlsxName, xlsxPath, imagesDir}]
  const folders = fs.readdirSync(rootDir).filter(f => {
    return fs.statSync(path.join(rootDir, f)).isDirectory();
  });

  for (const folder of folders) {
    const folderPath = path.join(rootDir, folder);
    const imagesDir = path.join(folderPath, "images");
    const xlsxFiles = fs.readdirSync(folderPath)
      .filter(f => f.endsWith(".xlsx") && !f.startsWith("~$"))
      .map(f => f.replace(".xlsx", ""));

    let imgPrefixes = new Set();
    if (fs.existsSync(imagesDir)) {
      const imgs = fs.readdirSync(imagesDir).filter(f => f.endsWith(".jpg"));
      for (const img of imgs) {
        const n = img.replace(".jpg", "");
        for (const xname of xlsxFiles) {
          if (n === xname || n.startsWith(xname + "-")) {
            imgPrefixes.add(xname);
            break;
          }
        }
      }
    }

    for (const xname of xlsxFiles) {
      if (!imgPrefixes.has(xname)) {
        result.push({
          folder,
          xlsxName: xname,
          xlsxPath: path.join(folderPath, `${xname}.xlsx`),
          imagesDir,
        });
      }
    }
  }
  return result;
}

function loadIdMap(folderPath) {
  const p = path.join(folderPath, ".id_map.json");
  if (!fs.existsSync(p)) return {};
  try {
    return JSON.parse(fs.readFileSync(p, "utf8"));
  } catch { return {}; }
}

function findProductIdForSafeName(idMap, safeName) {
  for (const [pid, sname] of Object.entries(idMap)) {
    if (sname === safeName) return pid;
  }
  return null;
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const familyFilter = args["family"] || null;
  const delayMs = parseInt(args["delay"] || "300", 10);

  console.log("Quet folder de tim xlsx khong co anh...");
  let missing = findXlsxWithoutImage(OUTPUT_DIR);
  if (familyFilter) {
    missing = missing.filter(m => m.folder.toLowerCase().includes(familyFilter.toLowerCase()));
  }
  console.log(`Tim thay ${missing.length} xlsx thieu anh.\n`);

  if (missing.length === 0) {
    console.log("Khong co gi de retry.");
    return;
  }

  let success = 0;
  let noImageOnSource = 0;
  let failed = 0;
  let noProductId = 0;
  const noImageList = [];

  for (const [idx, item] of missing.entries()) {
    const { folder, xlsxName, imagesDir } = item;
    const folderPath = path.join(OUTPUT_DIR, folder);
    const idMap = loadIdMap(folderPath);
    const productId = findProductIdForSafeName(idMap, xlsxName);

    process.stdout.write(`[${idx + 1}/${missing.length}] ${folder}/${xlsxName} ... `);

    if (!productId) {
      console.log("KHONG TIM THAY productId trong id_map");
      noProductId++;
      continue;
    }

    try {
      if (!fs.existsSync(imagesDir)) {
        fs.mkdirSync(imagesDir, { recursive: true });
      }

      const apiUrl = buildApiUrlFromProductId(productId);
      const payload = await fetchJson(apiUrl);
      // Khong dich (translate=false) - chi can image URLs
      const tpl = await buildExcelTemplateData({
        productUrl: "", apiUrl, payload, translate: false,
      });

      const imageUrls = tpl.productImageUrls || [];
      if (imageUrls.length === 0) {
        console.log("khong co anh tren SICK (source)");
        noImageOnSource++;
        noImageList.push(`${folder}/${xlsxName}`);
        continue;
      }

      const productCode = tpl.productHeader || xlsxName;
      const partNumber = tpl.partNumber || "";
      const imageBaseName = buildImageBaseName(productCode, partNumber);
      const imageFileName = toSafeFilename(imageBaseName);

      const saved = downloadAndWatermarkImages(imageUrls, imagesDir, imageFileName);
      if (saved && saved.length > 0) {
        success++;
        // saved msg da in tu downloadAndWatermarkImages
      } else {
        failed++;
        console.log("download/watermark FAIL");
      }
    } catch (err) {
      console.log(`ERROR: ${err.message}`);
      failed++;
    }

    await sleep(delayMs);
  }

  console.log();
  console.log("=".repeat(60));
  console.log(`Tong:                    ${missing.length}`);
  console.log(`Thanh cong:              ${success}`);
  console.log(`Khong co anh tren source: ${noImageOnSource}`);
  console.log(`Khong tim thay productId: ${noProductId}`);
  console.log(`Failed:                  ${failed}`);

  if (noImageList.length > 0) {
    console.log();
    console.log(`Cac product khong co anh tren SICK source (${noImageList.length}):`);
    for (const x of noImageList.slice(0, 30)) {
      console.log(`  ${x}`);
    }
    if (noImageList.length > 30) {
      console.log(`  ... va ${noImageList.length - 30} cai khac`);
    }
  }
}

main().catch(err => {
  console.error("FATAL:", err.message);
  process.exit(1);
});
