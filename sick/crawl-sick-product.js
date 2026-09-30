#!/usr/bin/env node


const { execFileSync } = require("node:child_process");
const { writeFileSync, mkdirSync, existsSync, unlinkSync } = require("node:fs");
const { resolve, join, basename } = require("node:path");
const { tmpdir } = require("node:os");
const { randomBytes } = require("node:crypto");

const DEFAULT_PRODUCT_URL =
  "https://www.sick.com/sg/en/catalog/products/detection-sensors/photoelectric-sensors/w16/wla16p-1hg12100zzz/p/p687909";
const DEFAULT_API_URL =
  "https://www.sick.com/api/fact-finder/records/detail/enSG-sick/p687909?usePersonalization=false&withCampaigns=false&withRecommendations=false&withSimilarProducts=false";
const DEFAULT_OUTPUT_FILE = "sick-product.xlsx";

function parseArgs(argv) {
  const args = {};
  for (let i = 0; i < argv.length; i += 1) {
    const current = argv[i];
    if (!current.startsWith("--")) continue;
    const key = current.slice(2);
    const next = argv[i + 1];
    if (!next || next.startsWith("--")) {
      args[key] = true;
      continue;
    }
    args[key] = next;
    i += 1;
  }
  return args;
}

function getProductIdFromUrl(productUrl) {
  const match = productUrl.match(/\/p\/(p\d+)(?:$|[/?#])/i);
  if (!match?.[1]) {
    throw new Error("Khong tim thay product id dang /p/p123456 trong URL.");
  }
  return match[1];
}

function buildApiUrlFromProductUrl(productUrl) {
  const productId = getProductIdFromUrl(productUrl);
  return `https://www.sick.com/api/fact-finder/records/detail/enSG-sick/${productId}?usePersonalization=false&withCampaigns=false&withRecommendations=false&withSimilarProducts=false`;
}

function buildApiUrlFromProductId(productId) {
  const pid = String(productId || "").trim();
  if (!pid) {
    throw new Error("Khong tim thay product id.");
  }
  return `https://www.sick.com/api/fact-finder/records/detail/enSG-sick/${pid}?usePersonalization=false&withCampaigns=false&withRecommendations=false&withSimilarProducts=false`;
}

function buildProductUrlFromProductId(productId) {
  const pid = String(productId || "").trim();
  if (!pid) {
    return "";
  }
  return `https://www.sick.com/sg/en/p/${pid}`;
}

function isCategoryUrl(input) {
  return /\/c\/g\d+(?:[/?#]|$)/i.test(String(input || ""));
}

function getCategoryCodeFromInput(input) {
  const text = String(input || "");
  const urlMatch = text.match(/\/c\/(g\d+)(?:[/?#]|$)/i);
  if (urlMatch?.[1]) return urlMatch[1].toLowerCase();
  const codeMatch = text.match(/\b(g\d+)\b/i);
  if (codeMatch?.[1]) return codeMatch[1].toLowerCase();
  return "";
}

async function fetchJson(url, retries = 3) {
  for (let attempt = 1; attempt <= retries; attempt++) {
    try {
      const response = await fetch(url, {
        headers: {
          accept: "application/json",
          "user-agent": "Mozilla/5.0 (Node.js crawler)",
        },
      });

      if (response.status === 429) {
        process.stderr.write(`Rate limited, retry ${attempt}/${retries}...\n`);
        await new Promise((r) => setTimeout(r, 2000 * attempt));
        continue;
      }

      if (!response.ok) {
        if (attempt < retries && response.status >= 500) {
          await new Promise((r) => setTimeout(r, 1000 * attempt));
          continue;
        }
        throw new Error(`Call API that bai (${response.status} ${response.statusText}).`);
      }

      return response.json();
    } catch (err) {
      if (attempt === retries) throw err;
      await new Promise((r) => setTimeout(r, 1000 * attempt));
    }
  }
}

async function fetchCategoryProductIds(categoryCode) {
  const code = getCategoryCodeFromInput(categoryCode);
  if (!code) {
    throw new Error("Khong tim thay category code dang g123456.");
  }

  const results = new Set();
  let page = 1;
  let pageCount = 1;
  while (page <= pageCount) {
    const url = `https://www.sick.com/api/fact-finder/search/enSG-sick?query=*&filter=DefType:ProductVariant&filter=CategoryCode:${code}&page=${page}&showPermutedSearchParams=false`;
    const data = await fetchJson(url);
    const hits = data?.hits ?? [];
    for (const hit of hits) {
      const rawId = hit?.id ?? hit?.DefArticleNo ?? hit?.defArticleNo;
      const id = String(rawId || "").trim();
      if (/^p\d+$/i.test(id)) {
        results.add(id.toLowerCase());
      }
    }
    pageCount = data?.paging?.pageCount ?? pageCount;
    page += 1;
  }

  return [...results];
}

// Import module dich thuat chuyen nganh (tu dien mo rong + bao ve thuat ngu ky thuat)
const { TERM_DICT, translateBatchSmart } = require("./sick-translate");

// Alias de tuong thich voi code cu
async function translateBatch(texts) {
  return translateBatchSmart(texts);
}

function safeParseJson(input) {
  if (typeof input !== "string" || input.trim() === "") return null;
  try {
    return JSON.parse(input);
  } catch {
    return null;
  }
}

function toText(value) {
  if (value === null || value === undefined) return "";
  if (typeof value === "string") return value;
  if (typeof value === "number" || typeof value === "boolean") return String(value);
  return JSON.stringify(value);
}

const SUP_DIGITS = "⁰¹²³⁴⁵⁶⁷⁸⁹";
const SUB_DIGITS = "₀₁₂₃₄₅₆₇₈₉";

function sanitizeText(text) {
  return toText(text)
    // Strip footnote-style <sup> markers (digits with closing paren), e.g. <sup>1)</sup>
    .replace(/<sup[^>]*>\s*\d{1,2}\)\s*<\/sup>/gi, "")
    // Convert math <sup>...</sup> to Unicode superscript: 10<sup>6</sup> -> 10⁶, 10<sup>-8</sup> -> 10⁻⁸
    .replace(/<sup[^>]*>([^<]*)<\/sup>/gi, (_, content) =>
      content
        .replace(/[0-9]/g, (c) => SUP_DIGITS[c])
        .replace(/-/g, "⁻")
        .replace(/\+/g, "⁺")
    )
    // Convert <sub>...</sub> to Unicode subscript for digits: B<sub>10d</sub> -> B₁₀d
    .replace(/<sub[^>]*>([^<]*)<\/sub>/gi, (_, content) =>
      content.replace(/[0-9]/g, (c) => SUB_DIGITS[c])
    )
    // Convert <br>, </li><li>, </p><p> sang newline de cell xuong dong de doc
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/li>\s*<li[^>]*>/gi, "\n")
    .replace(/<\/p>\s*<p[^>]*>/gi, "\n")
    .replace(/<li[^>]*>/gi, "")
    .replace(/<\/li>/gi, "")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/g, " ")
    // Normalize spaces nhung GIU LAI newline
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function stripFootnoteMarkers(text) {
  if (!text) return text;
  return text
    // Strip footnote markers embedded in values: " 1)", " 2)", " 12)" etc.
    // when followed by whitespace, end of string, or punctuation.
    // Requires whitespace BEFORE the digits to avoid eating last digits of legitimate
    // numbers like "(IEC 61508)" -> "(IEC 615" bug.
    .replace(/[ \t]+\d{1,2}\)(?=\s|$|[;,.:])/g, "")
    // Strip footnote marker that is the entire (remaining) value
    .replace(/^\s*\d{1,2}\)\s*$/g, "")
    // Normalize horizontal whitespace, preserve newlines
    .replace(/[ \t]+/g, " ")
    .replace(/\n[ \t]+/g, "\n")
    .replace(/[ \t]+\n/g, "\n")
    .trim();
}

function parseHtmlTableRows(html) {
  const rows = [];
  const trRegex = /<tr[^>]*>([\s\S]*?)<\/tr>/gi;
  let trMatch = trRegex.exec(html);

  while (trMatch) {
    const trContent = trMatch[1];
    const tdMatches = [...trContent.matchAll(/<td[^>]*>([\s\S]*?)<\/td>/gi)];
    if (tdMatches.length >= 2) {
      const key = sanitizeText(tdMatches[0][1]);
      const value = stripFootnoteMarkers(sanitizeText(tdMatches[1][1]));
      if (key || value) {
        rows.push({ key, value });
      }
    }
    trMatch = trRegex.exec(html);
  }

  return rows;
}

function parseTechData(master) {
  const rawTechData = safeParseJson(master.TechData);
  if (!Array.isArray(rawTechData)) return [];

  const specs = [];
  for (const entry of rawTechData) {
    const value = toText(entry?.Value);
    if (!value) continue;
    const titleText = sanitizeText(entry?.Title);

    if (/<table/i.test(value) && /<tr/i.test(value)) {
      // Section voi nhieu sub-rows -> emit section header + sub-rows
      const tableRows = parseHtmlTableRows(value);
      if (tableRows.length > 0) {
        if (titleText) {
          specs.push({ key: titleText, value: "", isSection: true });
        }
        specs.push(...tableRows);
      }
      continue;
    }
    // Single-value entry: 1 row co Title la key
    const plainValue = stripFootnoteMarkers(sanitizeText(value));
    if (plainValue) {
      specs.push({
        key: titleText || "Thông số",
        value: plainValue,
      });
    }
  }
  return specs;
}

function extractProductImageUrls(master) {
  const images = [];

  // DefPictureUrl - main product image
  const defPicture = safeParseJson(master.DefPictureUrl);
  if (Array.isArray(defPicture)) {
    for (const url of defPicture) {
      if (typeof url === "string" && url.startsWith("http")) {
        images.push(url);
      }
    }
  } else if (typeof master.DefPictureUrl === "string" && master.DefPictureUrl.startsWith("http")) {
    images.push(master.DefPictureUrl);
  }

  // Gallery images
  const gallery = safeParseJson(master.Gallery);
  if (Array.isArray(gallery)) {
    for (const item of gallery) {
      const url = item?.URL || item?.Url || item?.url;
      if (typeof url === "string" && url.startsWith("http") && !images.includes(url)) {
        images.push(url);
      }
    }
  }

  return images;
}

function toSlugFromProductName(name) {
  return sanitizeText(name)
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function toVietnameseSlug(text) {
  return sanitizeText(text)
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .replace(/đ/gi, "d")
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function toSafeFilename(name) {
  return sanitizeText(name)
    .replace(/[^a-zA-Z0-9_-]+/g, "_")
    .replace(/^_+|_+$/g, "");
}

function buildImageBaseName(productCode, partNumber) {
  // Anh file luc trung ten voi xlsx (productCode), khong them partNumber.
  const code = sanitizeText(productCode);
  const part = sanitizeText(partNumber);
  if (code) return code;
  if (part) return part;
  return "sick-product";
}

/**
 * Chuan hoa du lieu ve dung template Excel 2 ngon ngu, co dich sang tieng Viet.
 */
async function buildExcelTemplateData({ productUrl, apiUrl, payload, translate = true, productCategoryVi = "", productCategoryEn = "", brand = "SICK Sensor" }) {
  const master = payload?.record?.masterValues ?? {};
  const productCode = sanitizeText(master.DefName);
  const partNumber = sanitizeText(master.PartNumber);
  // Fallback chain: DefNameAddon (specific name) -> ProductFamilyName -> "Industrial sensor"
  // Do NOT default to "Photoelectric sensors" because that mislabels non-photoelectric products
  // (encoders, safety switches, ultrasonic, etc.)
  const categoryName = sanitizeText(
    master.DefNameAddon || master.ProductFamilyName || "Industrial sensor"
  );
  const description = sanitizeText(master.DefDescription || master.Description);

  const onlineDataSheet = safeParseJson(master.OnlineDataSheetUrl);
  let datasheetUrlEn = "";
  let datasheetUrlVi = "";
  if (Array.isArray(onlineDataSheet)) {
    const findUrl = (lang) => {
      const item = onlineDataSheet.find((it) => String(it?.Language || "").toLowerCase() === lang);
      return item ? sanitizeText(item.URL) : "";
    };
    datasheetUrlEn = findUrl("en") || sanitizeText(onlineDataSheet[0]?.URL || "");
    datasheetUrlVi = findUrl("vi") || datasheetUrlEn;  // Fallback EN neu khong co VI
  }
  // Backward-compat: giu firstDataSheetUrl cho code cu
  const firstDataSheetUrl = datasheetUrlEn;

  const detailSpecs = parseTechData(master)
    .filter((s) => s.key && (s.isSection || (s.value && s.value.trim())));
  // shortDescription: skip section headers (chi co key, khong value)
  const specsInOneLine = detailSpecs
    .filter((item) => !item.isSection && item.value)
    .map((item) => `${item.key}: ${item.value}`)
    .slice(0, 8)
    .join("; ");

  const shortTitleEn = description || `${categoryName} ${productCode} SICK`;
  const metaTitleEn = sanitizeText(master.MetaTitle) || `${productCode} | ${categoryName}`;
  const metaDescriptionEn = sanitizeText(master.MetaDescription);
  const productSlug = toSlugFromProductName(productCode) || "ma-san-pham-sick";

  // Thu thap moi text can dich
  let translations = {};
  if (translate) {
    const textsToTranslate = [
      categoryName,
      description,
      shortTitleEn,
      metaTitleEn,
      metaDescriptionEn,
      ...detailSpecs.map((s) => s.key),
      ...detailSpecs.map((s) => s.value),
    ];
    try {
      translations = await translateBatch(textsToTranslate);
    } catch (err) {
      process.stderr.write(`Canh bao: loi dich thuat - ${err.message}\n`);
    }
  }
  const tr = (text) => translations[text] || text;

  const categoryNameVi = tr(categoryName);
  const descriptionVi = tr(description);
  const shortTitleVi = descriptionVi || `${categoryNameVi} ${productCode} SICK`;
  const specsInOneLineVi = detailSpecs
    .filter((item) => !item.isSection && item.value)
    .map((item) => `${tr(item.key)}: ${tr(item.value)}`)
    .slice(0, 8)
    .join("; ");
  const metaDescriptionVi = `Mua ${categoryNameVi} ${productCode} chinh hang SICK. ${specsInOneLineVi || "Thong so cap nhat tu nha san xuat."}`;

  const detailSpecsBilingual = detailSpecs.map((s) => ({
    keyVi: tr(s.key),
    valueVi: s.isSection ? "" : tr(s.value),
    keyEn: s.key,
    valueEn: s.isSection ? "" : s.value,
    isSection: !!s.isSection,
  }));

  // === URL va slug ===
  const categorySlugVi = toVietnameseSlug(categoryNameVi);
  const categorySlugEn = toVietnameseSlug(categoryName);
  const urlExampleVi = `${productSlug}-sick`;
  const urlExampleEn = `${productSlug}-sick`;

  // === Meta title (khong brackets) ===
  // Neu category la "Archive" / "Lưu trữ" thi BO QUA prefix - khong co y nghia voi user
  const isArchiveCat = /^archive$/i.test(categoryName) || /^lưu trữ$/i.test(categoryNameVi);
  const titleCatVi = isArchiveCat ? "" : `${categoryNameVi} `;
  const titleCatEn = isArchiveCat ? "" : `${categoryName} `;
  const metaTitleViValue = `${titleCatVi}${productCode} SICK | AUMI`;
  const metaTitleEnValue = `${titleCatEn}${productCode} SICK | AUMI`;

  // === Meta description (155-165 chars) ===
  const metaDescViFull = `Mua ${titleCatVi}${productCode} chính hãng – thông số đầy đủ, giá tốt, giao nhanh. AUMI tư vấn kỹ thuật, CO/CQ, bảo hành theo tiêu chuẩn hãng, hỗ trợ lựa chọn model phù hợp.`;
  const metaDescEnFull = `Buy genuine ${titleCatEn}${productCode} – full specifications, good price, fast delivery. AUMI provides technical consulting, CO/CQ, warranty per manufacturer standards, and helps choose the right model.`;

  // === Templates (col C, F - static voi placeholder) ===
  const urlTemplateVi = `[Mã SP] - SICK (Ví dụ sản phẩm này là ${urlExampleVi})`;
  const urlTemplateEn = `[Mã SP] - SICK (Ví dụ sản phẩm này là ${urlExampleEn})`;

  const metaTitleTemplateVi = `[Loại sản phẩm] [Mã SP] SICK | AUMI (VD sản phẩm này là: ${categoryNameVi} ${productCode} SICK | AUMI)`;
  const metaTitleTemplateEn = `[Loại sản phẩm] [Mã SP] SICK | AUMI (VD sản phẩm này là: ${categoryName} ${productCode} SICK | AUMI)`;

  const metaDescTemplateVi = `Mua [Loại sản phẩm] [Mã SP] chính hãng – thông số đầy đủ, giá tốt, giao nhanh. AUMI tư vấn kỹ thuật, CO/CQ, bảo hành theo tiêu chuẩn hãng, hỗ trợ lựa chọn model phù hợp.
(VD: Mua ${categoryNameVi.toLowerCase()} ${productCode} chính hãng – thông số đầy đủ, giá tốt, giao nhanh. AUMI tư vấn kỹ thuật, CO/CQ, bảo hành theo tiêu chuẩn hãng, hỗ trợ lựa chọn model phù hợp.)`;
  const metaDescTemplateEn = `Buy genuine [Product type] [Product code] – full specifications, good price, fast delivery. AUMI provides technical consulting, CO/CQ, warranty per manufacturer standards, and helps choose the right model.
(VD: Buy genuine ${categoryName} ${productCode} – full specifications, good price, fast delivery. AUMI provides technical consulting, CO/CQ, warranty per manufacturer standards, and helps choose the right model.)`;

  const longDescTemplateVi = `Sản phẩm [Loại sản phẩm] [Mã SP] được thiết kế theo tiêu chuẩn công nghiệp châu Âu, cung cấp đầy đủ các thông số kỹ thuật quan trọng giúp kỹ sư dễ dàng lựa chọn và tích hợp vào hệ thống. Các thông số bao gồm dải đo/khả năng phát hiện, độ chính xác, thời gian phản hồi, tín hiệu đầu ra, nguồn cấp, cấp bảo vệ IP và điều kiện môi trường làm việc.

Tùy theo từng model và ứng dụng cụ thể, thiết bị có thể hỗ trợ nhiều tùy chọn kết nối, giao thức truyền thông hoặc các chức năng nâng cao nhằm đáp ứng yêu cầu về tự động hóa, an toàn và độ tin cậy trong môi trường công nghiệp.

Dưới đây là bảng thông số kỹ thuật chi tiết của sản phẩm:`;
  const longDescTemplateEn = `The product [Loại sản phẩm] [Mã SP] is designed in accordance with European industrial standards, providing all essential technical specifications to help engineers easily select and integrate it into their systems. These specifications include measuring range/detection capability, accuracy, response time, output signal, power supply, IP protection rating, and operating environmental conditions.

Depending on the specific model and application, the device may support various connection options, communication protocols, or advanced features to meet the requirements of automation, safety, and reliability in industrial environments.

Below is the detailed technical specification table of the product:`;

  // Thay [Loai san pham] va [Ma SP] bang gia tri thuc cho cot B va E
  // Neu category la "Archive" -> bo ca prefix + space sau (collapsed)
  const longCatVi = isArchiveCat ? "" : categoryNameVi;
  const longCatEn = isArchiveCat ? "" : categoryName;
  const longDescriptionViValue = longDescTemplateVi
    .replace(/Sản phẩm \[Loại sản phẩm\] /g, longCatVi ? `Sản phẩm ${longCatVi} ` : "Sản phẩm ")
    .replace(/\[Loại sản phẩm\]/g, longCatVi)
    .replace(/\[Mã SP\]/g, productCode);
  const longDescriptionEnValue = longDescTemplateEn
    .replace(/The product \[Loại sản phẩm\] /g, longCatEn ? `The product ${longCatEn} ` : "The product ")
    .replace(/\[Loại sản phẩm\]/g, longCatEn)
    .replace(/\[Mã SP\]/g, productCode);

  return {
    productHeader: productCode,
    partNumber,
    // Values (col B, E)
    codeVi: `${productCode} Part no.: ${partNumber}`,
    codeEn: `${productCode} Part no.: ${partNumber}`,
    productCategoryVi: productCategoryVi || "",
    productCategoryEn: productCategoryEn || "",
    brandVi: brand,
    brandEn: brand,
    urlVi: `${productSlug}-sick`,
    urlEn: `${productSlug}-sick`,
    keywordVi: `${productSlug} sick`,
    keywordEn: `${productSlug} sick`,
    metaTitleVi: metaTitleViValue,
    metaTitleEn: metaTitleEnValue,
    metaDescriptionVi: metaDescViFull,
    metaDescriptionEn: metaDescEnFull,
    shortDescriptionVi: specsInOneLineVi || descriptionVi,
    shortDescriptionEn: specsInOneLine || description,
    longDescriptionVi: longDescriptionViValue,
    longDescriptionEn: longDescriptionEnValue,
    // Templates (col C, F)
    urlTemplateVi,
    urlTemplateEn,
    metaTitleTemplateVi,
    metaTitleTemplateEn,
    metaDescTemplateVi,
    metaDescTemplateEn,
    longDescTemplateVi,
    longDescTemplateEn,
    // Misc
    detailSpecs: detailSpecsBilingual,
    datasheetUrl: firstDataSheetUrl,  // backward-compat
    datasheetUrlEn,
    datasheetUrlVi,
    productImageUrls: extractProductImageUrls(master),
    raw: { productUrl, apiUrl, masterValues: master },
  };
}

/**
 * Tai 1 anh san pham, tao canvas 500x500 trang, gan logo AUMI + hotline, luu JPG < 100kb.
 */
function downloadAndWatermarkImages(imageUrls, outputDir, productCode) {
  if (!imageUrls || imageUrls.length === 0) {
    process.stdout.write("Khong tim thay anh san pham tu API.\n");
    return [];
  }

  if (!existsSync(outputDir)) {
    mkdirSync(outputDir, { recursive: true });
  }

  const pythonCode = `
import json, sys, os, urllib.request, tempfile
from PIL import Image, ImageDraw, ImageFont

image_urls = json.loads(sys.argv[1])
output_dir = sys.argv[2]
product_code = sys.argv[3]

AUMI_LOGO_URL = "https://aumi.com.vn/wp-content/uploads/2020/09/logo.png"
AUMI_HOTLINE = "+84 917 991 589"

CANVAS_SIZE = 500

# Download AUMI logo
aumi_logo = None
try:
    logo_tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
    logo_tmp.close()
    req = urllib.request.Request(AUMI_LOGO_URL, headers={
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
        "Referer": "https://aumi.com.vn/",
    })
    with urllib.request.urlopen(req) as resp:
        with open(logo_tmp.name, 'wb') as f:
            f.write(resp.read())
    aumi_logo = Image.open(logo_tmp.name).convert("RGBA")
except Exception as e:
    print(f"Khong tai duoc logo AUMI: {e}")

# Load fonts
def load_font(size):
    for fp in [
        "C:/Windows/Fonts/arialbd.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/calibrib.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
        "/System/Library/Fonts/Helvetica.ttc",
    ]:
        if os.path.exists(fp):
            try:
                return ImageFont.truetype(fp, size)
            except Exception:
                continue
    return ImageFont.load_default()

font_hotline = load_font(18)
font_aumi_fallback = load_font(26)

saved_files = []

# Chi xu ly anh dau tien (anh dai dien)
if image_urls:
    url = image_urls[0]
    try:
        tmp = tempfile.NamedTemporaryFile(suffix='.png', delete=False)
        tmp.close()
        req = urllib.request.Request(url, headers={"User-Agent": "Mozilla/5.0"})
        with urllib.request.urlopen(req) as resp:
            with open(tmp.name, 'wb') as f:
                f.write(resp.read())

        src = Image.open(tmp.name).convert("RGBA")

        # Tao canvas trang 500x500
        canvas = Image.new("RGB", (CANVAS_SIZE, CANVAS_SIZE), (255, 255, 255))

        # Vung dat anh san pham: chua banner tren (khoang 70px), padding 10px xung quanh
        top_banner_h = 70
        padding = 15
        product_area_w = CANVAS_SIZE - 2 * padding
        product_area_h = CANVAS_SIZE - top_banner_h - padding

        # Resize anh san pham giu ti le
        ratio = min(product_area_w / src.width, product_area_h / src.height)
        new_w = int(src.width * ratio)
        new_h = int(src.height * ratio)
        product_img = src.resize((new_w, new_h), Image.LANCZOS)

        # Dat anh san pham vao giua vung san pham
        px = (CANVAS_SIZE - new_w) // 2
        py = top_banner_h + (product_area_h - new_h) // 2
        # Paste voi alpha mask de giu nen trong suot
        if product_img.mode == "RGBA":
            canvas.paste(product_img, (px, py), product_img)
        else:
            canvas.paste(product_img, (px, py))

        draw = ImageDraw.Draw(canvas)

        # === LOGO AUMI (top-left) ===
        logo_margin = 15
        logo_max_h = 50
        if aumi_logo:
            logo = aumi_logo.copy()
            logo_ratio = logo_max_h / logo.height
            logo_w = int(logo.width * logo_ratio)
            if logo_w > 180:
                logo_w = 180
                logo_ratio = logo_w / aumi_logo.width
                logo = aumi_logo.resize((logo_w, int(aumi_logo.height * logo_ratio)), Image.LANCZOS)
            else:
                logo = logo.resize((logo_w, logo_max_h), Image.LANCZOS)
            ly = (top_banner_h - logo.height) // 2
            canvas.paste(logo, (logo_margin, ly), logo)
        else:
            # Fallback: chu AUMI mau xanh
            draw.text((logo_margin, 15), "AUMI", fill=(0, 82, 155), font=font_aumi_fallback)
            draw.text((logo_margin, 45), "auto process", fill=(100, 100, 100), font=font_hotline)

        # === HOTLINE (top-right) ===
        hotline_text = AUMI_HOTLINE
        # Icon dien thoai (mobile phone silhouette)
        phone_w = 22
        phone_h = 34
        text_bbox = draw.textbbox((0, 0), hotline_text, font=font_hotline)
        text_w = text_bbox[2] - text_bbox[0]
        total_w = phone_w + 10 + text_w
        right_margin = 15
        start_x = CANVAS_SIZE - right_margin - total_w
        center_y = top_banner_h // 2

        icon_x = start_x
        icon_y = center_y - phone_h // 2

        # Phone body (filled rounded rectangle, dark)
        body_color = (40, 40, 40)
        screen_color = (255, 255, 255)
        draw.rounded_rectangle(
            [icon_x, icon_y, icon_x + phone_w, icon_y + phone_h],
            radius=4, fill=body_color
        )
        # Phone screen (white inner rectangle)
        draw.rounded_rectangle(
            [icon_x + 2, icon_y + 5, icon_x + phone_w - 2, icon_y + phone_h - 7],
            radius=1, fill=screen_color
        )
        # Home button (small circle at bottom)
        btn_size = 3
        btn_cx = icon_x + phone_w // 2
        btn_cy = icon_y + phone_h - 4
        draw.ellipse(
            [btn_cx - btn_size, btn_cy - btn_size, btn_cx + btn_size, btn_cy + btn_size],
            outline=screen_color, width=1
        )
        # Speaker (small line at top)
        spk_y = icon_y + 3
        draw.line(
            [icon_x + phone_w // 2 - 3, spk_y, icon_x + phone_w // 2 + 3, spk_y],
            fill=screen_color, width=1
        )

        # Ve text hotline
        draw.text(
            (icon_x + phone_w + 8, center_y - 10),
            hotline_text,
            fill=(40, 40, 40),
            font=font_hotline,
        )

        # Filename = ma san pham
        filename = f"{product_code}.jpg"
        out_path = os.path.join(output_dir, filename)

        # Save progressive quality de < 100kb
        quality = 92
        while quality >= 20:
            canvas.save(out_path, "JPEG", quality=quality, optimize=True)
            if os.path.getsize(out_path) <= 100 * 1024:
                break
            quality -= 4

        size_kb = os.path.getsize(out_path) / 1024
        saved_files.append(out_path)
        print(f"Saved: {out_path} ({size_kb:.1f} KB, 500x500)")

        os.unlink(tmp.name)
    except Exception as e:
        print(f"Loi xu ly anh {url}: {e}")

# Cleanup logo temp
try:
    if aumi_logo:
        os.unlink(logo_tmp.name)
except Exception:
    pass

print("__RESULT__" + json.dumps(saved_files))
`;

  const pythonCmd = process.platform === "win32" ? "python" : "python3";
  const id = randomBytes(8).toString("hex");
  const tmpScript = join(tmpdir(), `sick_watermark_${id}.py`);
  try {
    writeFileSync(tmpScript, pythonCode, "utf8");
    const result = execFileSync(
      pythonCmd,
      [tmpScript, JSON.stringify(imageUrls), outputDir, productCode],
      { encoding: "utf8", timeout: 120000 }
    );
    process.stdout.write(result);

    // Parse saved file paths from output
    const match = result.match(/__RESULT__(.+)/);
    if (match) {
      return JSON.parse(match[1]);
    }
  } catch (err) {
    process.stderr.write(`Loi khi xu ly anh: ${err.message}\n`);
  } finally {
    try { unlinkSync(tmpScript); } catch { }
  }
  return [];
}

function writeXlsxWithPython(outputPath, templateData) {
  const pythonCode = `
import json
import sys
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

output_path = sys.argv[1]
with open(sys.argv[2], 'r', encoding='utf-8') as f:
    data = json.load(f)

wb = Workbook()
ws = wb.active
ws.title = "SICK Product"

title_fill = PatternFill(start_color="2F75B5", end_color="2F75B5", fill_type="solid")
sub_title_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
label_fill = PatternFill(start_color="D9E2F3", end_color="D9E2F3", fill_type="solid")
keyword_fill = PatternFill(start_color="FFFF00", end_color="FFFF00", fill_type="solid")
thin_border_color = "BFBFBF"

from openpyxl.styles import Border, Side
thin = Side(border_style="thin", color=thin_border_color)
border = Border(left=thin, right=thin, top=thin, bottom=thin)

# Row 1: Product header (B1:E1 merged)
ws.merge_cells("B1:E1")
ws["B1"] = data.get("productHeader", "")
ws["B1"].fill = title_fill
ws["B1"].font = Font(color="FFFFFF", bold=True)
ws["B1"].alignment = Alignment(horizontal="left", vertical="center")

# Row 2: Headers Tieng Viet | English
ws.merge_cells("B2:C2")
ws.merge_cells("D2:E2")
ws["B2"] = "Tiếng Việt"
ws["D2"] = "English"
for cell in ("B2", "D2"):
    ws[cell].fill = sub_title_fill
    ws[cell].font = Font(bold=True)
    ws[cell].alignment = Alignment(horizontal="center", vertical="center")

# Rows 3+: Data rows (label, vi_value, en_value)
content_rows = [
    ("Mã", data.get("codeVi", ""), data.get("codeEn", "")),
    ("Danh mục sp", data.get("productCategoryVi", ""), data.get("productCategoryEn", "")),
    ("Thương hiệu", data.get("brandVi", "SICK Sensor"), data.get("brandEn", "SICK Sensor")),
    ("URL", data.get("urlVi", ""), data.get("urlEn", "")),
    ("Keyword chính", data.get("keywordVi", ""), data.get("keywordEn", "")),
    ("Meta title", data.get("metaTitleVi", ""), data.get("metaTitleEn", "")),
    ("Meta description (155 - 165 characters)", data.get("metaDescriptionVi", ""), data.get("metaDescriptionEn", "")),
    ("Mô tả ngắn của sản phẩm", data.get("shortDescriptionVi", ""), data.get("shortDescriptionEn", "")),
    ("Mô tả thông số", data.get("longDescriptionVi", ""), data.get("longDescriptionEn", "")),
]

start_row = 3
for idx, item in enumerate(content_rows):
    row = start_row + idx
    label, vi_val, en_val = item
    ws.cell(row=row, column=1, value=label)
    ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
    ws.cell(row=row, column=2, value=vi_val)
    ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
    ws.cell(row=row, column=4, value=en_val)
    ws.cell(row=row, column=1).font = Font(bold=True)
    ws.cell(row=row, column=1).fill = label_fill
    # Highlight vang cho dong Keyword chinh
    if label == "Keyword chính":
        for col_idx in range(2, 6):
            ws.cell(row=row, column=col_idx).fill = keyword_fill

detail_label_row = start_row + len(content_rows)
ws.cell(row=detail_label_row, column=1, value="Thông số chi tiết")
ws.cell(row=detail_label_row, column=1).font = Font(bold=True)
ws.cell(row=detail_label_row, column=1).fill = label_fill
ws.cell(row=detail_label_row, column=2, value="Tên thông số")
ws.cell(row=detail_label_row, column=3, value="Giá trị")
ws.cell(row=detail_label_row, column=4, value="Spec name")
ws.cell(row=detail_label_row, column=5, value="Value")
for col in ("B", "C", "D", "E"):
    ws[f"{col}{detail_label_row}"].fill = sub_title_fill
    ws[f"{col}{detail_label_row}"].font = Font(bold=True)

spec_start = detail_label_row + 1
specs = data.get("detailSpecs", [])
if not specs:
    specs = [{"keyVi": "Khong co du lieu", "valueVi": "-", "keyEn": "No data", "valueEn": "-"}]

# Section header style: in dam + nen xam nhat (giong web SICK)
section_fill = PatternFill(start_color="EAEDF3", end_color="EAEDF3", fill_type="solid")
section_font = Font(bold=True)

for idx, spec in enumerate(specs):
    row = spec_start + idx
    is_section = spec.get("isSection", False)
    key_vi = spec.get("keyVi", spec.get("key", ""))
    val_vi = spec.get("valueVi", spec.get("value", ""))
    key_en = spec.get("keyEn", spec.get("key", ""))
    val_en = spec.get("valueEn", spec.get("value", ""))

    if is_section:
        # Section header: merge B:C va D:E, in dam, nen xam nhat
        ws.merge_cells(start_row=row, start_column=2, end_row=row, end_column=3)
        ws.cell(row=row, column=2, value=key_vi)
        ws.merge_cells(start_row=row, start_column=4, end_row=row, end_column=5)
        ws.cell(row=row, column=4, value=key_en)
        for col_idx in range(2, 6):
            ws.cell(row=row, column=col_idx).fill = section_fill
            ws.cell(row=row, column=col_idx).font = section_font
    else:
        ws.cell(row=row, column=2, value=key_vi)
        ws.cell(row=row, column=3, value=val_vi)
        ws.cell(row=row, column=4, value=key_en)
        ws.cell(row=row, column=5, value=val_en)

link_row = spec_start + len(specs)
ws.cell(row=link_row, column=1, value="Link datasheet")
ws.cell(row=link_row, column=1).font = Font(bold=True)
ws.cell(row=link_row, column=1).fill = label_fill

# B:C = VI datasheet, D:E = EN datasheet (backward-compat: dung datasheetUrl neu khong co rieng)
ds_vi = data.get("datasheetUrlVi", "") or data.get("datasheetUrl", "")
ds_en = data.get("datasheetUrlEn", "") or data.get("datasheetUrl", "")

ws.merge_cells(start_row=link_row, start_column=2, end_row=link_row, end_column=3)
ws.cell(row=link_row, column=2, value=ds_vi)
if ds_vi:
    ws.cell(row=link_row, column=2).hyperlink = ds_vi
    ws.cell(row=link_row, column=2).style = "Hyperlink"

ws.merge_cells(start_row=link_row, start_column=4, end_row=link_row, end_column=5)
ws.cell(row=link_row, column=4, value=ds_en)
if ds_en:
    ws.cell(row=link_row, column=4).hyperlink = ds_en
    ws.cell(row=link_row, column=4).style = "Hyperlink"

for row in ws.iter_rows(min_row=1, max_row=link_row, min_col=1, max_col=5):
    for cell in row:
        cell.border = border
        cell.alignment = Alignment(wrap_text=True, vertical="top")

ws.column_dimensions["A"].width = 32
ws.column_dimensions["B"].width = 30
ws.column_dimensions["C"].width = 52
ws.column_dimensions["D"].width = 30
ws.column_dimensions["E"].width = 52
ws.freeze_panes = "B3"
wb.save(output_path)
print(output_path)
`;

  // Strip large fields that Python script doesn't read (raw API response can be MB).
  // Saves disk I/O across thousands of products.
  const { raw, productImageUrls, _logoPath, ...pythonPayload } = templateData;
  const encodedTemplate = JSON.stringify(pythonPayload);
  const logoPath = _logoPath || "";
  const pythonCmd = process.platform === "win32" ? "python" : "python3";

  // Defensive: ensure the destination folder exists right before save.
  // (Antivirus/Windows Defender sometimes deletes recently-created empty folders
  // between crawlOneSeries' mkdir and the actual file write.)
  const { dirname } = require("node:path");
  const parentDir = dirname(outputPath);
  if (!existsSync(parentDir)) {
    mkdirSync(parentDir, { recursive: true });
  }

  // Use OS temp dir to avoid path-with-spaces issues, antivirus interference,
  // and conflicts with parallel runs. Random ID per call.
  const id = randomBytes(8).toString("hex");
  const tmpScript = join(tmpdir(), `sick_xlsx_${id}.py`);
  const tmpData = join(tmpdir(), `sick_xlsx_${id}.json`);

  writeFileSync(tmpScript, pythonCode, "utf8");
  writeFileSync(tmpData, encodedTemplate, "utf8");

  // Verify both files exist before invoking python (defensive against AV/race)
  if (!existsSync(tmpData)) {
    throw new Error(`Tmp JSON not visible after write: ${tmpData}`);
  }

  try {
    execFileSync(pythonCmd, [tmpScript, outputPath, tmpData, logoPath], {
      stdio: "inherit",
    });
  } finally {
    try { unlinkSync(tmpScript); } catch { }
    try { unlinkSync(tmpData); } catch { }
  }
}

async function processSingleProduct({ productUrl, apiUrl, outputPath, imgDir, logoPath, skipImages }) {
  const payload = await fetchJson(apiUrl);
  const templateData = await buildExcelTemplateData({ productUrl, apiUrl, payload });

  if (logoPath) {
    templateData._logoPath = logoPath;
  }

  writeXlsxWithPython(outputPath, templateData);
  process.stdout.write(`Da xuat: ${outputPath}\n`);

  if (!skipImages) {
    const productCode = templateData.productHeader || "sick-product";
    const partNumber = templateData.partNumber || "";
    const imageBaseName = buildImageBaseName(productCode, partNumber);
    const imageFileName = toSafeFilename(imageBaseName);
    const imageUrls = templateData.productImageUrls || [];
    if (imageUrls.length > 0) {
      downloadAndWatermarkImages(imageUrls, imgDir, imageFileName);
    }
  }
  return templateData.productHeader;
}


async function processCategoryUrl(categoryInput, outputDir, logoPath, skipImages) {
  const categoryCode = getCategoryCodeFromInput(categoryInput);
  if (!categoryCode) {
    throw new Error("Khong tim thay category code dang g123456.");
  }

  if (!existsSync(outputDir)) {
    mkdirSync(outputDir, { recursive: true });
  }

  const productIds = await fetchCategoryProductIds(categoryCode);
  process.stdout.write(`\nCategory ${categoryCode}: ${productIds.length} san pham.\n`);
  if (productIds.length === 0) return;

  let success = 0;
  let failed = 0;
  for (const [idx, productId] of productIds.entries()) {
    process.stdout.write(`[${idx + 1}/${productIds.length}] ${productId}\n`);
    try {
      const productUrl = buildProductUrlFromProductId(productId) || `https://www.sick.com/sg/en/p/${productId}`;
      const apiUrl = buildApiUrlFromProductId(productId);
      const payload = await fetchJson(apiUrl);
      const templateData = await buildExcelTemplateData({ productUrl, apiUrl, payload });
      const productCode = templateData.productHeader || productId || "sick-product";
      const partNumber = templateData.partNumber || "";
      const safeName = toSafeFilename(productCode) || productId;
      const outputPath = resolve(outputDir, `${safeName}.xlsx`);

      if (logoPath) templateData._logoPath = logoPath;
      writeXlsxWithPython(outputPath, templateData);
      process.stdout.write(`    -> ${outputPath}\n`);

      if (!skipImages && templateData.productImageUrls?.length) {
        const imageBaseName = buildImageBaseName(productCode, partNumber);
        const imageFileName = toSafeFilename(imageBaseName);
        downloadAndWatermarkImages(templateData.productImageUrls, outputDir, imageFileName);
      }
      success += 1;
    } catch (err) {
      process.stderr.write(`    Loi: ${err.message}\n`);
      failed += 1;
    }
  }

  process.stdout.write(`\nHoan tat category ${categoryCode}: ${success} thanh cong, ${failed} loi.\n`);
  process.stdout.write(`Thu muc xuat: ${outputDir}\n`);
  if (!skipImages) process.stdout.write("Anh duoc luu chung thu muc voi file Excel.\n");
}

async function main() {
  const cliArgs = parseArgs(process.argv.slice(2));

  if (cliArgs.help) {
    const usage = `
Su dung:
  node crawl-sick-product.js [options]

Che do 1 san pham:
  --url <productUrl>     URL trang san pham SICK
  --api <apiUrl>         URL API truc tiep (bo qua --url)
  --output <file.xlsx>   File Excel dau ra (mac dinh: sick-product.xlsx)

Che do category:
  --category-url <url>   URL trang category (vd: https://www.sick.com/.../c/g433452?tab=selection)
  --out-dir <folder>     Thu muc luu Excel dau ra (mac dinh: ./sick-excel/<categoryCode>)

Tuy chon chung:
  --logo <logo.png>      File logo local (mac dinh: tai logo AUMI tu web)
  --img-dir <folder>     Thu muc luu anh san pham (mac dinh: ./sick-images)
  --no-images            Bo qua buoc tai va xu ly anh

Vi du:
  node crawl-sick-product.js --url "${DEFAULT_PRODUCT_URL}"
  node crawl-sick-product.js --category-url "https://www.sick.com/sg/en/catalog/products/detection-sensors/photoelectric-sensors/w16/c/g433452?tab=selection" --out-dir ./w16
`;
    process.stdout.write(usage);
    return;
  }

  const imgDir = resolve(cliArgs["img-dir"] ?? "sick-images");
  const logoPath = cliArgs.logo ? resolve(String(cliArgs.logo)) : null;
  const skipImages = !!cliArgs["no-images"];

  // Che do category
  const categoryInput = cliArgs["category-url"] ?? cliArgs.category ?? null;
  if (categoryInput || (cliArgs.url && isCategoryUrl(cliArgs.url) && !cliArgs.api)) {
    const picked = categoryInput ?? cliArgs.url;
    const categoryCode = getCategoryCodeFromInput(picked);
    const outDir = resolve(cliArgs["out-dir"] ?? join("sick-excel", categoryCode || "category"));
    await processCategoryUrl(picked, outDir, logoPath, skipImages);
    return;
  }

  // Che do 1 san pham
  const productUrl = cliArgs.url ?? DEFAULT_PRODUCT_URL;
  const apiUrl = cliArgs.api ?? buildApiUrlFromProductUrl(productUrl) ?? DEFAULT_API_URL;
  const outputPath = resolve(cliArgs.output ?? DEFAULT_OUTPUT_FILE);

  const payload = await fetchJson(apiUrl);
  const templateData = await buildExcelTemplateData({ productUrl, apiUrl, payload });

  if (cliArgs["save-json"]) {
    writeFileSync(resolve(String(cliArgs["save-json"])), JSON.stringify(payload, null, 2), "utf8");
  }

  if (logoPath) templateData._logoPath = logoPath;

  writeXlsxWithPython(outputPath, templateData);
  process.stdout.write(`\nDa xuat file Excel: ${outputPath}\n`);

  if (!skipImages) {
    const productCode = templateData.productHeader || "sick-product";
    const partNumber = templateData.partNumber || "";
    const imageBaseName = buildImageBaseName(productCode, partNumber);
    const imageFileName = toSafeFilename(imageBaseName);
    const imageUrls = templateData.productImageUrls || [];
    if (imageUrls.length > 0) {
      process.stdout.write(`\nDang tai anh san pham va gan logo AUMI...\n`);
      const savedFiles = downloadAndWatermarkImages(imageUrls, imgDir, imageFileName);
      process.stdout.write(`Hoan tat: ${savedFiles.length} anh da luu vao ${imgDir}\n`);
    } else {
      process.stdout.write("\nKhong tim thay anh san pham trong API response.\n");
    }
  }
}

// Run main only when executed directly
if (require.main === module) {
  main().catch((error) => {
    process.stderr.write(`\nLoi: ${error.message}\n`);
    process.exit(1);
  });
}

module.exports = {
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
  getCategoryCodeFromInput,
};
