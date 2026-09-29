"""
Download Product Overview PDF for each product family from SICK CDN.
Uses FactFinder API to find download entry, then downloads from CDN.
"""
import requests
import json
import os
import time
import re

BASE_DIR = "SICK_Products/SICK_Products"
API_URL = "https://www.sick.com/api/fact-finder/search/enSG-sick"
CDN_URL = "https://cdn.sick.com/media"

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept": "application/json",
}

def load_family_mapping():
    """Build mapping from folder name to family code."""
    with open("sick_products_raw.json", "r", encoding="utf-8") as f:
        data = json.load(f)

    name_to_code = {}
    for cat in data:
        for sub in cat.get("subcategories", []):
            for fam in sub.get("families", []):
                name = fam["name"]
                code = fam["code"]
                name_to_code[name.replace(" ", "_")] = code
                name_to_code[name] = code
                # Also handle special chars: comma → ,
                cleaned = re.sub(r'[,/]', '_', name).replace(" ", "_")
                name_to_code[cleaned] = code

    return name_to_code

def get_product_overview_path(session, family_code):
    """Query FactFinder for the product overview download entry."""
    download_id = f"dg{family_code[1:]}" if family_code.startswith("g") else f"d{family_code}"

    resp = session.get(
        f"{API_URL}?query=*&filter=DefArticleNo:{download_id}&hitsPerPage=1",
        timeout=30
    )
    if resp.status_code != 200:
        return None

    data = resp.json()
    hits = data.get("hits", [])
    if not hits:
        return None

    mv = hits[0].get("masterValues", {})
    download_json = mv.get("DownloadJson", [])
    if isinstance(download_json, str):
        download_json = json.loads(download_json)

    for d in download_json:
        if d.get("Language") == "en":
            return d.get("Path")

    return None

def download_pdf(session, cdn_path, save_path):
    """Download PDF from CDN."""
    url = f"{CDN_URL}{cdn_path}"
    resp = session.get(url, timeout=60, stream=True)
    if resp.status_code == 200 and resp.content[:4] == b"%PDF":
        os.makedirs(os.path.dirname(save_path), exist_ok=True)
        with open(save_path, "wb") as f:
            f.write(resp.content)
        return True, len(resp.content)
    return False, resp.status_code

def main():
    name_to_code = load_family_mapping()
    folders = sorted(f for f in os.listdir(BASE_DIR) if os.path.isdir(os.path.join(BASE_DIR, f)))

    print(f"Total folders: {len(folders)}")

    session = requests.Session()
    session.headers.update(HEADERS)

    downloaded = 0
    no_overview = []
    failed = []
    skipped = 0

    for i, folder in enumerate(folders, 1):
        datasheet_dir = os.path.join(BASE_DIR, folder, "datasheet")
        existing_pdfs = []
        if os.path.exists(datasheet_dir):
            existing_pdfs = [f for f in os.listdir(datasheet_dir) if f.endswith(".pdf")]
        if existing_pdfs:
            skipped += 1
            if i % 50 == 0:
                print(f"[{i}/{len(folders)}] Skipped (already exists): {folder}")
            continue

        # Find family code
        code = name_to_code.get(folder) or name_to_code.get(folder.replace("_", " "))
        if not code:
            failed.append(f"{folder}: no family code mapping")
            continue

        # Get product overview path from API
        try:
            cdn_path = get_product_overview_path(session, code)
            time.sleep(0.5)
        except Exception as e:
            failed.append(f"{folder}: API error - {e}")
            time.sleep(2)
            continue

        if not cdn_path:
            no_overview.append(folder)
            if i % 50 == 0:
                print(f"[{i}/{len(folders)}] No overview: {folder}")
            continue

        # Download PDF
        filename = os.path.basename(cdn_path)
        save_path = os.path.join(datasheet_dir, filename)

        try:
            success, result = download_pdf(session, cdn_path, save_path)
            time.sleep(0.3)
            if success:
                downloaded += 1
                if downloaded % 20 == 0 or i % 50 == 0:
                    print(f"[{i}/{len(folders)}] Downloaded: {folder} ({result} bytes)")
            else:
                failed.append(f"{folder}: download failed (status {result})")
        except Exception as e:
            failed.append(f"{folder}: download error - {e}")
            time.sleep(2)

    print(f"\n{'='*60}")
    print(f"RESULTS")
    print(f"{'='*60}")
    print(f"Total folders: {len(folders)}")
    print(f"Downloaded: {downloaded}")
    print(f"Skipped (already exists): {skipped}")
    print(f"No product overview: {len(no_overview)}")
    print(f"Failed: {len(failed)}")

    if no_overview:
        print(f"\nNo product overview ({len(no_overview)}):")
        for n in no_overview[:30]:
            print(f"  {n}")

    if failed:
        print(f"\nFailed ({len(failed)}):")
        for f in failed[:30]:
            print(f"  {f}")

if __name__ == "__main__":
    main()
