"""
Fix remaining 60 short description translation errors.
Extracts term-level changes from vi_value vs suggested_fix and applies them.
"""
import openpyxl
import unicodedata

ERROR_FILE = "loi dich san pham.xlsx"
PATH_PREFIX = "SICK_Products/"

def normalize(s):
    if s is None: return ""
    return unicodedata.normalize("NFKC", str(s)).strip()

def main():
    wb = openpyxl.load_workbook(ERROR_FILE)
    ws = wb.active

    fixed = 0
    errors = 0

    for row in range(2, ws.max_row + 1):
        field = ws.cell(row, 4).value
        if field != "Mô tả ngắn của sản phẩm":
            continue

        xlsx_path = PATH_PREFIX + ws.cell(row, 3).value
        vi_value = ws.cell(row, 6).value
        suggested_fix = ws.cell(row, 8).value

        if not vi_value or not suggested_fix:
            continue

        # Extract term-level replacements by comparing parts
        vi_parts = vi_value.split("; ")
        fix_parts = suggested_fix.split("; ")
        replacements = []
        for v, f in zip(vi_parts, fix_parts):
            if v != f:
                replacements.append((v, f))

        if not replacements:
            continue

        try:
            wb2 = openpyxl.load_workbook(xlsx_path)
            ws2 = wb2.active
            short_desc = ws2.cell(10, 2).value
            if not short_desc:
                wb2.close()
                continue

            original = str(short_desc)
            current = original
            for old, new in replacements:
                # Normalize for comparison but apply to original
                current_norm = normalize(current)
                old_norm = normalize(old)
                if old_norm in current_norm:
                    # Find position in normalized and replace
                    current = normalize(current).replace(old_norm, new)

            if current != normalize(original):
                ws2.cell(10, 2).value = current
                wb2.save(xlsx_path)
                fixed += 1

            wb2.close()
        except Exception as e:
            errors += 1
            print(f"Error processing {xlsx_path}: {e}")

    wb.close()
    print(f"Fixed: {fixed}, Errors: {errors}")

if __name__ == "__main__":
    main()
