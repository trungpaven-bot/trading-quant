
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)

print("--- Defined Names ---")
# List all names to find candidates for "KieuCongTrinh" and "KiemToan"
for name in wb.defined_names:
    defn = wb.defined_names[name]
    # Filter for names that look like variables (start with i_, or short codes)
    if str(name).startswith("i_") or str(name) in ["KT", "GiamSat", "Kieu", "Loai"]:
        dest = list(defn.destinations)
        print(f"{name}: {dest}")

print("\n--- Interpreting [2]Config ---")
# Check sheets to see which one maps to index 2 or has 'Config' in a hidden name
print(f"Sheet names: {wb.sheetnames}")
