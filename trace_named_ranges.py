
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)

print("--- Checking Named Range 'i_CPC' ---")
# Usually defined in workbook names
if 'i_CPC' in wb.defined_names:
    defn = wb.defined_names['i_CPC']
    print(f"Defined Name i_CPC: {list(defn.destinations)}")
else:
    print("i_CPC not found in defined names.")

print("--- Dumping Names that might be variables ---")
# Look for names like i_LoaiCongTrinh, i_CapCongTrinh etc.
for name in wb.defined_names:
    if str(name).startswith("i_") or str(name).startswith("v_"):
        print(f"{name}: {list(wb.defined_names[name].destinations)}")
