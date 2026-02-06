
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
ws_tmdt = wb["TMDT"]

print("--- Checking formula usage in TMDT sheet ---")
# Check if variables like i_CPC are used in TMDT sheet directly
variables = ["i_CPC", "i_CapCT", "TT09"] 
found = False
for r in range(1, 200):
    for c in range(1, 15):
        val = str(ws_tmdt.cell(row=r, column=c).value)
        if "=" in val:
            for v in variables:
                if v in val:
                    print(f"Row {r} Col {c}: {val}")
                    found = True
                    break
    if found and r > 150: break # Just a sample
