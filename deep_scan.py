
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb["TT09"]

print("--- Deep Scan for Labels (Rows 1-20, Cols 1-30) ---")
# The labels might be slightly different than my dictionary or have merged cells.
for r in range(1, 20):
    for c in range(1, 30):
        cell = ws.cell(row=r, column=c)
        val = cell.value
        if val and isinstance(val, str) and len(val) > 3:
             print(f"({r},{c}) [{cell.coordinate}]: {val}")
