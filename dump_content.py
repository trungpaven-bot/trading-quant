
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb["TT09"]

print("--- Dumping First 10 Rows, 15 Columns ---")
for r in range(1, 11):
    row_vals = []
    for c in range(1, 16):
        cell = ws.cell(row=r, column=c)
        val = cell.value
        if val:
            row_vals.append(f"({r},{c}): {repr(val)}")
    if row_vals:
        print(" | ".join(row_vals))
