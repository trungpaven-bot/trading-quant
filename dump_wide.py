
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb["TT09"]

print("--- Dumping Row 4 and 5, Cols 1-30 ---")
for r in [4, 5]:
    row_vals = []
    for c in range(1, 31):
        cell = ws.cell(row=r, column=c)
        val = cell.value
        if val:
            row_vals.append(f"({r},{c}): {repr(val)}")
    print(f"Row {r}: " + " | ".join(row_vals))
