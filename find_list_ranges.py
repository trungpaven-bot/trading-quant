
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True)
ws = wb["TT09"]

print("--- Searching for specific list items ---")
targets = ["1 bước", "2 bước", "3 bước", "độc lập", "nội bộ", "vốn nhà nước", "vốn khác"]

for r in range(1, 150):
    for c in range(1, 20):
        cell = ws.cell(row=r, column=c)
        val = str(cell.value).lower() if cell.value else ""
        for t in targets:
            if t in val:
                print(f"Found '{t}' in '{val}' at {cell.coordinate}")
