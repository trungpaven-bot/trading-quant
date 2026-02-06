
import openpyxl

file_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Tra dinh muc TT09.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True) # Load Value
ws = wb["TT09"]

print("--- Searching 'Thẩm định' again with VALUES to catch headers ---")
# The previous search missed some rows because maybe formulas return empty string?
# Or maybe the text is in merged cells.

for r in range(500, 1000):
    val = ws.cell(row=r, column=2).value
    if val and isinstance(val, str) and ("thẩm định" in val.lower()):
        print(f"Row {r}: {val}")
