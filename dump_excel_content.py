
import pandas as pd
import openpyxl

file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1.xlsx'
sheet_name = 'PA 3 (ok)'

# 1. Dump Values
try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    with open('sheet_dump_values.txt', 'w', encoding='utf-8') as f:
        f.write(df.to_string())
    print("Values dumped to sheet_dump_values.txt")
except Exception as e:
    print(f"Error reading values: {e}")

# 2. Dump Formulas
try:
    wb = openpyxl.load_workbook(file_path, data_only=False)
    sheet = wb[sheet_name]
    with open('sheet_dump_formulas.txt', 'w', encoding='utf-8') as f:
        for row in sheet.iter_rows(values_only=False):
            row_data = []
            for cell in row:
                # formatting to keep it compact
                val = cell.value
                if val:
                    # check if formula
                    if isinstance(val, str) and val.startswith('='):
                        row_data.append(f"[{cell.coordinate}: {val}]")
                    else:
                        row_data.append(f"[{cell.coordinate}: VAL]") # Just marking presence of value
            if row_data:
                f.write(", ".join(row_data) + "\n")
    print("Formulas dumped to sheet_dump_formulas.txt")
except Exception as e:
    print(f"Error reading formulas: {e}")
