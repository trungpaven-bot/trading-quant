
import pandas as pd

file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\2026.01.17 Tien do du an CCN PDP 2 - Chi phi HĐ V1.0.xlsx'

try:
    # Load with openpyxl to check colors
    import openpyxl
    wb = openpyxl.load_workbook(file_path)
    ws = wb['Tien do']
    
    print("Row 4 Headers:")
    for col in range(1, 25):
        val = ws.cell(row=4, column=col).value
        print(f"Col {col}: {val}")
    
except Exception as e:
    print(f"Error reading file: {e}")
