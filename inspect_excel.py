import pandas as pd
import sys

file_path = 'g:/My Drive/30_RESOURCES (Kho tai nguyen)/30.01_Tinh hieu qua du an/02_Tra dinh muc theo TT 09 va 12/2026.01.19 Tien do du an CCN PDP 1 - Chi phi HĐ V1.0_T.xlsx'

try:
    xl = pd.ExcelFile(file_path)
    for sheet in xl.sheet_names:
        print(f"\n{'='*50}")
        print(f"SHEET: {sheet}")
        print(f"{'='*50}")
        df = xl.parse(sheet, header=None)
        # Print first 20 rows and all columns
        print(df.head(30).to_string())
except Exception as e:
    print(f"Error: {e}")
