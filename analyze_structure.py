
import pandas as pd

file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1.xlsx'
sheet_name = 'PA 3 (ok)'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Iterate through rows and print non-null values with their indices
    print("Scanning first 50 rows for structure:")
    for i in range(50):
        row_data = df.iloc[i].dropna()
        if not row_data.empty:
            print(f"Row {i}: {row_data.to_dict()}")

except Exception as e:
    print(f"Error: {e}")
