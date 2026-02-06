
import pandas as pd
import os

# Define file path
file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1.xlsx'

# Load the specific sheet
sheet_name = 'PA 3 (ok)'

try:
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
    
    # Print the first 20 rows and ALL columns to get a good sense of the layout
    print("Shape:", df.shape)
    print("\nFirst 20 rows:")
    print(df.head(20).to_string())
    
    # Also verify if there are any other sheets that might be relevant or if the name is slightly different if it fails (unlikely given previous success finding file)
except Exception as e:
    print(f"Error reading excel: {e}")
