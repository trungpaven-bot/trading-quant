
import pandas as pd

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
df = pd.read_excel(file_path, sheet_name="TT09", header=None)

print("--- Data in TT09 rows 105-115 (Lookup Table Keys) ---")
print(df.iloc[105:116, 2:8].to_string()) # Cols C to H (indices 2 to 7)

print("\n--- Rows 1-10 to find 'Loại công trình' label again ---")
# Maybe it is merged?
print(df.iloc[0:10, 0:15].to_string())
