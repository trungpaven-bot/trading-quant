
import pandas as pd

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
df = pd.read_excel(file_path, sheet_name="TT09", header=None)

print("--- Rows 190-200 (i_CPC Usage) ---")
print(df.iloc[189:200, 0:5].to_string()) 

print("\n--- Rows 330-340 (i_CapCT Usage) ---")
print(df.iloc[329:340, 0:5].to_string())
