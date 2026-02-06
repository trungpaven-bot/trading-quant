
import pandas as pd

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
sheet_name = "TT09"

df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)
# pandas is 0-indexed, so row 90 is index 89.
subset = df.iloc[89:110] 

print(subset.to_string())
