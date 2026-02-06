
import pandas as pd

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
df = pd.read_excel(file_path, sheet_name="TT09", header=None)
print(df.iloc[20:60, 0:10].to_string())
