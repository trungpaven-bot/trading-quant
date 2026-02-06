
import pandas as pd

excel_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Du_lieu_Suat_von_dau_tu_2024.xlsx"
df = pd.read_excel(excel_path, header=None)

print("Shape:", df.shape)
print("First 50 rows preview:")
print(df.head(50).to_string())

# Find rows starting with "--- BẢNG" to identify table breaks
table_starts = df[df[0].str.contains("--- BẢNG", na=False)].index.tolist()
print("\nTable Start Indices (First 10):", table_starts[:10])

# Inspect a specific table (e.g. Table 4 which seemed large in previous logs)
if len(table_starts) >= 4:
    start = table_starts[3] # Table 4
    end = table_starts[4] if len(table_starts) > 4 else start + 20
    print(f"\n--- Content of Table 4 (Rows {start}-{end}) ---")
    print(df.iloc[start:end].to_string())
