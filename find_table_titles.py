
import pandas as pd

excel_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Du_lieu_Suat_von_dau_tu_2024.xlsx"
df = pd.read_excel(excel_path, header=None)

# Search for "Bảng" in the dataframe to see if titles were captured inside tables or if I need to fetch them from doc
print("Searching for 'Bảng' keyword in first 500 rows...")
found_titles = []
for idx, row in df.iterrows():
    row_str = " | ".join([str(x) for x in row if not pd.isna(x)])
    if "Bảng" in row_str and len(row_str) < 200: # Heuristic for title
        found_titles.append(f"Row {idx}: {row_str}")

for t in found_titles[:20]:
    print(t)
    
print(f"Total potential titles found: {len(found_titles)}")
