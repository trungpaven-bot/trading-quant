
import pandas as pd
import numpy as np

file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1.xlsx'
sheet_name = 'PA 3 (ok)'

df = pd.read_excel(file_path, sheet_name=sheet_name, header=None)

# Helper to clean and float
def to_float(x):
    try:
        return float(x)
    except:
        return 0.0

# 1. Total Investment Check (Rows 8-14 approx, indices 7-13)
# Row 7 is Total. Rows 8-13 are components.
# Indices:
# Row 7 (index 7): Total Investment
# Rows 8-13 (indices 8-13): Components

print("--- LOGIC CHECK 1: Investment Costs ---")
total_inv_declared = to_float(df.iloc[7, 2]) # Column C (index 2)
components_sum = 0
components = []
for i in range(8, 14): # Rows 8 to 13 (indices) -> Excel rows 9-14
    val = to_float(df.iloc[i, 2])
    name = df.iloc[i, 1]
    components.append((name, val))
    components_sum += val

print(f"Declared Total: {total_inv_declared:,.2f}")
print(f"Sum of Components: {components_sum:,.2f}")
print(f"Difference: {total_inv_declared - components_sum:,.2f}")
if abs(total_inv_declared - components_sum) > 1:
    print("WARNING: Components do not sum to Total Investment!")
else:
    print("OK: Investment costs match.")

# 2. Area Distribution Check
# Row 65 (index 64): Revenue from Industrial Land
# Row 66 (index 65): "Diện tích đất công nghiệp thu hút đầu tư hằng năm"
# Total Area is in Row 33 (index 33), value in col 3? Let's check dump.
# Dump: Row 33 col 2 (C) is 544049.

print("\n--- LOGIC CHECK 2: Land Area Distribution ---")
total_area_declared = to_float(df.iloc[33, 2]) # Index 33 corresponding to row 34? "Diện tích đất kinh doanh"
if total_area_declared == 0:
    # Try row 32 or 34
    total_area_declared = to_float(df.iloc[32, 2])

# The yearly distribution is likely in Row 65 (index 65 of df? No, explicit index 65 in dump is row 66)
# Let's find the row with text "Diện tích đất công nghiệp thu hút đầu tư hằng năm"
row_idx = -1
for i in range(50, 80):
    val = str(df.iloc[i, 1])
    if "Diện tích đất công nghiệp" in val and "hằng năm" in val:
        row_idx = i
        break

if row_idx != -1:
    print(f"Found Area Distribution at Row check {row_idx} (Excel Row {row_idx+1})")
    # Columns roughly 6 onwards (G onwards)
    yearly_areas = df.iloc[row_idx, 6:].apply(to_float)
    total_distributed = yearly_areas.sum()
    
    print(f"Total Area Declared: {total_area_declared:,.2f}")
    print(f"Total Area Distributed: {total_distributed:,.2f}")
    print(f"Difference: {total_area_declared - total_distributed:,.2f}")
    
    if abs(total_area_declared - total_distributed) > 1:
         print("WARNING: Distributed area does not match Total Area!")
    else:
         print("OK: Area distribution matches.")
else:
    print("Could not find Area Distribution row.")

# 3. Check Table Numbering
print("\n--- LOGIC CHECK 3: Table Numbering ---")
# Scan first column for "Bảng"
duplicates = {}
for i in range(df.shape[0]):
    val = str(df.iloc[i, 0])
    if "Bảng" in val:
        print(f"Found table at Row {i}: {val}")
        
