
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)
ws = wb["TT09"]

print("--- Checking Formulas in Rows 60-100 ---")
# Focus on Column E (percentage?) and H (Value?) 
# Based on earlier "dump_wide", data seemed to be in Cols 4-10 maybe?

for r in range(60, 100):
    row_data = []
    # Check Columns B to H
    for c_idx, c_char in enumerate(['B', 'C', 'D', 'E', 'F', 'G', 'H']):
        cell = ws[f"{c_char}{r}"]
        val = cell.value
        if isinstance(val, str) and val.startswith("="):
             row_data.append(f"{c_char}={val}")
        elif val:
             # Shorten text
             v_str = str(val)
             if len(v_str) > 20: v_str = v_str[:20] + "..."
             row_data.append(f"{c_char}={v_str}")
    
    if row_data:
         print(f"Row {r}: " + " | ".join(row_data))
