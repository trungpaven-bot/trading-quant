
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False)

print("--- Inspecting Named Range 'i_CPC' ---")
if 'i_CPC' in wb.defined_names:
    defn = wb.defined_names['i_CPC']
    print(f"Raw Entry: {defn}")
    # Destinations splits the reference
    for title, coord in defn.destinations:
        print(f"Sheet Title: {title}")
        print(f"Coordinate: {coord}")
        
    # Check sheet index mapping
    sheets = wb.sheetnames
    if len(sheets) > 2:
        print(f"Sheet at Index 2 is: {sheets[2]}") # 0, 1, 2
else:
    print("i_CPC not found")
