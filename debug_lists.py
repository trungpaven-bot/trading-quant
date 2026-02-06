
import openpyxl

file_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\File_Tra_Cuu_Final_SmartFix_NoBlanks.xlsx"

print(f"Loading {file_path}...")
try:
    wb = openpyxl.load_workbook(file_path)
    ws = wb["LISTS"]
    
    # Read Col A
    values = []
    for cell in ws["A"]:
        val = cell.value
        if val is not None and str(val).strip() != "":
            values.append(str(val).strip())
            
    print(f"Found {len(values)} non-empty values in Col A.")
    print("First 5:", values[:5])
    print("Last 5:", values[-5:])
    
    # Check what was in the named range
    if "ListTables" in wb.defined_names:
        print("Existing ListTables definition:", wb.defined_names["ListTables"].attr_text)

except Exception as e:
    print(f"Error: {e}")
