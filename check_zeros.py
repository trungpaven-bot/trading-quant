
import openpyxl

file_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Tra dinh muc TT09.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=True) # Load values to see where 0s are

print("--- Checking Previous Broken Link Locations ---")
# Based on logs from Step 495
check_list = [
    ("CHI_TIEU", "B12"), 
    ("TT09", "D942"),
    ("TT09", "E1192")
]

for sheet, coord in check_list:
    if sheet in wb.sheetnames:
        val = wb[sheet][coord].value
        print(f"{sheet}!{coord} Value: {val}")
