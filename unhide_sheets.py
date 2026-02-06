
import openpyxl

file_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\File_Tra_Cuu_Dep_Va_Thong_Minh.xlsx"
wb = openpyxl.load_workbook(file_path)

# Unhide all sheets so user can reference them manually
for sheet in wb.sheetnames:
    wb[sheet].sheet_state = 'visible'

print("Unhiding sheets...")
wb.save(file_path)
print("Done.")
