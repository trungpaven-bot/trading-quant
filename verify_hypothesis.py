
import openpyxl

file_path = r"g:\My Drive\00_INBOX (xu ly nhanh)\tra TMDT.xlsx"
wb = openpyxl.load_workbook(file_path, data_only=False) # Get Formulas

ws_tt09 = wb["TT09"]
print("--- Check F58 Formula in TT09 ---")
print(f"F58: {ws_tt09['F58'].value}")
# Likely =i_CPC

print("--- Check Definition of i_CPC again ---")
if 'i_CPC' in wb.defined_names:
    defn = wb.defined_names['i_CPC']
    print(f"i_CPC Define: {list(defn.destinations)}")
else:
    print("i_CPC not found")

# Is it possible [2]Config refers to a DIFFERENT sheet "Config" that is hidden or VeryHidden?
print("\n--- Sheet Names ---")
print(wb.sheetnames)

# Or maybe it refers to TMDT but the cell is blank because it EXPECTS the input FROM the Control?
# Correct!
# The Form Control writes TO the cell. If the Control is broken (unlink), the cell stays blank.
# The user sees "None" because the link is broken.
# BUT the User needs to know which cell to LINK TO so that formulas work.

# Let's verify if F58 (=i_CPC) evaluates to something if I temporarily WRITE a value to TMDT!C118.
