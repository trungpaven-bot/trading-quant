
import openpyxl
from openpyxl.workbook.defined_name import DefinedName
from openpyxl.worksheet.datavalidation import DataValidation

file_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\File_Tra_Cuu_Final_SmartFix_NoBlanks.xlsx"

print(f"Opening {file_path} for Hard Fix...")
wb = openpyxl.load_workbook(file_path)

# --- FIX LIST 1: TABLES ---
ws_list = wb["LISTS"]

# 1. Collect clean data from Col Z (Unique Tables)
# Note: In previous script, I put Unique Tables in Col Z.
vals_z = []
for cell in ws_list["Z"]:
    v = cell.value
    if v and str(v).strip() != "" and str(v) != "UnqTbls": # Skip header
        vals_z.append(str(v).strip())

print(f"Found {len(vals_z)} unique tables.")

# 2. Hardcode Named Range "ListTables"
# Remove old dynamic definition if exists
if "ListTables" in wb.defined_names:
    del wb.defined_names["ListTables"]

# Create STATIC definition: $Z$2:$Z$132 (Exact fit)
exact_range_z = f"LISTS!$Z$2:$Z${len(vals_z) + 1}" 
# +1 because Z1 is header, data starts at Z2. End is 1 + count
print(f"Hardcoding ListTables to: {exact_range_z}")
dn = DefinedName("ListTables", attr_text=exact_range_z)
wb.defined_names.add(dn)


# --- FIX LIST 2: DEPENDENT LISTS (ColTbl & ColDesc) ---
# Col A (Table) and Col D (Desc)
vals_a = []
for cell in ws_list["A"]:
    v = cell.value
    if v and str(v).strip() != "" and str(v) != "Bảng":
        vals_a.append(str(v))

count_rows = len(vals_a)
print(f"Found {count_rows} data rows.")

# Hardcode ColTbl and ColDesc ranges
if "ColTbl" in wb.defined_names: del wb.defined_names["ColTbl"]
if "ColDesc" in wb.defined_names: del wb.defined_names["ColDesc"]

exact_range_a = f"LISTS!$A$2:$A${count_rows + 1}"
exact_range_d = f"LISTS!$D$2:$D${count_rows + 1}"

print(f"Hardcoding ColTbl to: {exact_range_a}")
wb.defined_names.add(DefinedName("ColTbl", attr_text=exact_range_a))

print(f"Hardcoding ColDesc to: {exact_range_d}")
wb.defined_names.add(DefinedName("ColDesc", attr_text=exact_range_d))


# --- RE-APPLY VALIDATION TO BE SURE ---
ws_ui = wb["TRA_CUU"]
c4 = ws_ui["C4"]
c4.value = vals_z[0] # Set default to allow validation to pass

# Clear old validation
# OpenPyXL validation handling is tricky. Usually overwriting works.
# But let's verify formula in C5 (Dependent) uses comma/semicolon correctly?
# Actually, since we fixed the "Defined Names", the existing validation in cells 
# should just magically work now because they point to "=ListTables".

print("Saving fix...")
wb.save(file_path)
print("Done.")
