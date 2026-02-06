
import win32com.client as win32
import os

# Define paths
file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1.xlsx'
output_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1_Professional.xlsx'

excel = None
wb = None

try:
    print("Launching Excel...")
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False

    print(f"Opening {file_path}...")
    wb = excel.Workbooks.Open(file_path)
    
    # Identify source sheet
    source_name = "PA 3 (ok)"
    try:
        ws_source = wb.Worksheets(source_name)
    except:
        print(f"Sheet {source_name} not found!")
        wb.Close(SaveChanges=False)
        excel.Quit()
        exit()

    # Create a copy
    print("Creating new sheet 'PA 3 (Professional)'...")
    ws_source.Copy(Before=ws_source)
    ws_new = wb.ActiveSheet
    ws_new.Name = "PA 3 (Professional)"
    
    # --- 1. PREP SPACE ---
    # Find start of Table 1.2 (Allocations) to insert space before it.
    print("Finding insertion point...")
    found_alloc = ws_new.Columns(1).Find(What="Bảng 1.2", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
    
    if found_alloc:
        insert_row = found_alloc.Row
        print(f"Inserting 50 rows at Row {insert_row}...")
        ws_new.Rows(f"{insert_row}:{insert_row+49}").Insert()
        
        target_row_op = insert_row + 2
        target_row_price = insert_row + 25
    else:
        print("Could not find 'Bảng 1.2'. Defaulting to Row 50.")
        insert_row = 50
        ws_new.Rows("50:100").Insert()
        target_row_op = 52
        target_row_price = 75

    # --- 2. MOVE OPERATING COSTS (V) ---
    # Search for header "Nhóm thông tin về chi phí hoạt động"
    print("Moving Operating Costs Block...")
    found_op = ws_new.Cells.Find(What="Nhóm thông tin về chi phí hoạt động", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
    
    if found_op:
        # Define block range: F4 to K20 (Approx)
        # We start from found_op row/col.
        r_start = found_op.Row
        c_start = found_op.Column
        # Let's assume block is ~15 rows high, 6 cols wide.
        # Check specific bounds if possible, but raw block cut is usually safe if space is clear.
        block_op = ws_new.Range(ws_new.Cells(r_start, c_start), ws_new.Cells(r_start+16, c_start+7)) # Generous box
        
        dest_op = ws_new.Cells(target_row_op, 2) # Column B
        block_op.Cut(Destination=dest_op)
        
        # Add Header/Style
        ws_new.Cells(target_row_op-1, 2).Value = "V. THÔNG SỐ CHI PHÍ HOẠT ĐỘNG (Đã chuyển)"
        ws_new.Cells(target_row_op-1, 2).Font.Bold = True
        ws_new.Cells(target_row_op-1, 2).Font.Color = 255 # Red

    # --- 3. MOVE UNIT PRICES (VII) ---
    # Search for header "Nhóm thông tin về đơn giá kinh doanh"
    print("Moving Unit Prices Block...")
    found_price = ws_new.Cells.Find(What="Nhóm thông tin về đơn giá kinh doanh", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
    
    if found_price:
        r_start = found_price.Row
        c_start = found_price.Column
        # This block is likely below the first one, or in Col H.
        block_price = ws_new.Range(ws_new.Cells(r_start, c_start), ws_new.Cells(r_start+16, c_start+6))
        
        dest_price = ws_new.Cells(target_row_price, 2)
        block_price.Cut(Destination=dest_price)
        
        ws_new.Cells(target_row_price-1, 2).Value = "VI. THÔNG SỐ ĐƠN GIÁ KINH DOANH (Đã chuyển)"
        ws_new.Cells(target_row_price-1, 2).Font.Bold = True
        ws_new.Cells(target_row_price-1, 2).Font.Color = 255
        
    # --- 4. FIX TABLE NUMBERING ---
    # Find "Bảng 1.2 : Bảng tính doanh thu" (The second 1.2)
    # Since we inserted rows, it's pushed down.
    # Start search from Row 100 downward.
    print("Fixing table numbering...")
    found_rev = ws_new.Range("A100:E200").Find(What="Bảng tính doanh thu", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
    
    if found_rev:
        val = found_rev.Value
        if "1.2" in val:
            print(f"Renaming {val} to 1.3")
            found_rev.Value = val.replace("1.2", "1.3")
            
    # --- 5. FORMATTING TOUCHUPS ---
    # Set Column B width to be wider for readability
    ws_new.Columns("B").ColumnWidth = 50
    ws_new.Columns("C").ColumnWidth = 15
    ws_new.Columns("D").ColumnWidth = 10
    
    # --- 6. SAVE ---
    print(f"Saving to {output_path}...")
    wb.SaveAs(output_path)
    print("Done.")

except Exception as e:
    print(f"Error: {e}")

finally:
    if wb:
        try: wb.Close(SaveChanges=False)
        except: pass
    if excel:
        excel.Quit()
