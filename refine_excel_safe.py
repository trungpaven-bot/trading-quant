
import win32com.client as win32
import os

# Define paths
file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1.xlsx'
output_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\FS CCN Phan Dinh Phung 1_Refined_Safe.xlsx'

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
        try:
            wb.Close(SaveChanges=False)
        except: pass
        excel.Quit()
        exit()

    # Create a copy
    print("Copying worksheet...")
    ws_source.Copy(Before=ws_source)
    ws_new = wb.ActiveSheet
    ws_new.Name = "PA 3 (Refined)"
    
    # 1. Move Operating Costs Block
    # Look for "Nhóm thông tin về chi phí hoạt động"
    found = ws_new.Cells.Find(What="Nhóm thông tin về chi phí hoạt động", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
    
    if found:
        print(f"Found Operating Costs at {found.Address}")
        # Assuming the block starts at 'found' and extends say 15 rows and 6 columns (F to K is 6 cols)
        start_row = found.Row
        start_col_idx = found.Column
        
        # We want to insert space below the "Capital Structure" (Cơ cấu nguồn vốn) section.
        # Find "Cơ cấu nguồn vốn"
        found_cap = ws_new.Cells.Find(What="Cơ cấu nguồn vốn đầu tư", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
        
        insert_row = 25 # Default fallback
        if found_cap:
            # Capital structure usually has 3-5 rows below it.
            # Let's say we check down 6 rows
            insert_row = found_cap.Row + 6
        
        print(f"Inserting rows at Row {insert_row}")
        # Insert enough rows (approx 20)
        ws_new.Rows(f"{insert_row}:{insert_row+20}").Insert()
        
        # Cut and Paste
        # We need to determine the exact range to cut. 
        # From previous analysis, it was around F4/F5 to K20/L20.
        # Let's grab a generous block from found.Row down 18 rows, and from found.Column to found.Column + 8 (safe bet)
        # Or better, just Cut the specific range if we are confident.
        # Previous script said: F5:K20.
        
        # Let's try to find the end of the block?
        # It's risky to guess. I'll stick to a reasonably sized block F4:L22 (covering potential merges)
        # Converting column index to letter is annoying in pure win32 without helpers, but Ranges work with Cells.
        
        cut_range = ws_new.Range(ws_new.Cells(start_row, start_col_idx), ws_new.Cells(start_row+17, start_col_idx+6)) # 18 rows, 7 cols
        dest_cell = ws_new.Cells(insert_row + 1, 2) # Column B
        
        print(f"Cutting from {cut_range.Address} to {dest_cell.Address}")
        cut_range.Cut(Destination=dest_cell)
        
        # Add a Label above it
        ws_new.Cells(insert_row, 2).Value = "IV. CHI PHÍ HOẠT ĐỘNG KHÁC"
        ws_new.Cells(insert_row, 2).Font.Bold = True
        ws_new.Cells(insert_row, 2).Font.Color = 255 # Red
    
    # 2. Fix Table Numbering
    # Find "Bảng 1.2" that appears AFTER the first one?
    # Or just Find "Bảng  1.2 : Bảng tính doanh thu"
    found_table = ws_new.Cells.Find(What="Bảng tính doanh thu", LookIn=win32.constants.xlValues, LookAt=win32.constants.xlPart)
    if found_table:
        val = found_table.Value
        if "1.2" in val:
            print(f"Renaming table at {found_table.Address}")
            found_table.Value = val.replace("1.2", "1.3")

    # Save
    print(f"Saving to {output_path}...")
    wb.SaveAs(output_path)

except Exception as e:
    print(f"Error: {e}")

finally:
    if wb:
        try: wb.Close(SaveChanges=False)
        except: pass
    if excel:
        excel.Quit()
