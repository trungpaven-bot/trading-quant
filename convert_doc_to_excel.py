
import win32com.client
import pandas as pd
import os

word_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\QD_409.2025.QD-BXD_BXD_cong bo suat von dau tu 2024.doc"
excel_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Du_lieu_Suat_von_dau_tu_2024.xlsx"

print(f"Reading: {word_path}")

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(word_path)
    
    print(f"Total Tables: {doc.Tables.Count}")
    
    all_data = []
    
    # Iterate all tables
    for i, tbl in enumerate(doc.Tables):
        table_idx = i + 1
        print(f"Processing Table {table_idx}...", end='\r')
        
        try:
            # Helper to get text faster? No, just iterate for robustness.
            # We add a Header Row identifying the table first
            all_data.append([f"--- BẢNG {table_idx} ---", "", "", "", ""]) 
            
            row_count = tbl.Rows.Count
            col_count = tbl.Columns.Count
            
            # Heuristic: skip tiny tables (layout tables) if needed, but safe to keep all
            
            for r in range(1, row_count + 1):
                row_vals = []
                for c in range(1, col_count + 1):
                    try:
                        txt = tbl.Cell(r, c).Range.Text.replace('\r\x07', '').strip()
                        row_vals.append(txt)
                    except:
                        row_vals.append("")
                all_data.append(row_vals)
            
            all_data.append([""]) # Empty row between tables
            
        except Exception as e:
            print(f"\nError in Table {table_idx}: {e}")
            all_data.append([f"Error reading Table {table_idx}"])

    doc.Close(SaveChanges=False)
    word.Quit()
    
    print("\nWriting to Excel...")
    df = pd.DataFrame(all_data)
    df.to_excel(excel_path, index=False, header=False)
    print(f"Saved to: {excel_path}")

except Exception as e:
    print(f"Fatal Error: {e}")
    try: text = doc.Close(SaveChanges=False) 
    except: pass
    try: word.Quit() 
    except: pass
