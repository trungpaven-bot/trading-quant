
import win32com.client
import pandas as pd
import re

word_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\QD_409.2025.QD-BXD_BXD_cong bo suat von dau tu 2024.doc"
output_excel = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Du_lieu_Co_Tieu_De_Bang.xlsx"

def clean_text(text):
    if not text: return ""
    return re.sub(r'[\x00-\x1f\x7f-\x9f]', ' ', text).strip()

print(f"Opening Word: {word_path}")

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(word_path)
    
    all_data = []
    
    # Iterate Tables
    count = doc.Tables.Count
    print(f"Found {count} tables.")
    
    for i in range(1, count + 1):
        print(f"Processing Table {i}/{count}...", end='\r')
        tbl = doc.Tables(i)
        
        # 1. Capture Title (Look backwards from Table)
        # Strategy: Get Range just before table
        # We try to get the 3 previous paragraphs and search for "Bảng"
        title = f"Bảng {i} (Không tên)"
        try:
            # Move cursor to start of table
            rng = tbl.Range
            rng.Collapse(1) # wdCollapseStart
            
            # Look back up to 3 paragraphs
            found_title = False
            for _ in range(3):
                rng.MoveStart(4, -1) # wdParagraph, -1
                txt = clean_text(rng.Text)
                if len(txt) > 5 and ("Bảng" in txt or "STT" in txt or "Suất vốn" in txt):
                    title = txt
                    found_title = True
                    break
            
            if not found_title:
                title = f"Bảng {i} - {txt[:50]}..."
        except:
            pass
            
        # 2. Extract Data
        try:
            row_count = tbl.Rows.Count
            col_count = tbl.Columns.Count
            
            # Skip tiny tables
            if row_count < 2 or col_count < 2: continue

            for r in range(1, row_count + 1):
                row_vals = [title] # Col 0 is Table Name
                for c in range(1, col_count + 1):
                    try:
                        valid_c = min(c, 10) # Limit cols to avoid crazy merged cells
                        txt = tbl.Cell(r, valid_c).Range.Text
                        row_vals.append(clean_text(txt))
                    except:
                        row_vals.append("")
                all_data.append(row_vals)
                
        except Exception as e:
            # print(f"Error reading content table {i}: {e}")
            pass

    doc.Close(SaveChanges=False)
    word.Quit()
    
    print("\nSaving to Excel...")
    # Normalize
    max_cols = max([len(row) for row in all_data]) if all_data else 0
    padded_data = [row + [""]*(max_cols-len(row)) for row in all_data]
    
    df = pd.DataFrame(padded_data)
    df.to_excel(output_excel, index=False, header=False)
    print(f"Done. Saved to {output_excel}")

except Exception as e:
    print(f"\nFatal Error: {e}")
    try: doc.Close(SaveChanges=False) 
    except: pass
    try: word.Quit() 
    except: pass
