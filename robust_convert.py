
import win32com.client
import pandas as pd
import os
import re

word_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\QD_409.2025.QD-BXD_BXD_cong bo suat von dau tu 2024.doc"
csv_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Du_lieu_Suat_von_dau_tu_2024.csv"
excel_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\Du_lieu_Suat_von_dau_tu_2024.xlsx"

def clean_text(text):
    if not text: return ""
    # Remove control characters
    text = re.sub(r'[\x00-\x1f\x7f-\x9f]', ' ', text)
    return text.strip()

print(f"Reading Word Doc: {word_path}")

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    doc = word.Documents.Open(word_path)
    
    all_data = []
    
    for i, tbl in enumerate(doc.Tables):
        table_idx = i + 1
        print(f"Processing Table {table_idx}...", end='\r')
        all_data.append([f"--- BẢNG {table_idx} ---", "", "", ""])
        
        try:
            row_count = tbl.Rows.Count
            col_count = tbl.Columns.Count
            
            for r in range(1, row_count + 1):
                row_vals = []
                for c in range(1, col_count + 1):
                    try:
                        valid_c = min(c, tbl.Columns.Count) # Safety
                        txt = tbl.Cell(r, valid_c).Range.Text
                        row_vals.append(clean_text(txt))
                    except:
                        row_vals.append("")
                all_data.append(row_vals)
            all_data.append([""])
            
        except Exception as e:
            # Table structure error (merged cells often cause row access fail)
            all_data.append([f"Error reading Table {table_idx}: Structure complex/merged"])

    doc.Close(SaveChanges=False)
    word.Quit()
    
    print("\nSaving to Excel...")
    # Normalize lengths
    max_cols = max([len(row) for row in all_data]) if all_data else 0
    padded_data = [row + [""]*(max_cols-len(row)) for row in all_data]
    
    df = pd.DataFrame(padded_data)
    # Save as Excel
    df.to_excel(excel_path, index=False, header=False)
    print(f"Success! Saved to: {excel_path}")

except Exception as e:
    print(f"\nFatal Error: {e}")
    try: doc.Close(SaveChanges=False)
    except: pass
    try: word.Quit()
    except: pass
