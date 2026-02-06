
import win32com.client
import os
import pandas as pd

file_path = r"g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\QD_409.2025.QD-BXD_BXD_cong bo suat von dau tu 2024.doc"

print(f"Processing: {file_path}")

try:
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    
    doc = word.Documents.Open(file_path)
    print(f"Document opened. Tables count: {doc.Tables.Count}")
    
    all_tables_data = []

    # Iterate over tables
    # Limit to first 20 tables to look for Suat Von Dau Tu data (usually it has many tables)
    count = 0
    for tbl in doc.Tables:
        count += 1
        table_data = []
        try:
            # We assume regular tables. Complex merged tables might be tricky.
            # Using Range.Text to get content.
            # Convert simple table to list of lists
            
            # Optimization: retrieving cell by cell via COM is SLOW.
            # A faster way is copying to clipboard or using `ConvertToText` but that modifies doc.
            # Let's try reading cell by cell but limiting rows for inspection first.
            
            row_count = tbl.Rows.Count
            col_count = tbl.Columns.Count
            print(f"Table {count}: {row_count} rows, {col_count} cols")
            
            if row_count > 100:
                print("  (Large table, reading first 10 rows only for preview)")
                scan_rows = 10
            else:
                scan_rows = row_count
                
            for r in range(1, scan_rows + 1):
                row_vals = []
                try:
                    # Rows(r) might fail if there are vertically merged cells. 
                    # Accessing cells via Range is safer but harder to coordinate.
                    # Let's try regular Rows(r).Cells
                    for c in range(1, col_count + 1):
                        try:
                            # Note: Word cells end with \r\x07. Strip it.
                            txt = tbl.Cell(r, c).Range.Text
                            txt = txt.replace('\r\x07', '').strip()
                            row_vals.append(txt)
                        except:
                            row_vals.append("<Merge/Error>")
                    table_data.append(row_vals)
                except:
                    table_data.append(["<Row Error>"])
            
            all_tables_data.append(table_data)
        except Exception as e:
            print(f"Error reading Table {count}: {e}")
        
        if count >= 10: 
            print("Stopping after 10 tables for initial inspection.")
            break

    doc.Close(SaveChanges=False)
    word.Quit()
    
    # Dump Preview
    print("\n--- Tables Content Preview ---")
    for i, tbl in enumerate(all_tables_data):
        print(f"\n--- Table {i+1} ---")
        for row in tbl:
            print(row)

except Exception as e:
    print(f"Fatal Error: {e}")
    try:
        doc.Close(SaveChanges=False)
    except:
        pass
    try:
        word.Quit()
    except:
        pass
