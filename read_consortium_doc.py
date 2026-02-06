
from docx import Document
import os

file_path = r'g:\My Drive\30_RESOURCES (Kho tai nguyen)\30.01_Tinh hieu qua du an\02_Tra dinh muc theo TT 09 va 12\2026-01-30_ThoaThuanLienDanh_KhoLNG_NamTanTap.docx'

try:
    doc = Document(file_path)
    full_text = []
    for para in doc.paragraphs:
        if para.text.strip():
            full_text.append(para.text)
            
    # Print the first 50 lines to get the Parties details
    print("\n".join(full_text[:50]))
    
except Exception as e:
    print(f"Error reading docx: {e}")
