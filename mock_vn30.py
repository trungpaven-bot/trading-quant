import pandas as pd
import random

def generate_mock_data():
    vn30_list = [
        "ACB", "BCM", "BID", "BVH", "CTG", "FPT", "GAS", "GVR", "HDB", "HPG",
        "MBB", "MSN", "MWG", "PLX", "POW", "SAB", "SHB", "SSB", "SSI", "STB",
        "TCB", "TPB", "VCB", "VHM", "VIB", "VIC", "VJC", "VNM", "VPB", "VRE"
    ]
    
    data = []
    print("Đang sinh dữ liệu mô phỏng cho VN30...")
    
    for ticker in vn30_list:
        # Giả lập giá và chỉ số
        price = random.randint(15000, 120000)
        pe = round(random.uniform(8.0, 25.0), 2)
        pb = round(random.uniform(1.0, 4.0), 2)
        roe = round(random.uniform(10.0, 25.0), 2) # %
        
        # Logic giả lập: Bank thường P/E thấp, Retail P/E cao
        if ticker in ["VCB", "ACB", "TCB", "MBB"]:
            pe = round(random.uniform(9.0, 14.0), 2)
        if ticker in ["MWG", "FPT", "VNM"]:
            pe = round(random.uniform(15.0, 22.0), 2)
            
        data.append({
            "Ticker": ticker,
            "Price": price,
            "P/E": pe,
            "P/B": pb,
            "ROE": roe,
            "Date": "2026-02-05" # Mock date
        })
        
    df = pd.DataFrame(data)
    
    # Logic tính toán (giống script ban đầu)
    df['DanhGia'] = df.apply(
        lambda x: 'Re' if (x['P/E'] < 15 and x['ROE'] > 15) else 'BinhThuong', 
        axis=1
    )
    
    output_file = "VN30_Analysis.xlsx"
    df.to_excel(output_file, index=False)
    print(f"Đã tạo file Excel giả lập tại: {output_file}")
    print(df.head())

if __name__ == "__main__":
    generate_mock_data()
