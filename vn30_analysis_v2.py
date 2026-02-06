from vnstock import Vnstock
import pandas as pd
from datetime import datetime, timedelta
import time

def analyze_vn30_v2():
    vn30_list = [
        "ACB", "BCM", "BID", "BVH", "CTG", "FPT", "GAS", "GVR", "HDB", "HPG",
        "MBB", "MSN", "MWG", "PLX", "POW", "SAB", "SHB", "SSB", "SSI", "STB",
        "TCB", "TPB", "VCB", "VHM", "VIB", "VIC", "VJC", "VNM", "VPB", "VRE"
    ]
    
    print("Khởi tạo Vnstock client...")
    # Khởi tạo client 1 lần nếu được, hoặc trong loop
    
    final_data = []
    
    end_date = datetime.now().strftime('%Y-%m-%d')
    start_date = (datetime.now() - timedelta(days=5)).strftime('%Y-%m-%d')

    print(f"Bắt đầu lấy dữ liệu (Vnstock v3.4.2 OOP) từ {start_date} đến {end_date}...")
    
    for ticker in vn30_list:
        try:
            print(f"Processing {ticker}...", end=" ")
            
            # Khởi tạo đối tượng stock
            # Source 'VCI' thường ổn định cho dữ liệu lịch sử
            stock = Vnstock().stock(symbol=ticker, source='VCI')
            
            # 1. Lấy giá lịch sử (OHLC)
            # Cú pháp mới: stock.quote.history(...)
            df_hist = stock.quote.history(start=start_date, end=end_date, interval='1D')
            
            if df_hist is None or df_hist.empty:
                print("No price data.")
                continue
                
            last_price = float(df_hist.iloc[-1]['close'])
            
            # 2. Lấy chỉ số tài chính (Ratio)
            # Cú pháp mới: stock.finance.ratio(...)
            df_ratio = stock.finance.ratio(period='quarter', lang='vi')
            
            pe = 0.0
            roe = 0.0
            pb = 0.0
            
            if df_ratio is not None and not df_ratio.empty:
                # Dữ liệu ratio thường có cột: ticker, period, year, quarter, priceToEarning, priceToBook, roe...
                # Cần check tên cột chính xác. Thông thường là camelCase.
                latest = df_ratio.iloc[0] # Vì thường sort mới nhất lên đầu hoặc cuối, cần check data thực tế.
                # Giả sử lấy dòng đầu tiên (quý gần nhất)
                pe = float(latest.get('priceToEarning', 0))
                pb = float(latest.get('priceToBook', 0))
                roe = float(latest.get('roe', 0))
                
            print(f"OK (P={last_price:,.0f}, PE={pe})")
            
            final_data.append({
                "Ticker": ticker,
                "Price": last_price,
                "P/E": pe,
                "P/B": pb,
                "ROE": roe * 100 if roe < 1 else roe
            })
            
        except Exception as e:
            print(f"Err: {e}")
            
    # Xuất Excel
    if final_data:
        df = pd.DataFrame(final_data)
        
        df['DanhGia'] = df.apply(lambda x: 'Re' if (0 < x['P/E'] < 15 and x['ROE'] > 12) else '-', axis=1)
        
        output = "VN30_Analysis_Final.xlsx"
        df.to_excel(output, index=False)
        print(f"Xong! File saved: {output}")
        print(df.head())
    else:
        print("Không có dữ liệu.")

if __name__ == "__main__":
    analyze_vn30_v2()
