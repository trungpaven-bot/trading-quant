from vnstock import Vnstock
import pandas as pd
from datetime import datetime, timedelta
import time

def analyze_vn30_final():
    vn30_list = [
        "ACB", "BCM", "BID", "BVH", "CTG", "FPT", "GAS", "GVR", "HDB", "HPG",
        "MBB", "MSN", "MWG", "PLX", "POW", "SAB", "SHB", "SSB", "SSI", "STB",
        "TCB", "TPB", "VCB", "VHM", "VIB", "VIC", "VJC", "VNM", "VPB", "VRE"
    ]
    
    final_data = []
    
    end_date = datetime.now().strftime('%Y-%m-%d')
    # Tăng range lên 10 ngày để tránh cuối tuần/lễ
    start_date = (datetime.now() - timedelta(days=10)).strftime('%Y-%m-%d')

    print(f"Bắt đầu lấy dữ liệu (Vnstock v3.4.2 OOP) từ {start_date} đến {end_date}...")
    
    for ticker in vn30_list:
        try:
            print(f"Processing {ticker}...", end=" ")
            
            stock = Vnstock().stock(symbol=ticker, source='VCI')
            
            # 1. Lấy giá
            # Chạy retry nhẹ nếu lỗi mạng
            df_hist = None
            try:
                df_hist = stock.quote.history(start=start_date, end=end_date, interval='1D')
            except:
                pass
                
            if df_hist is None or df_hist.empty:
                print("No price data.")
                continue
                
            last_price = float(df_hist.iloc[-1]['close'])
            
            # 2. Lấy chỉ số tài chính (Ratio)
            df_ratio = None
            try:
                df_ratio = stock.finance.ratio(period='quarter', lang='vi')
            except:
                pass
            
            pe = 0.0
            roe = 0.0
            pb = 0.0
            
            if df_ratio is not None and not df_ratio.empty:
                # Flat multi-index columns để dễ truy cập
                df_ratio.columns = ['_'.join(col).strip() if isinstance(col, tuple) else col for col in df_ratio.columns.values]
                
                # Tìm cột chứa P/E, ROE (case insensitive)
                cols = df_ratio.columns.tolist()
                
                latest = df_ratio.iloc[0] # Dòng mới nhất
                
                # Helper tìm giá trị
                def get_val(keywords):
                    for col in cols:
                        if all(k.lower() in col.lower() for k in keywords):
                            return float(latest.get(col, 0))
                    return 0.0

                pe = get_val(['chỉ tiêu định giá', 'p/e'])
                if pe == 0: pe = get_val(['p/e']) # Fallback
                
                pb = get_val(['chỉ tiêu định giá', 'p/b'])
                if pb == 0: pb = get_val(['p/b'])
                
                roe = get_val(['chỉ tiêu sinh lời', 'roe'])
                if roe == 0: roe = get_val(['roe'])

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
            
    if final_data:
        df = pd.DataFrame(final_data)
        
        # Logic lọc
        df['DanhGia'] = df.apply(lambda x: 'Re' if (0 < x['P/E'] < 15 and x['ROE'] > 12) else '-', axis=1)
        
        output = "VN30_Analysis_Final.xlsx"
        df.to_excel(output, index=False)
        print(f"\n====================================")
        print(f"Xong! File saved: {output}")
        print(f"====================================")
        print(df.head())
    else:
        print("Không có dữ liệu.")

if __name__ == "__main__":
    analyze_vn30_final()
