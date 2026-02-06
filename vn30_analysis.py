import pandas as pd
from vnstock import stock_historical_data, stock_evaluation
from datetime import datetime, timedelta
import time

def analyze_vn30():
    # Danh sách VN30
    vn30_list = [
        "ACB", "BCM", "BID", "BVH", "CTG", "FPT", "GAS", "GVR", "HDB", "HPG",
        "MBB", "MSN", "MWG", "PLX", "POW", "SAB", "SHB", "SSB", "SSI", "STB",
        "TCB", "TPB", "VCB", "VHM", "VIB", "VIC", "VJC", "VNM", "VPB", "VRE"
    ]

    print(f"Đang lấy dữ liệu cho {len(vn30_list)} mã VN30...")
    
    today = datetime.now().strftime("%Y-%m-%d")
    start_date = (datetime.now() - timedelta(days=5)).strftime("%Y-%m-%d") # Lấy 5 ngày gần nhất để chắc chắn có dữ liệu
    
    final_data = []

    for ticker in vn30_list:
        try:
            # 1. Lấy giá đóng cửa gần nhất
            df_price = stock_historical_data(symbol=ticker, start_date=start_date, end_date=today, resolution='1D', type='stock')
            
            if df_price.empty:
                print(f"[Warn] Không có dữ liệu giá cho {ticker}")
                continue
                
            last_price = float(df_price.iloc[-1]['close'])
            trading_date = str(df_price.iloc[-1]['time'])

            # 2. Lấy chỉ số định giá (P/E, P/B, ROE)
            # Hàm stock_evaluation trả về dữ liệu định giá theo ngày
            df_eval = stock_evaluation(symbol=ticker, period=1, time_window='D')
            
            pe = 0.0
            pb = 0.0
            roe = 0.0 # ROE có thể cần công thức hoặc nguồn khác, nhưng thử lấy từ eval nếu có
            
            if not df_eval.empty:
                # Cột thường thấy: PE, PB, ROE
                # Chuyển tên cột về chữ thường hoặc kiểm tra kỹ
                last_eval = df_eval.iloc[-1]
                pe = float(last_eval.get('PE', 0))
                pb = float(last_eval.get('PB', 0))
                # Một số version vnstock trả về 'ROE' trong evaluation, nếu không sẽ là 0
                roe = float(last_eval.get('ROE', 0))

            print(f"  > {ticker}: Giá={last_price:,.0f}, P/E={pe}, ROE={roe}")

            final_data.append({
                "Mã CP": ticker,
                "Giá hiện tại": last_price,
                "P/E": pe,
                "P/B": pb,
                "ROE (%)": roe * 100 if roe < 1 else roe, # Chuẩn hóa nếu cần
                "Ngày GD": trading_date
            })
            
        except Exception as e:
            print(f"[Err] Lỗi khi xử lý {ticker}: {e}")
            time.sleep(1) # Nghỉ chút nếu lỗi để tránh spam request

    # Xuất Excel
    if final_data:
        df_result = pd.DataFrame(final_data)
        
        # Logic lọc cổ phiếu "Rẻ"
        # Điều kiện ví dụ: P/E < 15 VÀ ROE > 12%
        def check_cheap(row):
            if 0 < row['P/E'] < 15 and row['ROE (%)'] > 12:
                return 'Hấp dẫn'
            return '-'
            
        df_result['Nhận định'] = df_result.apply(check_cheap, axis=1)
        
        output_path = "VN30_Analysis_Realtime.xlsx"
        df_result.to_excel(output_path, index=False)
        print(f"\n==========================================")
        print(f"✅ Đã xuất dữ liệu thành công ra file: {output_path}")
        print(f"==========================================")
        print(df_result.head())
    else:
        print("Không thu thập được dữ liệu nào.")

if __name__ == "__main__":
    analyze_vn30()
