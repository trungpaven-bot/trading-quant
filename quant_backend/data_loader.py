import yfinance as yf
import pandas as pd
from datetime import datetime, timedelta

class DataLoader:
    def __init__(self):
        # VN30 list chuẩn hóa cho Yahoo Finance (.VN cho sàn HOSE)
        self.vn30_tickers = [
            "ACB.VN", "BCM.VN", "BID.VN", "BVH.VN", "CTG.VN", "FPT.VN", "GAS.VN", 
            "GVR.VN", "HDB.VN", "HPG.VN", "MBB.VN", "MSN.VN", "MWG.VN", "PLX.VN", 
            "POW.VN", "SAB.VN", "SHB.VN", "SSB.VN", "SSI.VN", "STB.VN", "TCB.VN", 
            "TPB.VN", "VCB.VN", "VHM.VN", "VIB.VN", "VIC.VN", "VJC.VN", "VNM.VN", 
            "VPB.VN", "VRE.VN"
        ]

    def fetch_price_history(self, tickers=None, period="1y", interval="1d"):
        """
        Tải dữ liệu giá (Open, High, Low, Close, Volume)
        tickers: List mã (ví dụ ['AAPL', 'MSFT', 'HPG.VN']). Nếu None lấy VN30.
        period: 1d, 5d, 1mo, 3mo, 6mo, 1y, 2y, 5y, 10y, ytd, max
        interval: 1m, 2m... 1h, 1d, 5d, 1wk, 1mo
        """
        target_tickers = tickers if tickers else self.vn30_tickers
        
        print(f"Downloading data for {len(target_tickers)} tickers...")
        
        # Batch download cực nhanh
        data = yf.download(
            tickers=target_tickers, 
            period=period, 
            interval=interval, 
            group_by='ticker', 
            auto_adjust=True, 
            prepost=True, 
            threads=True
        )
        
        # Yahoo trả về MultiIndex columns nếu tải nhiều mã
        # Cần chuẩn hóa lại thành dạng dễ dùng: Dictionary of DataFrames hoặc Panel
        # Ở đây tôi sẽ trả về DataFrame chỉ chứa Close price cho việc tính toán Matrix (giống yêu cầu cũ)
        # Nếu muốn full OHLCV, xử lý riêng.
        
        return data

    def get_close_price_matrix(self, tickers=None, period="6mo"):
        """
        Helper function để lấy ma trận giá đóng cửa (dùng cho Network/Portfolio optimization)
        """
        target_tickers = tickers if tickers else self.vn30_tickers
        data = yf.download(target_tickers, period=period, auto_adjust=True, progress=False)
        
        # Lấy cột Close
        if 'Close' in data.columns:
            close_data = data['Close']
            # Drop các cột toàn NaN (nếu mã sai)
            close_data.dropna(axis=1, how='all', inplace=True)
            # Fill forward hole
            close_data.ffill(inplace=True)
            return close_data
        else:
            return pd.DataFrame()

    def get_ticker_info(self, ticker):
        """
        Lấy thông tin cơ bản: Sector, PE, Market Cap, Business Summary
        """
        try:
            t = yf.Ticker(ticker)
            return t.info
        except:
            return {}

    def get_financials(self, ticker):
        """
        Lấy báo cáo tài chính (Income Statement, Balance Sheet, Cash Flow)
        """
        t = yf.Ticker(ticker)
        return {
            "income_stmt": t.income_stmt.to_dict(),
            "balance_sheet": t.balance_sheet.to_dict(),
            "cash_flow": t.cashflow.to_dict()
        }
