import pandas as pd
import numpy as np

class AiOracle:
    def __init__(self, loader, network_engine):
        self.loader = loader
        self.network = network_engine

    def ask(self, ticker):
        """
        AI Oracle Logic (Rule-based):
        1. Fundamental Snapshot: Lấy P/E, ROE từ yfinance.
        2. Technical Signal: So sánh Giá hiện tại vs MA20.
        """
        try:
            # --- 1. FUNDAMENTAL SNAPSHOT (Soi Cơ bản) ---
            # Lấy thông tin từ yfinance (đã cache trong loader hoặc gọi trực tiếp)
            info = self.loader.get_ticker_info(ticker)
            
            # Xử lý dữ liệu thô
            pe = info.get('trailingPE', 0)
            roe = info.get('returnOnEquity', 0)
            price = info.get('currentPrice', 0)
            
            # Nếu không có giá từ Info (lỗi API), thử lấy từ lịch sử
            if price == 0 and not self.network.prices.empty and ticker in self.network.prices:
                price = self.network.prices[ticker].iloc[-1]

            # Logic Fundamental Đơn giản
            fund_signal = "TRUNG LẬP"
            if 0 < pe < 12 and roe > 0.15: fund_signal = "RẺ (HẤP DẪN)"
            elif pe > 25: fund_signal = "ĐẮT (CẨN TRỌNG)"
            
            # --- 2. AI ORACLE (Technical Rule-based) ---
            # Lấy lịch sử giá để tính MA20
            # Ta có thể dùng dữ liệu từ Network Engine (đã load sẵn VN30 6 tháng)
            prices_series = None
            if ticker in self.network.prices.columns:
                prices_series = self.network.prices[ticker]
            else:
                # Nếu mã không nằm trong VN30 đã load, cần fetch riêng (nhưng để nhanh ta tạm skip hoặc fetch nóng)
                # Ở đây giả định user hỏi mã trong VN30 trước
                pass
            
            tech_verdict = "KHÔNG ĐỦ DỮ LIỆU"
            ma20 = 0
            
            if prices_series is not None and len(prices_series) >= 20:
                ma20 = prices_series.rolling(window=20).mean().iloc[-1]
                current_price = prices_series.iloc[-1]
                
                # Rule-based Logic (Chính xác tuyệt đối)
                if current_price > ma20:
                    tech_verdict = "XU HƯỚNG TĂNG (NẮM GIỮ)"
                else:
                    tech_verdict = "XU HƯỚNG GIẢM (QUAN SÁT)"
                    
            # --- 3. Đóng gói kết quả ---
            analysis_text = (
                f"🤖 AI ORACLE ALERTS:\n"
                f"- Tín hiệu Kỹ thuật: {tech_verdict}\n"
                f"  (Giá {current_price:,.0f} vs MA20 {ma20:,.0f})\n"
                f"- Tín hiệu Cơ bản: {fund_signal}\n"
                f"  (P/E={pe:.1f}, ROE={roe*100:.1f}%)"
            )
            
            return {
                "ticker": ticker,
                "fundamental": {
                    "pe": round(pe, 2) if pe else 0,
                    "roe": f"{roe*100:.1f}%" if roe else "N/A",
                    "signal": fund_signal
                },
                "technical": {
                    "price": current_price,
                    "ma20": round(ma20, 2),
                    "signal": tech_verdict
                },
                "full_analysis": analysis_text
            }
            
        except Exception as e:
            return {
                "ticker": ticker,
                "error": str(e),
                "full_analysis": "Hệ thống đang bận hoặc không tìm thấy mã."
            }
