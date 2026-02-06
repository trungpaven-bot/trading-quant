import pandas as pd
import numpy as np

class AiOracle:
    def __init__(self, loader, network_engine):
        self.loader = loader
        self.network = network_engine

    def ask(self, ticker):
        """
        Phân tích định lượng tổng hợp cho 1 mã.
        """
        verdict = "NEUTRAL"
        score = 50 # 0-100 (0=Strong Sell, 100=Strong Buy)
        reasons = []
        
        try:
            # 1. Lấy dữ liệu cơ bản (Basic Info)
            info = self.loader.get_ticker_info(ticker)
            pe = info.get('trailingPE', 0)
            pb = info.get('priceToBook', 0)
            sector = info.get('sector', 'Unknown')
            
            # 2. Phân tích kĩ thuật & Mạng lưới (Technical & Network)
            # Lấy data VN30 để xây context (nếu ticker nằm trong đó)
            # API hiện tại load toàn bộ VN30, ta giả sử ticker là 1 phần của VN30 hoặc ta load riêng
            df_price = self.network.prices
            
            trend_signal = 0
            neighbor_impact = "Trung lập"
            
            if ticker in df_price.columns:
                # Chạy lại tính toán spillover nếu chưa có (nhẹ)
                spillover_df = self.network.compute_spillover_momentum()
                row = spillover_df.loc[ticker]
                
                own_mom = row['Own_Momentum']
                net_mom = row['Neighbor_Momentum']
                final_sig = row['Network_Signal']
                
                # Logic chấm điểm Trend
                if final_sig > 0.05: # Tăng > 5%
                    score += 20
                    reasons.append(f"Xu hướng Tăng mạnh (Signal: {final_sig:.1%}).")
                elif final_sig < -0.05:
                    score -= 20
                    reasons.append(f"Xu hướng Giảm (Signal: {final_sig:.1%}).")
                
                # Logic Spilover
                if net_mom > own_mom:
                    neighbor_impact = "Tích cực"
                    reasons.append("Được hỗ trợ mạnh bởi dòng tiền ngành (Neighbors > Own).")
                elif net_mom < own_mom and net_mom < 0:
                    neighbor_impact = "Tiêu cực"
                    reasons.append("Bị kìm hãm bởi thị trường chung xấu.")

            # 3. Logic chấm điểm Value (P/E)
            # Giả định sơ bộ: PE < 12 là rẻ, > 20 là đắt (cần chỉnh theo ngành thực tế)
            if 0 < pe < 12:
                score += 15
                reasons.append(f"Định giá hấp dẫn (P/E={pe:.1f}).")
            elif pe > 25:
                score -= 10
                reasons.append(f"Định giá khá cao (P/E={pe:.1f}).")
                
            # 4. Tổng hợp Verdict
            if score >= 75: verdict = "STRONG BUY"
            elif score >= 60: verdict = "BUY"
            elif score <= 25: verdict = "STRONG SELL"
            elif score <= 40: verdict = "SELL"
            
            prompt_summary = (
                f"Phân tích mã {ticker} ({sector}):\n"
                f"- Hệ thống chấm điểm: {score}/100 ({verdict})\n"
                f"- Định giá: P/E={pe:.1f}, P/B={pb:.1f}\n"
                f"- Đà tăng trưởng mạng lưới: {neighbor_impact}\n"
                f"- Lý do chính: {'; '.join(reasons)}"
            )
            
            return {
                "ticker": ticker,
                "verdict": verdict,
                "score": score,
                "analysis": prompt_summary,
                "raw_data": {
                    "pe": pe,
                    "network_signal": float(final_sig) if 'final_sig' in locals() else 0
                }
            }
            
        except Exception as e:
            return {
                "ticker": ticker,
                "verdict": "UNKNOWN",
                "score": 0,
                "analysis": f"Không đủ dữ liệu phân tích. Lỗi: {str(e)}"
            }
