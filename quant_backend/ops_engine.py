import numpy as np
import pandas as pd

class OPSEngine:
    def __init__(self, prices: pd.DataFrame):
        """
        prices: DataFrame (Index=Date, Columns=Ticker, Value=Close)
        """
        self.prices = prices
        # Tính Relative Returns (x_t = p_t / p_{t-1})
        # pct_change() = (p_t - p_{t-1})/p_{t-1} -> x_t = 1 + pct_change
        self.rel_returns = 1 + self.prices.pct_change().dropna()
        self.tickers = self.prices.columns.tolist()
        self.n_assets = len(self.tickers)
        self.dates = self.rel_returns.index

    def run(self, strategy="EG", eta=0.05, tc=0.001):
        """
        Chạy thuật toán OPS.
        strategy: 'EG' (Exponentiated Gradient), 'UCRP' (Universal Constant Rebalanced - Benchmark)
        eta: Learning rate (chỉ dùng cho EG)
        tc: Transaction cost (0.1% mỗi lần trade)
        """
        T, N = self.rel_returns.shape
        if T == 0: return {}

        # 1. Khởi tạo
        # Weights (b): Tỷ trọng danh mục. Ban đầu chia đều.
        weights = np.ones(N) / N
        
        # History lưu lại để vẽ chart
        history_weights = [weights]
        wealth = [1.0] # Bắt đầu với 1 đơn vị tiền
        
        # Ma trận return dạng numpy cho nhanh
        X = self.rel_returns.values

        # 2. Vòng lặp giao dịch (Online Loop)
        for t in range(T):
            x_t = X[t] # Vector biến động giá ngày t (ví dụ: [1.01, 0.99, ...])
            
            # --- Tính lợi nhuận ngày t ---
            # Portfolio return = dot_product(weights, asset_returns)
            # R_t = \sum (b_{t,i} * x_{t,i})
            day_return = np.dot(weights, x_t)
            
            # Trừ phí giao dịch (Transaction Cost - Simplified approximation)
            # Trong thực tế cần so sánh w_{t-1} và w_t để tính turnover
            cost = 0
            if t > 0:
                # Turnover: Lượng thay đổi danh mục
                # w' = weights sau khi giá biến động nhưng chưa rebalance
                prev_w_post = (history_weights[-1] * X[t-1]) / np.dot(history_weights[-1], X[t-1])
                turnover = np.sum(np.abs(weights - prev_w_post))
                cost = turnover * tc
            
            day_return_net = day_return - cost
            
            # Cập nhật tổng tài sản
            wealth.append(wealth[-1] * day_return_net)

            # --- Cập nhật Weights cho ngày mai (Algo Logic) ---
            if strategy == "EG":
                # EG Update Rule:
                # w_{new} = w_{old} * exp(eta * gradient)
                # Gradient của log-return function xấp xỉ x_t / (w . x_t)
                grad = x_t / day_return
                
                # Exponentiated Gradient update
                new_w = weights * np.exp(eta * grad)
                new_w /= np.sum(new_w) # Normalize lại về 1
                weights = new_w
                
            elif strategy == "UCRP":
                # Buy and Hold (Fixed Rebalance)
                # Giữ nguyên target weight ban đầu (hoặc không làm gì)
                # Ở đây giả định UCRP là rebalance về đều nhau mỗi ngày (1/N)
                weights = np.ones(N) / N
            
            history_weights.append(weights)

        # 3. Đóng gói kết quả
        df_wealth = pd.Series(wealth, index=[self.prices.index[0]] + list(self.dates))
        df_weights = pd.DataFrame(history_weights[1:], index=self.dates, columns=self.tickers)
        
        return {
            "final_wealth": wealth[-1],
            "total_return": (wealth[-1] - 1) * 100,
            "equity_curve": df_wealth,
            "weights": df_weights
        }

    def get_latest_allocation(self, strategy="EG", eta=0.05):
        """
        Chỉ chạy để lấy tỷ trọng khuyến nghị cho ngày mai
        """
        res = self.run(strategy, eta)
        if not res: return {}
        
        last_weights = res["weights"].iloc[-1]
        # Chỉ lấy các mã có tỷ trọng > 1% để đỡ rác
        allocation = last_weights[last_weights > 0.01].sort_values(ascending=False)
        return allocation.to_dict()
