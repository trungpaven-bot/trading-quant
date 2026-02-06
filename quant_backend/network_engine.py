import pandas as pd
import numpy as np
import networkx as nx

class FinancialNetwork:
    def __init__(self, price_df: pd.DataFrame, window_size=20):
        """
        price_df: DataFrame index=Date, columns=Ticker, values=ClosePrice
        window_size: Rolling window size (days)
        """
        self.prices = price_df
        self.returns = self.prices.pct_change().dropna()
        self.correlation_matrix = None
        self.graph = None
        self.adj_matrix = None # Ma trận kề (trọng số)

    def build_network(self, threshold=0.5):
        """
        Xây dựng mạng lưới dựa trên tương quan Pearson.
        """
        # 1. Tính Correlation Matrix
        self.correlation_matrix = self.returns.corr(method='pearson')
        
        # 2. Khởi tạo đồ thị vô hướng
        G = nx.Graph()
        stocks = self.correlation_matrix.columns
        G.add_nodes_from(stocks)
        
        # 3. Thêm cạnh và xây dựng Adjacency Matrix
        # Chúng ta dùng ma trận numpy để tính toán spillover cho nhanh
        n = len(stocks)
        adj = np.zeros((n, n))
        
        for i in range(len(stocks)):
            for j in range(i + 1, len(stocks)):
                stock_a = stocks[i]
                stock_b = stocks[j]
                corr_val = self.correlation_matrix.iloc[i, j]
                
                if abs(corr_val) > threshold:
                    weight = abs(corr_val)
                    G.add_edge(stock_a, stock_b, weight=weight, correlation=corr_val)
                    adj[i, j] = weight
                    adj[j, i] = weight # Vô hướng
        
        self.graph = G
        self.adj_matrix = adj
        return G

    def analyze_centrality(self):
        """Tính toán các chỉ số mạng lưới."""
        if self.graph is None: raise ValueError("Build network first")
        degree = nx.degree_centrality(self.graph)
        pagerank = nx.pagerank(self.graph, weight='weight')
        clustering = nx.clustering(self.graph, weight='weight')
        stats = pd.DataFrame({'Degree': degree, 'PageRank': pagerank, 'Clustering': clustering})
        return stats.sort_values(by='PageRank', ascending=False)

    def compute_spillover_momentum(self, lookback=20, alpha=0.3):
        """
        Tính toán Network Trend-Following (Spillover).
        lookback: Số ngày tính momentum (ví dụ 20 ngày = 1 tháng).
        alpha: Trọng số của 'hàng xóm' (0.3 nghĩa là 30% tin vào hàng xóm).
        """
        if self.adj_matrix is None: raise ValueError("Build network first")
        
        # 1. Tính Momentum riêng (Own Momentum)
        # Return lũy kế trong N ngày gần nhất: (P_t / P_{t-n}) - 1
        momentum = (self.prices.iloc[-1] / self.prices.iloc[-lookback] - 1).fillna(0)
        mom_vec = momentum.values # Vector R
        
        # 2. Tính Momentum hàng xóm (Neighbor Momentum)
        # Normalize adjacency matrix để tổng trọng số hàng xóm = 1 (tránh scale bậy)
        # Row sum
        row_sums = self.adj_matrix.sum(axis=1)
        # Tránh chia cho 0
        row_sums[row_sums == 0] = 1.0 
        norm_adj = self.adj_matrix / row_sums[:, np.newaxis]
        
        neighbor_mom = norm_adj @ mom_vec
        
        # 3. Kết hợp (Spillover Signal)
        # S = (1 - alpha) * R_own + alpha * R_neighbor
        signal = (1 - alpha) * mom_vec + alpha * neighbor_mom
        
        # Đóng gói kết quả
        result = pd.DataFrame({
            'Own_Momentum': mom_vec,
            'Neighbor_Momentum': neighbor_mom,
            'Network_Signal': signal
        }, index=self.prices.columns)
        
        return result.sort_values(by='Network_Signal', ascending=False)

    def export_json_for_d3(self):
        if self.graph is None: return {}
        return nx.node_link_data(self.graph)
