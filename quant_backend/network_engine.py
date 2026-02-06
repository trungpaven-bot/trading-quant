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

    def build_network(self, threshold=0.5):
        """
        Xây dựng mạng lưới dựa trên tương quan Pearson.
        threshold: Ngưỡng tương quan để vẽ cạnh (0.0 -> 1.0)
        """
        # 1. Tính Correlation Matrix
        self.correlation_matrix = self.returns.corr(method='pearson')
        
        # 2. Khởi tạo đồ thị vô hướng (Undirected Graph)
        G = nx.Graph()
        
        # Thêm nodes
        stocks = self.correlation_matrix.columns
        G.add_nodes_from(stocks)
        
        # 3. Thêm cạnh dựa trên threshold
        # Chỉ xét tam giác trên của ma trận để tránh lặp
        for i in range(len(stocks)):
            for j in range(i + 1, len(stocks)):
                stock_a = stocks[i]
                stock_b = stocks[j]
                corr_val = self.correlation_matrix.iloc[i, j]
                
                # Nếu tương quan cao > threshold -> Có liên kết
                if abs(corr_val) > threshold:
                    # Trọng số cạnh = độ mạnh tương quan
                    G.add_edge(stock_a, stock_b, weight=abs(corr_val), correlation=corr_val)
        
        self.graph = G
        return G

    def analyze_centrality(self):
        """
        Tính toán các chỉ số quan trọng của mạng lưới.
        """
        if self.graph is None:
            raise ValueError("Cần chạy build_network() trước.")
            
        # Degree Centrality: Ai có nhiều kết nối nhất? (Dễ bị ảnh hưởng bởi thị trường chung)
        degree = nx.degree_centrality(self.graph)
        
        # PageRank: Ai là người dẫn dắt quan trọng?
        pagerank = nx.pagerank(self.graph, weight='weight')
        
        # Clustering: Ai thuộc nhóm gắn kết chặt chẽ?
        clustering = nx.clustering(self.graph, weight='weight')
        
        # Tổng hợp kết quả
        stats = pd.DataFrame({
            'Degree': degree,
            'PageRank': pagerank,
            'Clustering': clustering
        })
        
        # Sắp xếp theo PageRank (Độ ảnh hưởng)
        return stats.sort_values(by='PageRank', ascending=False)

    def export_json_for_d3(self):
        """
        Xuất dữ liệu format JSON để frontend vẽ graph (d3.js / react-force-graph)
        """
        if self.graph is None: return {}
        
        data = nx.node_link_data(self.graph)
        # Bổ sung thêm thông tin cho Link để vẽ màu (Xanh = Tương quan dương, Đỏ = Âm)
        return data
