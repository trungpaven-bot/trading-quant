from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from quant_backend.data_loader import DataLoader
from quant_backend.network_engine import FinancialNetwork
from quant_backend.ops_engine import OPSEngine
import pandas as pd

app = FastAPI()

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Khởi tạo DataLoader (Singleton)
loader = DataLoader()

@app.get("/")
def read_root():
    return {"status": "ok", "service": "TitanLabs Quant Engine (yfinance Powered)"}

@app.get("/api/analyze-network")
def analyze_market_network(period: str = "6mo", threshold: float = 0.5):
    """
    Phân tích mạng lưới VN30 dùng yfinance.
    period: 1mo, 3mo, 6mo, 1y...
    threshold: Ngưỡng tương quan (0.5 - 0.9)
    """
    try:
        print(f"Fetch matrix period={period}...")
        prices = loader.get_close_price_matrix(period=period)
        
        if prices.empty:
            raise HTTPException(status_code=404, detail="No data fetched")
            
        print("Building graph...")
        net_engine = FinancialNetwork(prices, window_size=30)
        net_engine.build_network(threshold=threshold)
        
        stats = net_engine.analyze_centrality()
        graph_data = net_engine.export_json_for_d3()
        
        stats_dict = stats.reset_index().rename(columns={'index': 'Ticker'}).to_dict(orient='records')
        
        return {
            "graph": graph_data,
            "centrality": stats_dict,
            "meta": {
                "tickers": len(prices.columns),
                "data_points": len(prices),
                "period": period
            }
        }
    except Exception as e:
        print(f"Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/stock/{ticker}")
def get_stock_detail(ticker: str):
    """
    Lấy thông tin chi tiết mã (Info + Financials).
    Ví dụ: ticker='HPG.VN' hoặc 'AAPL'
    """
    info = loader.get_ticker_info(ticker)
    if not info:
        raise HTTPException(status_code=404, detail="Ticker not found")
    return info

@app.get("/api/stock/{ticker}/financials")
def get_stock_financials(ticker: str):
    """
    Lấy báo cáo tài chính
    """
    return loader.get_financials(ticker)

@app.get("/api/optimize-portfolio")
def optimize_portfolio(strategy: str = "EG", eta: float = 0.05, period: str = "6mo"):
    """
    Chạy thuật toán tối ưu danh mục (OPS).
    strategy: 'EG' (Exponentiated Gradient) hoặc 'UCRP' (Benchmark)
    eta: Learning rate (mặc định 0.05)
    """
    try:
        print(f"Loading data for OPS ({period})...")
        # 1. Tải giá
        prices = loader.get_close_price_matrix(period=period)
        if prices.empty:
            raise HTTPException(status_code=404, detail="No price data available")

        # 2. Chạy OPS
        print(f"Running OPS Strategy: {strategy}...")
        engine = OPSEngine(prices)
        result = engine.run(strategy=strategy, eta=eta)
        
        # Lấy allocation mới nhất
        latest_alloc = engine.get_latest_allocation(strategy=strategy, eta=eta)
        
        # Format lại dữ liệu Equity Curve để Frontend vẽ chart
        # Chuyển index datetime thành string
        equity_chart = []
        if not result.get("equity_curve", pd.Series()).empty:
            s = result["equity_curve"]
            equity_chart = [{"date": str(d)[:10], "value": v} for d, v in s.items()]

        return {
            "performance": {
                "total_return_pct": result.get("total_return", 0),
                "final_wealth": result.get("final_wealth", 0)
            },
            "latest_allocation": latest_alloc,
            "equity_curve": equity_chart,
            "meta": {
                "strategy": strategy,
                "period": period,
                "tickers_count": engine.n_assets
            }
        }

    except Exception as e:
        print(f"OPS Error: {e}")
        raise HTTPException(status_code=500, detail=str(e))


if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
