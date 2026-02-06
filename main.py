from fastapi import FastAPI, HTTPException
from fastapi.middleware.cors import CORSMiddleware
from quant_backend.data_loader import DataLoader
from quant_backend.network_engine import FinancialNetwork
from quant_backend.ops_engine import OPSEngine
from quant_backend.ai_oracle import AiOracle
import pandas as pd

app = FastAPI(title="TradingQuant API")

app.add_middleware(
    CORSMiddleware,
    allow_origins=["*"],
    allow_credentials=True,
    allow_methods=["*"],
    allow_headers=["*"],
)

# Global instances
loader = DataLoader()
# Cache một network engine để dùng chung cho Oracle
# Lưu ý: Trong production nên dùng Redis hoặc lru_cache
global_network = None

def get_or_create_network(period="6mo"):
    global global_network
    if global_network is None:
        print("Initializing Global Network...")
        prices = loader.get_close_price_matrix(period=period)
        if not prices.empty:
            global_network = FinancialNetwork(prices)
            global_network.build_network(threshold=0.5)
    return global_network

@app.get("/")
def read_root():
    return {"status": "ok", "service": "TradingQuant AI Engine"}

@app.get("/api/analyze-network")
def analyze_market_network(period: str = "6mo", threshold: float = 0.5):
    try:
        prices = loader.get_close_price_matrix(period=period)
        if prices.empty: raise HTTPException(status_code=404, detail="No data")
            
        net_engine = FinancialNetwork(prices)
        net_engine.build_network(threshold=threshold)
        
        # Cập nhật global để Oracle dùng ké
        global global_network
        global_network = net_engine
        
        # 1. Cơ bản
        stats = net_engine.analyze_centrality()
        graph = net_engine.export_json_for_d3()
        
        # 2. Nâng cao: Tính Momentum Spillover
        spillover = net_engine.compute_spillover_momentum()
        
        # Merge kết quả
        # Stats là DataFrame index=Ticker, Spillover cũng vậy
        combined_stats = stats.join(spillover)
        stats_dict = combined_stats.reset_index().rename(columns={'index': 'Ticker'}).to_dict(orient='records')
        
        return {
            "graph": graph,
            "market_stats": stats_dict,
            "meta": {"period": period}
        }
    except Exception as e:
        print(f"Err: {e}")
        raise HTTPException(status_code=500, detail=str(e))

@app.get("/api/oracle/{ticker}")
def ask_oracle(ticker: str):
    """
    Hỏi ý kiến AI Oracle về mã này.
    Ví dụ: HPG.VN, VNM.VN
    """
    # Đảm bảo network đã có dữ liệu để so sánh
    net = get_or_create_network()
    if net is None:
        raise HTTPException(status_code=503, detail="System initializing...")
        
    oracle = AiOracle(loader, net)
    result = oracle.ask(ticker)
    return result

@app.get("/api/optimize-portfolio")
def optimize_portfolio(strategy: str = "EG", eta: float = 0.05, period: str = "6mo"):
    try:
        prices = loader.get_close_price_matrix(period=period)
        if prices.empty: raise HTTPException(status_code=404, detail="No data")

        engine = OPSEngine(prices)
        result = engine.run(strategy=strategy, eta=eta)
        alloc = engine.get_latest_allocation(strategy=strategy, eta=eta)
        
        chart = []
        if not result.get("equity_curve", pd.Series()).empty:
            chart = [{"date": str(d)[:10], "value": v} for d, v in result["equity_curve"].items()]

        return {
            "performance": {
                "return": result.get("total_return", 0),
                "wealth": result.get("final_wealth", 0)
            },
            "allocation": alloc,
            "chart": chart
        }
    except Exception as e:
        raise HTTPException(status_code=500, detail=str(e))

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
