from fastapi import FastAPI, HTTPException, Request
from fastapi.middleware.cors import CORSMiddleware
from quant_backend.data_loader import DataLoader
from quant_backend.network_engine import FinancialNetwork
from quant_backend.ops_engine import OPSEngine
from quant_backend.ai_oracle import AiOracle
import pandas as pd
import yfinance as yf
import numpy as np

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
    return {"status": "ok", "message": "TradingQuant Server is Running", "server_time": pd.Timestamp.now().isoformat()}

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

@app.post("/api/network-trend")
async def network_trend_scanner(request: Request):
    try:
        # 1. Nhận danh sách mã từ Web
        body = await request.json()
        tickers_raw = body.get("tickers", "BTC-USD, ETH-USD") # Chuỗi nhập vào
        lookback = int(body.get("lookback", 20))
        
        # Xử lý chuỗi nhập: Tách dấu phẩy, xóa khoảng trắng
        ticker_list = [t.strip().upper() for t in tickers_raw.split(",") if t.strip()]
        
        if not ticker_list:
            return {"status": "error", "message": "Chưa nhập mã nào!"}

        print(f"🔍 Scanning: {ticker_list} (Lookback: {lookback}d)")
        
        # 2. Tải dữ liệu LIVE từ Yahoo (Chỉ tải đủ số ngày cần thiết)
        # Tải dư ra chút để đảm bảo đủ nến (lookback * 1.5)
        period_str = f"{int(lookback * 2)}d" if lookback < 100 else "1y" # Tải dư ra chút
        
        # Download data
        data = yf.download(ticker_list, period=period_str, progress=False, auto_adjust=True)
        
        results = []
        
        # 3. Tính toán hiệu suất (Performance)
        # Logic xử lý DataFrame của yfinance (khá phức tạp do MultiIndex)
        close_data = pd.DataFrame()
        
        # Trường hợp 1 mã
        if len(ticker_list) == 1:
            if 'Close' in data.columns:
                close_data = data[['Close']].copy()
                close_data.columns = ticker_list # Rename to ticker
            else:
                 # Đôi khi yfinance trả về Series trực tiếp nếu auto_adjust=True? Không, thường là DF.
                 # Dự phòng
                 close_data = data
        else:
             # Trường hợp nhiều mã: Columns là (Price, Ticker) hoặc chỉ Ticker nếu chỉ request Close?
             # Khi request nhiều chỉ số (OHLC), 'Close' là level 0.
             if 'Close' in data.columns:
                 close_data = data['Close']
             else:
                 close_data = data # Có thể user request chỉ Close? (Hiện tại download mặc định lấy all)

        # Loop calculate
        for ticker in ticker_list:
            try:
                series = None
                if ticker in close_data.columns:
                    series = close_data[ticker]
                
                if series is not None and not series.empty:
                    series = series.dropna()
                    if len(series) >= lookback:
                        start_price = float(series.iloc[-lookback])
                        end_price = float(series.iloc[-1])
                        
                        if start_price > 0:
                            change_pct = ((end_price - start_price) / start_price) * 100
                            results.append({"ticker": ticker, "performance": round(change_pct, 2)})
                        else:
                             results.append({"ticker": ticker, "performance": 0, "note": "Giá = 0"})
                    else:
                        results.append({"ticker": ticker, "performance": 0, "note": "Không đủ dữ liệu"})
                else:
                    results.append({"ticker": ticker, "performance": 0, "note": "Không tìm thấy mã"})
            except Exception as e:
                print(f"Err calc {ticker}: {e}")
                results.append({"ticker": ticker, "performance": 0, "note": "Lỗi tính toán"})

        # 4. Sắp xếp từ Tăng mạnh nhất -> Giảm mạnh nhất
        results.sort(key=lambda x: x["performance"], reverse=True)
        
        return {
            "status": "success",
            "data": results,
            "best_performer": results[0]["ticker"] if results else "N/A"
        }

    except Exception as e:
        print(f"❌ SCAN ERROR: {e}")
        return {"status": "error", "message": str(e)}

# ==============================================================================
# PHẦN BỔ SUNG CHO: PORTFOLIO OPTIMIZATION & BACKTEST
# (Thêm vào backend/main.py hoặc thay thế toàn bộ file)
# ==============================================================================

# --- HÀM HỖ TRỢ TÍNH TOÁN TÀI CHÍNH ---
def get_historical_data(tickers, period="1y"):
    try:
        # Tải dữ liệu đóng cửa điều chỉnh
        data = yf.download(tickers, period=period, progress=False, group_by='ticker', auto_adjust=True)
        # Xử lý format trả về của yfinance
        if len(tickers) == 1:
             # Nếu 1 mã, nó trả về DataFrame có cột Close
             cols = data.columns
             if 'Close' in cols: return data['Close'].to_frame(name=tickers[0])
             else: return data
        else:
             # Nếu nhiều mã, data là MultiIndex (Ticker, OHLCV) hoặc (OHLCV, Ticker) tùy version
             # Thường là columns Level 0 = Price Type hoặc Ticker
             # Cách an toàn nhất là lấy Close của từng thằng
             df_close = pd.DataFrame()
             for t in tickers:
                 try:
                     # Check format
                     if isinstance(data.columns, pd.MultiIndex):
                         # Try to extract
                         try: s = data.xs(t, level=0, axis=1)['Close'] 
                         except: s = data['Close'][t]
                     else:
                         # Flat format
                         s = data['Close'][t] # Nếu format cũ
                     df_close[t] = s
                 except: pass
             return df_close
    except Exception as e: 
        print(f"DL Error: {e}")
        return None

def calculate_metrics(returns):
    # Tính Sharpe, Biến động, CAGR
    mean_return = returns.mean() * 252
    cov_matrix = returns.cov() * 252
    volatility = returns.std() * (252 ** 0.5)
    
    # Sharpe (Giả sử risk-free = 0)
    sharpe = mean_return / volatility if volatility > 0 else 0
    return mean_return, volatility, sharpe

# --- G. API TỐI ƯU HÓA DANH MỤC (OPTIMIZE) ---
@app.post("/api/portfolio/optimize")
async def optimize_portfolio_mc(request: Request):
    try:
        body = await request.json()
        tickers_raw = body.get("assets", "HPG, VNM, FPT") # Lấy danh sách mã
        
        # Xử lý chuỗi ticker
        tickers = [t.strip().upper() for t in tickers_raw.split(",") if t.strip()]
        
        # Thêm đuôi .VN nếu là mã Việt Nam (logic đơn giản: 3 chữ cái -> thêm .VN)
        final_tickers = []
        for t in tickers:
            if len(t) == 3 and t.isalpha() and t != "BTC" and t != "ETH" and t != "USD": 
                final_tickers.append(t + ".VN")
            else: 
                final_tickers.append(t)
            
        if len(final_tickers) < 2:
            return {"status": "error", "message": "Cần ít nhất 2 mã để tối ưu hóa."}

        # 1. Tải dữ liệu 1 năm
        df = get_historical_data(final_tickers, period="1y")
        if df is None or df.empty:
            return {"status": "error", "message": "Không tải được dữ liệu."}
            
        # 2. Tính lợi nhuận ngày (Log returns)
        log_ret = np.log(df / df.shift(1)).dropna()
        
        if log_ret.empty:
             return {"status": "error", "message": "Dữ liệu không đủ để tính toán."}

        # 3. Chạy Mô phỏng Monte Carlo (Tìm bộ tỷ trọng tốt nhất)
        num_portfolios = 2000 # Chạy 2000 kịch bản
        best_sharpe = -100
        best_weights = []
        
        num_assets = len(log_ret.columns) # Dùng số cột thực tế tải được
        found_tickers = log_ret.columns.tolist()

        mean_returns = log_ret.mean()
        cov_matrix = log_ret.cov()

        for _ in range(num_portfolios):
            weights = np.random.random(num_assets)
            weights /= np.sum(weights) # Chuẩn hóa để tổng = 1
            
            # Tính toán hiệu suất Portfolio này
            port_return = np.sum(mean_returns * weights) * 252
            port_vol = np.sqrt(np.dot(weights.T, np.dot(cov_matrix * 252, weights)))
            sharpe = port_return / port_vol if port_vol > 0 else 0
            
            if sharpe > best_sharpe:
                best_sharpe = sharpe
                best_weights = weights
        
        # 4. Đóng gói kết quả
        result_weights = {}
        for i, ticker in enumerate(found_tickers):
            result_weights[ticker.replace(".VN", "")] = round(best_weights[i], 2)
            
        return {
            "status": "success",
            "optimal_weights": result_weights,
            "metrics": {
                "expected_return": round(best_sharpe * 0.15 * 100, 2), # %
                "sharpe_ratio": round(best_sharpe, 2)
            }
        }

    except Exception as e:
        print(f"Optimize Error: {e}")
        return {"status": "error", "message": str(e)}

# --- H. API BACKTEST (KIỂM THỬ QUÁ KHỨ) ---
@app.post("/api/portfolio/backtest")
async def backtest_portfolio(request: Request):
    try:
        body = await request.json()
        tickers_raw = body.get("assets", "HPG, VNM")
        # Giả sử nhận weights từ optimization hoặc chia đều
        
        # Xử lý ticker (tương tự trên)
        tickers = [t.strip().upper() for t in tickers_raw.split(",") if t.strip()]
        final_tickers = []
        for t in tickers:
            if len(t) == 3 and t.isalpha() and t != "BTC" and t != "ETH": 
                final_tickers.append(t + ".VN")
            else: 
                final_tickers.append(t)
        
        # 1. Tải dữ liệu dài hơn (3 năm để backtest)
        df = get_historical_data(final_tickers, period="3y")
        if df is None or df.empty: return {"status": "error", "message": "No data"}
        
        # 2. Giả lập Backtest (Chiến lược Buy & Hold chia đều tiền)
        # Chuẩn hóa về 100 điểm bắt đầu
        normalized = (df / df.iloc[0]) * 100
        
        # Tạo đường Equity Curve (Tổng hợp)
        # Nếu có weights từ request thì dùng, không thì chia đều
        normalized['Portfolio'] = normalized.mean(axis=1) # Chia đều tỷ trọng (Simple)
        
        # Lấy dữ liệu để vẽ chart
        dates = normalized.index.strftime('%Y-%m-%d').tolist()
        values = normalized['Portfolio'].tolist()
        
        # Tính Max Drawdown
        rolling_max = normalized['Portfolio'].cummax()
        drawdown = (normalized['Portfolio'] - rolling_max) / rolling_max
        max_dd = drawdown.min() * 100

        return {
            "status": "success",
            "dates": dates,
            "equity_curve": values,
            "metrics": {
                "total_return": round(values[-1] - 100, 2), # %
                "max_drawdown": round(max_dd, 2)
            }
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

if __name__ == "__main__":
    import uvicorn
    uvicorn.run(app, host="0.0.0.0", port=8000)
