"use client"
import { useState, useEffect } from "react"
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input"
import { Search, Globe, ChevronRight, Activity, TrendingUp, TrendingDown, DollarSign, Wallet } from "lucide-react"

export default function Dashboard() {
  const [ticker, setTicker] = useState("")
  const [oracleResult, setOracleResult] = useState<any>(null)
  const [loadingOracle, setLoadingOracle] = useState(false)

  const [ntfTickers, setNtfTickers] = useState("BTC-USD, ETH-USD, HPG.VN, FPT.VN")
  const [ntfLookback, setNtfLookback] = useState(20)
  const [ntfResult, setNtfResult] = useState<any[]>([])
  const [loadingNtf, setLoadingNtf] = useState(false)

  // Portfolio States
  const [portAssets, setPortAssets] = useState("HPG.VN, VNM.VN, FPT.VN, VCB.VN")
  const [maxWeight, setMaxWeight] = useState(0.40) // Mặc định 40%
  const [optResult, setOptResult] = useState<any>(null)
  const [btResult, setBtResult] = useState<any>(null)
  const [loadingPort, setLoadingPort] = useState(false)

  // Server Status State
  const [serverStatus, setServerStatus] = useState<"checking" | "online" | "offline">("checking")

  // API URL từ biến môi trường
  const API_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000"

  // 1. Server Heartbeat Logic
  const checkServerStatus = async () => {
    setServerStatus("checking")
    try {
      // Timeout 8s
      const controller = new AbortController()
      const timeoutId = setTimeout(() => controller.abort(), 8000)

      const res = await fetch(`${API_URL}/`, { signal: controller.signal })
      clearTimeout(timeoutId)

      if (res.ok) {
        setServerStatus("online")
      } else {
        setServerStatus("offline")
      }
    } catch (e) {
      setServerStatus("offline")
    }
  }

  useEffect(() => {
    checkServerStatus()
    // Ping mỗi 60s để giữ kết nối hoặc check trạng thái
    const interval = setInterval(checkServerStatus, 60000)
    return () => clearInterval(interval)
  }, [])

  // 2. Handlers
  const handleAskOracle = async () => {
    if (!ticker) return
    setLoadingOracle(true)
    try {
      const res = await fetch(`${API_URL}/api/oracle/${ticker.toUpperCase()}`)
      const data = await res.json()
      setOracleResult(data)
    } catch (e) {
      console.error(e)
    } finally {
      setLoadingOracle(false)
    }
  }

  const handleRunNtf = async () => {
    setLoadingNtf(true)
    try {
      const res = await fetch(`${API_URL}/api/network-trend`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ tickers: ntfTickers, lookback: ntfLookback })
      })
      const data = await res.json()
      if (data.status === "success") {
        setNtfResult(data.data)
      }
    } catch (e) {
      console.error(e)
    } finally {
      setLoadingNtf(false)
    }
  }

  const handleOptimize = async () => {
    setLoadingPort(true)
    setOptResult(null)
    try {
      const res = await fetch(`${API_URL}/api/portfolio/optimize`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          assets: portAssets,
          max_weight: maxWeight
        })
      })
      const data = await res.json()
      if (data.status === "success") setOptResult(data)
    } catch (e) { console.error(e) }
    finally { setLoadingPort(false) }
  }

  const handleBacktest = async () => {
    setLoadingPort(true)
    setBtResult(null)
    try {
      const res = await fetch(`${API_URL}/api/portfolio/backtest`, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ assets: portAssets })
      })
      const data = await res.json()
      if (data.status === "success") setBtResult(data)
    } catch (e) { console.error(e) }
    finally { setLoadingPort(false) }
  }

  return (
    <div className="space-y-6">
      {/* Header & Server Status */}
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">TradingQuant Dashboard</h1>
          <p className="text-slate-500">Phân tích thị trường & Tối ưu danh mục đầu tư.</p>
        </div>
        <div className="flex gap-2">
          <div
            onClick={checkServerStatus}
            className={`px-3 py-1 rounded-full text-xs font-bold flex items-center cursor-pointer transition-all border select-none ${serverStatus === "online" ? "bg-green-50 text-green-700 border-green-200" :
                serverStatus === "checking" ? "bg-amber-50 text-amber-700 border-amber-200" :
                  "bg-red-50 text-red-700 border-red-200"
              }`}
            title="Bấm để lay server dậy (Ping)"
          >
            <span className={`w-2 h-2 rounded-full mr-2 ${serverStatus === "online" ? "bg-green-500 animate-pulse" :
                serverStatus === "checking" ? "bg-amber-500 animate-bounce" :
                  "bg-red-500"
              }`}></span>
            {serverStatus === "online" ? "Server Online" :
              serverStatus === "checking" ? "Waking up..." : "Server Sleeping"}
          </div>
        </div>
      </div>

      <div className="grid grid-cols-1 md:grid-cols-2 gap-6">

        {/* --- 1. AI ORACLE & FUNDAMENTAL SNAPSHOT --- */}
        <Card className="border-teal-100 shadow-sm">
          <CardHeader className="bg-teal-50/50 border-b border-teal-100 pb-4">
            <div className="flex items-center justify-between">
              <CardTitle className="text-teal-900 flex items-center gap-2">
                <Activity className="h-5 w-5" /> AI Oracle Snapshot
              </CardTitle>
            </div>
            <CardDescription>Soi cơ bản & Tư vấn tín hiệu kỹ thuật nhanh.</CardDescription>
          </CardHeader>
          <CardContent className="p-6 space-y-4">
            <div className="flex gap-2">
              <Input
                placeholder="Nhập mã (VD: HPG.VN, BTC-USD)..."
                value={ticker}
                onChange={(e) => setTicker(e.target.value)}
                onKeyDown={(e) => e.key === 'Enter' && handleAskOracle()}
              />
              <Button onClick={handleAskOracle} disabled={loadingOracle} className="bg-teal-600 hover:bg-teal-700 text-white">
                {loadingOracle ? "Đang soi..." : "Phân tích"}
              </Button>
            </div>

            {oracleResult && (
              <div className="bg-slate-50 p-4 rounded-lg border border-slate-100 space-y-3 animate-in fade-in slide-in-from-top-2">
                <div className="flex justify-between items-start">
                  <div>
                    <h3 className="font-bold text-lg text-slate-800">{oracleResult.ticker}</h3>
                    <div className="text-xs text-slate-500">Dữ liệu Realtime</div>
                  </div>
                  <div className={`px-3 py-1 rounded text-xs font-bold ${oracleResult.technical?.signal?.includes("TĂNG") ? "bg-green-100 text-green-700" :
                      oracleResult.technical?.signal?.includes("GIẢM") ? "bg-red-100 text-red-700" : "bg-gray-100 text-gray-700"
                    }`}>
                    {oracleResult.technical?.signal || "N/A"}
                  </div>
                </div>

                <div className="grid grid-cols-2 gap-4">
                  <div className="p-3 bg-white rounded border">
                    <div className="text-xs text-slate-400 mb-1">Cơ bản (Fundamental)</div>
                    <div className="font-semibold text-sm">P/E: <span className="text-slate-900">{oracleResult.fundamental?.pe}</span></div>
                    <div className="font-semibold text-sm">ROE: <span className="text-slate-900">{oracleResult.fundamental?.roe}</span></div>
                    <div className="text-xs font-bold text-blue-600 mt-1">{oracleResult.fundamental?.signal}</div>
                  </div>
                  <div className="p-3 bg-white rounded border">
                    <div className="text-xs text-slate-400 mb-1">Kỹ thuật (Technical)</div>
                    <div className="font-semibold text-sm">Giá: {oracleResult.technical?.price?.toLocaleString()}</div>
                    <div className="font-semibold text-sm text-slate-500">MA20: {oracleResult.technical?.ma20?.toLocaleString()}</div>
                  </div>
                </div>

                <div className="text-sm bg-blue-50 text-blue-800 p-3 rounded italic">
                  " {oracleResult.full_analysis.split('\n')[0]} ... "
                </div>
              </div>
            )}
          </CardContent>
        </Card>

        {/* --- 2. NETWORK TREND FOLLOWING --- */}
        <Card className="border-indigo-100 shadow-sm">
          <CardHeader className="bg-indigo-50/50 border-b border-indigo-100 pb-4">
            <div className="flex items-center justify-between">
              <CardTitle className="text-indigo-900 flex items-center gap-2">
                <Globe className="h-5 w-5" /> Network Trend Scanner
              </CardTitle>
            </div>
            <CardDescription>So sánh sức mạnh nhóm ngành / List theo dõi.</CardDescription>
          </CardHeader>
          <CardContent className="p-6 space-y-4">
            <div className="space-y-2">
              <label className="text-xs font-medium text-slate-600">Danh sách mã (cách nhau dấu phẩy)</label>
              <textarea
                className="flex min-h-[80px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm ring-offset-background placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-50"
                value={ntfTickers}
                onChange={(e) => setNtfTickers(e.target.value)}
              />
            </div>
            <div className="flex gap-2 items-center">
              <div className="w-1/3">
                <label className="text-xs font-medium text-slate-600">Lookback (ngày)</label>
                <Input
                  type="number"
                  value={ntfLookback}
                  onChange={(e) => setNtfLookback(parseInt(e.target.value))}
                />
              </div>
              <Button onClick={handleRunNtf} disabled={loadingNtf} className="flex-1 mt-6 bg-indigo-600 hover:bg-indigo-700 text-white">
                {loadingNtf ? "Đang quét..." : "Quét Xu Hướng (Live)"}
              </Button>
            </div>

            {ntfResult.length > 0 && (
              <div className="mt-4 border rounded-lg overflow-hidden max-h-60 overflow-y-auto scrollbar-thin">
                <table className="w-full text-sm text-left">
                  <thead className="bg-slate-50 text-slate-500 font-medium border-b sticky top-0">
                    <tr>
                      <th className="px-4 py-2">Mã</th>
                      <th className="px-4 py-2 text-right">Hiệu suất</th>
                      <th className="px-4 py-2">Note</th>
                    </tr>
                  </thead>
                  <tbody className="divide-y divide-slate-100">
                    {ntfResult.map((item, idx) => (
                      <tr key={idx} className="hover:bg-slate-50/50">
                        <td className="px-4 py-2 font-medium">{item.ticker}</td>
                        <td className={`px-4 py-2 text-right font-bold ${item.performance > 0 ? "text-green-600" : "text-red-500"}`}>
                          {item.performance > 0 ? "+" : ""}{item.performance}%
                        </td>
                        <td className="px-4 py-2 text-xs text-slate-400">{item.note || "-"}</td>
                      </tr>
                    ))}
                  </tbody>
                </table>
              </div>
            )}
          </CardContent>
        </Card>

        {/* --- 3. PORTFOLIO OPTIMIZER & BACKTEST --- */}
        <Card className="md:col-span-2 border-orange-100 shadow-sm">
          <CardHeader className="bg-orange-50/50 border-b border-orange-100 pb-4">
            <div className="flex items-center justify-between">
              <CardTitle className="text-orange-900 flex items-center gap-2">
                <Wallet className="h-5 w-5" /> Portfolio Optimizer (Monte Carlo)
              </CardTitle>
            </div>
            <CardDescription>Tìm tỷ trọng tối ưu và kiểm thử (Backtest) danh mục đầu tư.</CardDescription>
          </CardHeader>
          <CardContent className="p-6">
            <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
              <div className="space-y-4">
                <div className="space-y-2">
                  <label className="text-xs font-medium text-slate-600">Assets (Cổ phiếu dự kiến mua)</label>
                  <textarea
                    className="flex min-h-[80px] w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
                    value={portAssets}
                    onChange={(e) => setPortAssets(e.target.value)}
                  />
                  <p className="text-xs text-slate-400">Hỗ trợ cổ phiếu VN, Mỹ & Crypto.</p>
                </div>

                <div className="space-y-2 bg-orange-50 p-3 rounded-lg border border-orange-100">
                  <label className="text-xs font-bold text-orange-800 flex justify-between">
                    <span>Đa dạng hóa (Max Weight)</span>
                    <span>{(maxWeight * 100).toFixed(0)}%</span>
                  </label>
                  <div className="flex items-center gap-2">
                    <input
                      type="range" min="0.1" max="1.0" step="0.05"
                      className="w-full h-2 bg-orange-200 rounded-lg appearance-none cursor-pointer accent-orange-600"
                      value={maxWeight}
                      onChange={(e) => setMaxWeight(parseFloat(e.target.value))}
                    />
                  </div>
                  <p className="text-[10px] text-orange-600 italic">
                    *Không dùng quá {(maxWeight * 100).toFixed(0)}% vốn cho 1 mã.
                  </p>
                </div>

                <div className="flex gap-2">
                  <Button onClick={handleOptimize} disabled={loadingPort} className="flex-1 bg-orange-600 hover:bg-orange-700 text-white">
                    Optimize ⚡
                  </Button>
                  <Button onClick={handleBacktest} disabled={loadingPort} variant="outline" className="flex-1 border-orange-200 text-orange-700 hover:bg-orange-50">
                    Backtest 📉
                  </Button>
                </div>
              </div>

              {/* Result Area */}
              <div className="md:col-span-2 bg-slate-50 rounded-xl border border-dashed border-slate-200 p-4 flex flex-col justify-center items-center min-h-[200px]">
                {!optResult && !btResult && !loadingPort && (
                  <div className="text-slate-400 text-sm text-center">
                    Nhập danh sách mã và bấm nút để bắt đầu mô phỏng.<br />
                    <span className="text-xs opacity-70">Hệ thống sẽ chạy 2,000 kịch bản ngẫu nhiên tìm tỷ trọng tối ưu.</span>
                  </div>
                )}

                {loadingPort && (
                  <div className="flex flex-col items-center animate-in fade-in">
                    <div className="w-8 h-8 border-4 border-orange-500 border-t-transparent rounded-full animate-spin mb-2"></div>
                    <span className="text-sm text-slate-500 animate-pulse">Đang chạy mô phỏng Monte Carlo...</span>
                  </div>
                )}

                {optResult && (
                  <div className="w-full animate-in fade-in zoom-in-95">
                    <h3 className="font-bold text-lg text-slate-800 mb-4 flex items-center gap-2">
                      <TrendingUp className="h-5 w-5 text-green-600" /> Kết quả Tối Ưu Hóa (Sharpe cao nhất)
                    </h3>
                    <div className="grid grid-cols-1 md:grid-cols-2 gap-8">
                      <div>
                        <div className="text-sm font-medium mb-2 text-slate-600">Phân bổ khuyến nghị:</div>
                        <div className="space-y-2">
                          {Object.entries(optResult.optimal_weights).map(([ticker, weight]: any) => (
                            <div key={ticker} className="flex items-center justify-between bg-white p-2 rounded border shadow-sm">
                              <span className="font-bold text-slate-700">{ticker}</span>
                              <div className="flex items-center gap-2">
                                <div className="h-2 bg-orange-200 rounded-full w-20 overflow-hidden">
                                  <div className="h-full bg-orange-500" style={{ width: `${weight * 100}%` }}></div>
                                </div>
                                <span className="font-mono text-orange-600 font-bold min-w-[3rem] text-right">{(weight * 100).toFixed(0)}%</span>
                              </div>
                            </div>
                          ))}
                        </div>
                      </div>
                      <div className="space-y-4">
                        <div className="p-4 bg-white rounded border border-orange-100 shadow-sm flex items-center justify-between">
                          <div>
                            <div className="text-xs text-slate-400">Lợi nhuận kỳ vọng / năm</div>
                            <div className="text-2xl font-bold text-green-600">+{optResult.metrics.expected_return}%</div>
                          </div>
                          <TrendingUp className="h-8 w-8 text-green-100" />
                        </div>
                        <div className="p-4 bg-white rounded border border-slate-100 shadow-sm flex items-center justify-between">
                          <div>
                            <div className="text-xs text-slate-400">Sharpe Ratio (Hiệu quả)</div>
                            <div className="text-xl font-bold text-slate-800">{optResult.metrics.sharpe_ratio}</div>
                          </div>
                          <Activity className="h-8 w-8 text-slate-100" />
                        </div>
                      </div>
                    </div>
                  </div>
                )}

                {btResult && (
                  <div className="w-full animate-in fade-in zoom-in-95">
                    <h3 className="font-bold text-lg text-slate-800 mb-4 flex items-center gap-2">
                      <Activity className="h-5 w-5 text-blue-600" /> Kết quả Backtest (3 Năm qua)
                    </h3>
                    <div className="grid grid-cols-3 gap-4 mb-4">
                      <div className="p-3 bg-white rounded border text-center">
                        <div className="text-xs text-slate-400">Tổng Lợi Nhuận</div>
                        <div className={`text-xl font-bold ${btResult.metrics.total_return > 0 ? "text-green-600" : "text-red-500"}`}>
                          {btResult.metrics.total_return > 0 ? "+" : ""}{btResult.metrics.total_return}%
                        </div>
                      </div>
                      <div className="p-3 bg-white rounded border text-center">
                        <div className="text-xs text-slate-400">Rủi ro (Max DD)</div>
                        <div className="text-xl font-bold text-red-600">{btResult.metrics.max_drawdown}%</div>
                      </div>
                      <div className="p-3 bg-white rounded border text-center">
                        <div className="text-xs text-slate-400">Số phiên</div>
                        <div className="text-xl font-bold text-slate-700">{btResult.equity_curve.length}</div>
                      </div>
                    </div>
                  </div>
                )}
              </div>
            </div>
          </CardContent>
        </Card>

      </div>
    </div>
  )
}
