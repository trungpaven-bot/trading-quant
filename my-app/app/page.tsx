"use client"
import { useState, useEffect } from "react"
import { Card, CardContent, CardHeader, CardTitle, CardDescription } from "@/components/ui/card"
import { Button } from "@/components/ui/button"
import { Input } from "@/components/ui/input" // Giả sử đã có component Input basic
import { Search, Globe, ChevronRight, Activity, TrendingUp, TrendingDown, DollarSign } from "lucide-react"

// Component Reusable
const StatCard = ({ title, value, sub, color = "text-slate-900" }: any) => (
  <Card>
    <CardContent className="p-6">
      <div className="text-sm font-medium text-slate-500 mb-1">{title}</div>
      <div className={`text-2xl font-bold ${color}`}>{value}</div>
      <div className="text-xs text-slate-400 mt-1">{sub}</div>
    </CardContent>
  </Card>
)

export default function Dashboard() {
  const [ticker, setTicker] = useState("")
  const [oracleResult, setOracleResult] = useState<any>(null)
  const [loadingOracle, setLoadingOracle] = useState(false)

  const [ntfTickers, setNtfTickers] = useState("BTC-USD, ETH-USD, HPG.VN, FPT.VN")
  const [ntfLookback, setNtfLookback] = useState(20)
  const [ntfResult, setNtfResult] = useState<any[]>([])
  const [loadingNtf, setLoadingNtf] = useState(false)

  // API URL từ biến môi trường
  const API_URL = process.env.NEXT_PUBLIC_API_URL || "http://localhost:8000"

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

  return (
    <div className="space-y-6">
      <div className="flex items-center justify-between">
        <div>
          <h1 className="text-2xl font-bold tracking-tight text-slate-900">TradingQuant Dashboard</h1>
          <p className="text-slate-500">Phân tích thị trường & Tối ưu danh mục đầu tư.</p>
        </div>
        <div className="flex gap-2">
          <div className="px-3 py-1 bg-green-100 text-green-700 rounded-full text-xs font-bold flex items-center">
            <span className="w-2 h-2 bg-green-500 rounded-full mr-2 animate-pulse"></span>
            Market Open
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
              <input
                className="flex h-10 w-full rounded-md border border-input bg-background px-3 py-2 text-sm ring-offset-background file:border-0 file:bg-transparent file:text-sm file:font-medium placeholder:text-muted-foreground focus-visible:outline-none focus-visible:ring-2 focus-visible:ring-ring focus-visible:ring-offset-2 disabled:cursor-not-allowed disabled:opacity-50"
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
                <input
                  type="number"
                  className="flex h-10 w-full rounded-md border border-input bg-background px-3 py-2 text-sm"
                  value={ntfLookback}
                  onChange={(e) => setNtfLookback(parseInt(e.target.value))}
                />
              </div>
              <Button onClick={handleRunNtf} disabled={loadingNtf} className="flex-1 mt-4 bg-indigo-600 hover:bg-indigo-700 text-white">
                {loadingNtf ? "Đang quét..." : "Quét Xu Hướng (Live)"}
              </Button>
            </div>

            {/* Result Table */}
            {ntfResult.length > 0 && (
              <div className="mt-4 border rounded-lg overflow-hidden">
                <table className="w-full text-sm text-left">
                  <thead className="bg-slate-50 text-slate-500 font-medium border-b">
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

      </div>
    </div>
  )
}
