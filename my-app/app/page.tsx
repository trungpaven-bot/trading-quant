import { ArrowUpRight, TrendingUp, DollarSign, Activity } from "lucide-react";
import { Button } from "@/components/ui/button";

function MetricCard({ title, value, change, trend }: { title: string, value: string, change: string, trend: "up" | "down" }) {
  return (
    <div className="bg-white p-6 rounded-xl border shadow-sm">
      <div className="flex justify-between items-start mb-2">
        <span className="text-slate-500 text-sm font-medium">{title}</span>
        {trend === "up" ? <TrendingUp className="h-4 w-4 text-green-500" /> : <Activity className="h-4 w-4 text-slate-400" />}
      </div>
      <div className="text-2xl font-bold text-slate-900 mb-1">{value}</div>
      <div className={`text-xs font-medium ${trend === "up" ? "text-green-600" : "text-red-500"} flex items-center`}>
        {trend === "up" ? "+" : ""}{change} so với hôm qua
      </div>
    </div>
  )
}

export default function Home() {
  return (
    <div className="space-y-8">
      {/* Welcome Section */}
      <div className="flex justify-between items-center">
        <div>
          <h1 className="text-2xl font-bold text-slate-900 tracking-tight">Xin chào, Admin 👋</h1>
          <p className="text-slate-500">Đây là tổng quan thị trường hôm nay.</p>
        </div>
        <Button className="bg-emerald-600 hover:bg-emerald-700 text-white shadow-lg shadow-emerald-600/20">
          <ArrowUpRight className="mr-2 h-4 w-4" /> Xuất báo cáo
        </Button>
      </div>

      {/* Metrics Grid */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-6">
        <MetricCard title="VN-Index" value="1,245.32" change="12.5 (1.02%)" trend="up" />
        <MetricCard title="Thanh khoản" value="23.5K Tỷ" change="15%" trend="up" />
        <MetricCard title="Mã tăng" value="320" change="Chiếm 65%" trend="up" />
        <MetricCard title="Mã giảm" value="115" change="Chiếm 23%" trend="down" />
      </div>

      {/* Main Content Area */}
      <div className="grid grid-cols-1 lg:grid-cols-3 gap-8">
        {/* Chart Section (Placeholder) */}
        <div className="lg:col-span-2 bg-white p-6 rounded-xl border shadow-sm min-h-[400px] flex flex-col">
          <h3 className="font-bold text-lg mb-4">Biểu biến động VN30</h3>
          <div className="flex-1 bg-slate-50 rounded-lg flex items-center justify-center border border-dashed border-slate-200">
            <span className="text-slate-400 text-sm">Biểu đồ đang được cập nhật...</span>
          </div>
        </div>

        {/* Top Movers */}
        <div className="bg-white p-6 rounded-xl border shadow-sm flex flex-col">
          <h3 className="font-bold text-lg mb-4">Top Tăng Trưởng (VN30)</h3>
          <div className="space-y-4">
            {[
              { code: "MBB", name: "MB Bank", price: "24,500", change: "+4.2%" },
              { code: "FPT", name: "FPT Corp", price: "102,100", change: "+3.8%" },
              { code: "HPG", name: "Hoa Phat Group", price: "29,300", change: "+2.1%" },
              { code: "SSI", name: "SSI Securities", price: "34,100", change: "+1.9%" },
              { code: "MWG", name: "Mobile World", price: "48,200", change: "+1.5%" },
            ].map((stock) => (
              <div key={stock.code} className="flex items-center justify-between p-3 hover:bg-slate-50 rounded-lg transition-colors cursor-pointer group">
                <div className="flex items-center gap-3">
                  <div className="w-10 h-10 rounded-full bg-slate-100 flex items-center justify-center font-bold text-xs text-slate-600 group-hover:bg-emerald-100 group-hover:text-emerald-700 transition-colors">
                    {stock.code}
                  </div>
                  <div>
                    <div className="font-bold text-sm text-slate-900">{stock.code}</div>
                    <div className="text-xs text-slate-500">{stock.name}</div>
                  </div>
                </div>
                <div className="text-right">
                  <div className="font-medium text-sm">{stock.price}</div>
                  <div className="text-xs font-bold text-emerald-600">{stock.change}</div>
                </div>
              </div>
            ))}
          </div>
          <Button variant="outline" className="w-full mt-auto pt-4 border-t border-0">Xem tất cả</Button>
        </div>
      </div>
    </div>
  );
}
