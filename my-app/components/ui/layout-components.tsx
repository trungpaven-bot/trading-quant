import Link from "next/link";
import { LayoutDashboard, LineChart, PieChart, Wallet, BookOpen, Settings, Search, User } from "lucide-react";
import { Button } from "@/components/ui/button";

export function Sidebar() {
    return (
        <div className="w-64 border-r bg-white h-full hidden lg:flex flex-col fixed top-0 left-0 bg-slate-50">
            <div className="p-6 h-16 flex items-center border-b">
                <div className="w-8 h-8 bg-black text-white rounded-md flex items-center justify-center font-bold mr-2">TL</div>
                <span className="font-bold text-xl tracking-tight">TitanLabs</span>
            </div>

            <div className="flex-1 py-6 px-3 gap-1 flex flex-col">
                <Link href="/">
                    <Button variant="ghost" className="w-full justify-start text-slate-600 hover:text-slate-900 hover:bg-slate-100 font-medium">
                        <LayoutDashboard className="mr-2 h-4 w-4" />
                        Trang chủ
                    </Button>
                </Link>
                <Link href="/market">
                    <Button variant="ghost" className="w-full justify-start text-slate-600 hover:text-slate-900 hover:bg-slate-100 font-medium">
                        <LineChart className="mr-2 h-4 w-4" />
                        PT Thị trường
                    </Button>
                </Link>
                <Link href="/valuation">
                    <Button variant="ghost" className="w-full justify-start text-slate-600 hover:text-slate-900 hover:bg-slate-100 font-medium">
                        <PieChart className="mr-2 h-4 w-4" />
                        Định giá
                    </Button>
                </Link>
                <Link href="/knowledge">
                    <Button variant="ghost" className="w-full justify-start text-slate-600 hover:text-slate-900 hover:bg-slate-100 font-medium">
                        <BookOpen className="mr-2 h-4 w-4" />
                        Kiến thức
                    </Button>
                </Link>
            </div>

            <div className="p-4 border-t space-y-2">
                <div className="bg-gradient-to-br from-amber-50 to-orange-50 border border-amber-100 p-4 rounded-xl">
                    <div className="flex items-center gap-2 mb-2">
                        <span className="text-xs font-bold text-amber-800 uppercase">Nâng cấp VIP</span>
                    </div>
                    <p className="text-xs text-amber-700 mb-3">Mở khóa toàn bộ báo cáo phân tích chuyên sâu.</p>
                    <Button size="sm" className="w-full bg-gradient-to-r from-amber-500 to-orange-500 text-white border-0 hover:from-amber-600 hover:to-orange-600">Đăng ký ngay</Button>
                </div>
                <Button variant="ghost" className="w-full justify-start text-slate-600">
                    <Settings className="mr-2 h-4 w-4" /> Cài đặt
                </Button>
            </div>
        </div>
    );
}

export function Header() {
    return (
        <header className="h-16 bg-white border-b flex items-center justify-between px-6 sticky top-0 z-40 lg:ml-64">
            <div className="flex-1 max-w-xl">
                <div className="relative">
                    <Search className="absolute left-3 top-1/2 -translate-y-1/2 h-4 w-4 text-slate-400" />
                    <input
                        type="text"
                        placeholder="Tìm kiếm mã cổ phiếu..."
                        className="w-full pl-10 pr-4 py-2 rounded-lg border border-slate-200 bg-slate-50 focus:outline-none focus:ring-2 focus:ring-primary/20 transition-all text-sm"
                    />
                </div>
            </div>
            <div className="flex items-center gap-4">
                <Button size="icon" variant="ghost" className="rounded-full">
                    <User className="h-5 w-5 text-slate-600" />
                </Button>
            </div>
        </header>
    )
}
