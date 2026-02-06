import React from 'react';
import { Facebook, Instagram, MapPin, Phone, Clock } from 'lucide-react';

const Footer = () => {
    return (
        <footer className="bg-stone-900 text-stone-400 py-16">
            <div className="container mx-auto px-4 grid md:grid-cols-3 gap-12">
                <div>
                    <h3 className="font-serif text-3xl text-white font-bold mb-6">THIARA <span className="text-base font-sans font-light text-rose-400">Home Spa</span></h3>
                    <p className="mb-6 text-stone-400 leading-relaxed">
                        Chuyên gia chăm sóc da khoa học & cá nhân hóa ngay tại khu đô thị của bạn.
                        Nơi vẻ đẹp bắt nguồn từ sự thấu hiểu và an yên.
                    </p>
                    <div className="flex gap-4">
                        <a href="#" className="w-10 h-10 rounded-full bg-stone-800 flex items-center justify-center hover:bg-rose-600 hover:text-white transition-all duration-300 group">
                            <Facebook size={18} className="group-hover:scale-110 transition-transform" />
                        </a>
                        <a href="#" className="w-10 h-10 rounded-full bg-stone-800 flex items-center justify-center hover:bg-rose-600 hover:text-white transition-all duration-300 group">
                            <Instagram size={18} className="group-hover:scale-110 transition-transform" />
                        </a>
                    </div>
                </div>

                <div>
                    <h4 className="text-white font-bold text-lg mb-6 font-serif">Liên Hệ</h4>
                    <ul className="space-y-4">
                        <li className="flex items-start gap-4 group">
                            <div className="w-8 h-8 rounded-full bg-stone-800 flex items-center justify-center shrink-0 group-hover:bg-rose-900 transition-colors">
                                <MapPin size={16} className="text-rose-500 group-hover:text-white transition-colors" />
                            </div>
                            <span className="group-hover:text-stone-200 transition-colors">Tòa Park 5, KĐT [Tên Khu Đô Thị], [Quận/Thành Phố]</span>
                        </li>
                        <li className="flex items-center gap-4 group">
                            <div className="w-8 h-8 rounded-full bg-stone-800 flex items-center justify-center shrink-0 group-hover:bg-rose-900 transition-colors">
                                <Phone size={16} className="text-rose-500 group-hover:text-white transition-colors" />
                            </div>
                            <span className="group-hover:text-stone-200 transition-colors">090.xxx.xxxx (Zalo/Call)</span>
                        </li>
                        <li className="flex items-center gap-4 group">
                            <div className="w-8 h-8 rounded-full bg-stone-800 flex items-center justify-center shrink-0 group-hover:bg-rose-900 transition-colors">
                                <Clock size={16} className="text-rose-500 group-hover:text-white transition-colors" />
                            </div>
                            <span className="group-hover:text-stone-200 transition-colors">09:00 - 20:00 (Đặt lịch trước)</span>
                        </li>
                    </ul>
                </div>

                <div>
                    <h4 className="text-white font-bold text-lg mb-6 font-serif">Chính Sách</h4>
                    <ul className="space-y-3">
                        <li><a href="#" className="hover:text-rose-500 transition-colors flex items-center gap-2"><span className="w-1 h-1 bg-rose-500 rounded-full"></span> Quy trình đặt lịch</a></li>
                        <li><a href="#" className="hover:text-rose-500 transition-colors flex items-center gap-2"><span className="w-1 h-1 bg-rose-500 rounded-full"></span> Chính sách bảo hành da</a></li>
                        <li><a href="#" className="hover:text-rose-500 transition-colors flex items-center gap-2"><span className="w-1 h-1 bg-rose-500 rounded-full"></span> Cam kết riêng tư</a></li>
                        <li><a href="#" className="hover:text-rose-500 transition-colors flex items-center gap-2"><span className="w-1 h-1 bg-rose-500 rounded-full"></span> Câu hỏi thường gặp</a></li>
                    </ul>
                </div>
            </div>
            <div className="container mx-auto px-4 mt-16 pt-8 border-t border-stone-800 text-center text-stone-500 text-sm">
                <p>© 2025 Thiara Home Spa. All rights reserved.</p>
            </div>
        </footer>
    );
};

export default Footer;
