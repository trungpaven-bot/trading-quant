import React from 'react';
import { motion } from 'framer-motion';

const Hero = ({ onBookNow }) => {
    return (
        <section className="relative h-screen min-h-[600px] flex items-center justify-center overflow-hidden">
            {/* Background Image with Overlay */}
            <div className="absolute inset-0 z-0">
                <img
                    src="https://images.unsplash.com/photo-1540555700478-4be289fbecef?auto=format&fit=crop&q=80&w=2000"
                    alt="Spa Background"
                    className="w-full h-full object-cover scale-105 animate-subtle-zoom"
                />
                <div className="absolute inset-0 bg-black/30 md:bg-black/40 backdrop-blur-[2px]"></div>
                <div className="absolute inset-0 bg-gradient-to-b from-black/30 via-transparent to-stone-900/20"></div>
            </div>

            <div className="container mx-auto px-4 relative z-10 text-center text-white">
                <motion.div
                    initial={{ opacity: 0, y: 20 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.8, delay: 0.2 }}
                >
                    <span className="inline-block py-1 px-4 rounded-full bg-white/10 backdrop-blur-md text-sm font-medium tracking-widest uppercase mb-6 border border-white/20">
                        ✨ Home Spa Chuyên Nghiệp & Riêng Tư
                    </span>
                </motion.div>

                <motion.h1
                    initial={{ opacity: 0, y: 30 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.8, delay: 0.4 }}
                    className="font-serif text-5xl md:text-7xl font-bold mb-6 leading-tight drop-shadow-lg"
                >
                    Đánh Thức Vẻ Đẹp <br /> <span className="text-rose-200 italic">Từ Sự Thấu Hiểu</span>
                </motion.h1>

                <motion.p
                    initial={{ opacity: 0, y: 30 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.8, delay: 0.6 }}
                    className="text-lg md:text-xl mb-10 max-w-2xl mx-auto text-stone-100 font-light leading-relaxed"
                >
                    Không gian trị liệu 1:1 độc quyền. Phác đồ cá nhân hóa từ chuyên gia.
                    Nơi bạn tìm thấy phiên bản rạng rỡ nhất của chính mình.
                </motion.p>

                <motion.div
                    initial={{ opacity: 0, y: 30 }}
                    animate={{ opacity: 1, y: 0 }}
                    transition={{ duration: 0.8, delay: 0.8 }}
                    className="flex flex-col md:flex-row justify-center gap-5"
                >
                    <button
                        onClick={onBookNow}
                        className="bg-rose-600 text-white px-8 py-4 rounded-full text-lg font-semibold hover:bg-rose-700 transition-all shadow-xl shadow-rose-900/20 transform hover:-translate-y-1"
                    >
                        Đặt Lịch Tư Vấn Miễn Phí
                    </button>
                    <a
                        href="#services"
                        className="bg-white/10 backdrop-blur-md border border-white/40 text-white px-8 py-4 rounded-full text-lg font-semibold hover:bg-white hover:text-rose-900 transition-all shadow-lg"
                    >
                        Xem Dịch Vụ
                    </a>
                </motion.div>
            </div>
        </section>
    );
};

export default Hero;
