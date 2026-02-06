import React from 'react';
import { Clock } from 'lucide-react';
import { SERVICES } from '../../data/mockData';
import { motion } from 'framer-motion';

const Services = ({ onBookNow }) => {
    return (
        <section id="services" className="py-24 bg-stone-50">
            <div className="container mx-auto px-4">
                <div className="text-center mb-16">
                    <motion.h2
                        initial={{ opacity: 0, y: 20 }}
                        whileInView={{ opacity: 1, y: 0 }}
                        viewport={{ once: true }}
                        className="font-serif text-4xl md:text-5xl font-bold text-stone-900 mb-6"
                    >
                        Dịch Vụ Trị Liệu
                    </motion.h2>
                    <motion.p
                        initial={{ opacity: 0, y: 20 }}
                        whileInView={{ opacity: 1, y: 0 }}
                        viewport={{ once: true }}
                        transition={{ delay: 0.2 }}
                        className="text-stone-600 max-w-2xl mx-auto text-lg"
                    >
                        Các liệu trình được thiết kế tinh gọn, tập trung vào hiệu quả thực tế cho các vấn đề da phổ biến của phụ nữ hiện đại.
                    </motion.p>
                </div>

                <div className="grid md:grid-cols-3 gap-10">
                    {SERVICES.map((service, index) => (
                        <motion.div
                            key={service.id}
                            initial={{ opacity: 0, y: 30 }}
                            whileInView={{ opacity: 1, y: 0 }}
                            viewport={{ once: true }}
                            transition={{ duration: 0.5, delay: index * 0.1 }}
                            className="bg-white rounded-2xl overflow-hidden shadow-sm hover:shadow-2xl transition-all duration-500 group flex flex-col"
                        >
                            <div className="relative h-64 overflow-hidden">
                                <img
                                    src={service.image}
                                    alt={service.name}
                                    className="w-full h-full object-cover transition-transform duration-700 group-hover:scale-110"
                                />
                                {service.tag && (
                                    <span className="absolute top-4 right-4 bg-rose-600 text-white text-xs font-bold px-4 py-1.5 rounded-full shadow-lg tracking-wide uppercase">
                                        {service.tag}
                                    </span>
                                )}
                                <div className="absolute inset-0 bg-gradient-to-t from-black/50 to-transparent opacity-0 group-hover:opacity-100 transition-opacity duration-300"></div>
                            </div>
                            <div className="p-8 flex-1 flex flex-col">
                                <h3 className="text-2xl font-bold text-stone-900 mb-3 font-serif leading-tight group-hover:text-rose-700 transition-colors">
                                    {service.name}
                                </h3>
                                <div className="flex justify-between items-center mb-5 pb-5 border-b border-stone-100">
                                    <span className="text-rose-600 font-bold text-xl">{service.price}</span>
                                    <span className="text-stone-500 text-sm flex items-center gap-1 bg-stone-100 px-3 py-1 rounded-full">
                                        <Clock size={14} /> {service.duration}
                                    </span>
                                </div>
                                <p className="text-stone-600 text-sm mb-8 line-clamp-3 leading-relaxed flex-1">
                                    {service.description}
                                </p>
                                <button
                                    onClick={onBookNow}
                                    className="w-full border-2 border-rose-600 text-rose-600 py-3 rounded-xl font-bold hover:bg-rose-600 hover:text-white transition-all duration-300 uppercase text-sm tracking-wider"
                                >
                                    Đặt Lịch Ngay
                                </button>
                            </div>
                        </motion.div>
                    ))}
                </div>
            </div>
        </section>
    );
};

export default Services;
