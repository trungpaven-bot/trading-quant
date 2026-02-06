import React from 'react';
import { Star, Quote } from 'lucide-react';
import { REVIEWS } from '../../data/mockData';
import { motion } from 'framer-motion';

const Testimonials = () => {
    return (
        <section id="reviews" className="py-24 bg-white relative overflow-hidden">
            {/* Decorative Elements */}
            <div className="absolute top-0 left-0 w-64 h-64 bg-rose-50 rounded-full mix-blend-multiply filter blur-3xl opacity-70 -translate-x-1/2 -translate-y-1/2"></div>
            <div className="absolute bottom-0 right-0 w-96 h-96 bg-rose-50 rounded-full mix-blend-multiply filter blur-3xl opacity-70 translate-x-1/3 translate-y-1/3"></div>

            <div className="container mx-auto px-4 relative z-10">
                <div className="text-center mb-16">
                    <h2 className="font-serif text-4xl md:text-5xl font-bold text-stone-900 mb-4">Lời Yêu Thương Từ Khách Hàng</h2>
                    <p className="text-stone-500">Những chia sẻ chân thực nhất về trải nghiệm tại Thiara Home Spa</p>
                </div>

                <div className="grid md:grid-cols-3 gap-8">
                    {REVIEWS.map((review, index) => (
                        <motion.div
                            key={review.id}
                            initial={{ opacity: 0, y: 20 }}
                            whileInView={{ opacity: 1, y: 0 }}
                            viewport={{ once: true }}
                            transition={{ duration: 0.5, delay: index * 0.2 }}
                            className="bg-white p-8 rounded-2xl border border-stone-100 shadow-lg shadow-stone-200/50 relative group hover:-translate-y-1 transition-transform duration-300"
                        >
                            <Quote className="absolute top-6 right-6 text-rose-100 w-12 h-12 group-hover:text-rose-200 transition-colors" />

                            <div className="flex text-yellow-400 mb-6">
                                {[...Array(5)].map((_, i) => <Star key={i} size={18} fill="currentColor" className="drop-shadow-sm" />)}
                            </div>

                            <p className="text-stone-600 italic mb-8 leading-relaxed relative z-10">"{review.content}"</p>

                            <div className="flex items-center gap-4 pt-6 border-t border-stone-50">
                                <div className="w-12 h-12 bg-gradient-to-br from-rose-100 to-rose-200 rounded-full flex items-center justify-center text-rose-700 font-bold text-lg shadow-inner">
                                    {review.name.charAt(0)}
                                </div>
                                <div>
                                    <h4 className="font-bold text-stone-900 text-base font-serif">{review.name}</h4>
                                    <p className="text-xs text-stone-500 uppercase tracking-wide font-medium">{review.role}</p>
                                </div>
                            </div>
                        </motion.div>
                    ))}
                </div>
            </div>
        </section>
    );
};

export default Testimonials;
