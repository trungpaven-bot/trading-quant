import React from 'react';
import { User, Sparkles, ShieldCheck } from 'lucide-react';
import { motion } from 'framer-motion';

const Features = () => {
    const features = [
        {
            icon: <User size={32} />,
            title: "Chuyên Gia 1:1",
            description: "Trực tiếp chủ spa với kinh nghiệm chuyên sâu thăm khám và thực hiện liệu trình cho bạn."
        },
        {
            icon: <Sparkles size={32} />,
            title: "Không Gian Riêng Tư",
            description: "Mô hình Home Spa tại căn hộ cao cấp, chỉ nhận 1 khách/khung giờ. Yên tĩnh tuyệt đối."
        },
        {
            icon: <ShieldCheck size={32} />,
            title: "Phác Đồ Chuẩn Y Khoa",
            description: "Sử dụng dược mỹ phẩm chính hãng. Không xâm lấn, không nghỉ dưỡng, hiệu quả bền vững."
        }
    ];

    return (
        <section className="py-20 bg-white">
            <div className="container mx-auto px-4">
                <div className="grid md:grid-cols-3 gap-8">
                    {features.map((feature, index) => (
                        <motion.div
                            key={index}
                            initial={{ opacity: 0, y: 20 }}
                            whileInView={{ opacity: 1, y: 0 }}
                            viewport={{ once: true }}
                            transition={{ duration: 0.5, delay: index * 0.2 }}
                            className="text-center p-8 rounded-2xl bg-rose-50/30 hover:bg-rose-50 transition-colors border border-transparent hover:border-rose-100"
                        >
                            <div className="w-16 h-16 bg-rose-100 rounded-full flex items-center justify-center mx-auto mb-6 text-rose-600 shadow-inner">
                                {feature.icon}
                            </div>
                            <h3 className="text-xl font-bold mb-3 text-stone-900 font-serif">{feature.title}</h3>
                            <p className="text-stone-600 leading-relaxed">{feature.description}</p>
                        </motion.div>
                    ))}
                </div>
            </div>
        </section>
    );
};

export default Features;
