import React, { useState } from 'react';
import { X, CheckCircle, Clock, ShieldCheck, ChevronRight, ChevronLeft, Calendar } from 'lucide-react';
import { SERVICES, TIME_SLOTS } from '../../data/mockData';
import { motion, AnimatePresence } from 'framer-motion';
import clsx from 'clsx';

const BookingModal = ({ isOpen, onClose }) => {
    const [step, setStep] = useState(1);
    const [selectedService, setSelectedService] = useState(null);
    const [selectedDate, setSelectedDate] = useState('');
    const [selectedTime, setSelectedTime] = useState('');
    const [customerInfo, setCustomerInfo] = useState({ name: '', phone: '', notes: '' });
    const [isSubmitting, setIsSubmitting] = useState(false);

    if (!isOpen) return null;

    const handleNext = () => setStep(step + 1);
    const handleBack = () => setStep(step - 1);

    const handleSubmit = (e) => {
        e.preventDefault();
        setIsSubmitting(true);
        // Simulate API call
        setTimeout(() => {
            setIsSubmitting(false);
            setStep(4); // Success step
        }, 1500);
    };

    const getNext7Days = () => {
        const days = [];
        for (let i = 0; i < 7; i++) {
            const d = new Date();
            d.setDate(d.getDate() + i);
            days.push({
                date: d.toISOString().split('T')[0],
                dayName: new Intl.DateTimeFormat('vi-VN', { weekday: 'short' }).format(d),
                dayNum: d.getDate()
            });
        }
        return days;
    };

    return (
        <div className="fixed inset-0 z-[60] flex items-center justify-center p-4 bg-stone-900/60 backdrop-blur-sm animate-fade-in">
            <motion.div
                initial={{ opacity: 0, scale: 0.95 }}
                animate={{ opacity: 1, scale: 1 }}
                exit={{ opacity: 0, scale: 0.95 }}
                className="bg-white rounded-3xl w-full max-w-2xl max-h-[90vh] overflow-hidden shadow-2xl flex flex-col"
            >
                {/* Header Modal */}
                <div className="p-5 border-b border-stone-100 flex justify-between items-center bg-rose-50/50">
                    <div>
                        <h3 className="font-bold text-rose-900 text-xl font-serif">
                            {step === 1 && "Bước 1: Chọn Liệu Trình"}
                            {step === 2 && "Bước 2: Chọn Thời Gian"}
                            {step === 3 && "Bước 3: Thông Tin Của Bạn"}
                            {step === 4 && "Hoàn Tất Đặt Lịch"}
                        </h3>
                        <div className="flex gap-2 mt-2">
                            {[1, 2, 3].map(i => (
                                <div key={i} className={clsx("h-1 rounded-full transition-all duration-300", i <= step ? "w-8 bg-rose-500" : "w-2 bg-rose-200")}></div>
                            ))}
                        </div>
                    </div>
                    <button onClick={onClose} className="text-stone-400 hover:text-rose-600 p-2 rounded-full hover:bg-rose-100 transition-colors">
                        <X size={24} />
                    </button>
                </div>

                {/* Body Modal */}
                <div className="flex-1 overflow-y-auto p-6 md:p-8 custom-scrollbar">
                    <AnimatePresence mode="wait">
                        {/* STEP 1: CHỌN DỊCH VỤ */}
                        {step === 1 && (
                            <motion.div
                                key="step1"
                                initial={{ opacity: 0, x: 20 }}
                                animate={{ opacity: 1, x: 0 }}
                                exit={{ opacity: 0, x: -20 }}
                                className="space-y-4"
                            >
                                {SERVICES.map((service) => (
                                    <div
                                        key={service.id}
                                        onClick={() => setSelectedService(service)}
                                        className={clsx(
                                            "border rounded-2xl p-4 cursor-pointer transition-all flex gap-5 items-start group relative overflow-hidden",
                                            selectedService?.id === service.id
                                                ? "border-rose-500 bg-rose-50/50 ring-1 ring-rose-500 shadow-md"
                                                : "border-stone-200 hover:border-rose-300 hover:shadow-lg"
                                        )}
                                    >
                                        <img src={service.image} alt={service.name} className="w-24 h-24 object-cover rounded-xl hidden sm:block shadow-sm" />
                                        <div className="flex-1 z-10">
                                            <div className="flex justify-between items-start mb-2">
                                                <h4 className="font-bold text-stone-900 text-lg font-serif group-hover:text-rose-700 transition-colors">{service.name}</h4>
                                                <span className="font-bold text-rose-600 bg-white px-3 py-1 rounded-full text-sm shadow-sm border border-rose-100">{service.price}</span>
                                            </div>
                                            <p className="text-sm text-stone-500 mb-3 flex items-center gap-2 font-medium">
                                                <Clock size={14} className="text-rose-400" /> {service.duration}
                                            </p>
                                            <p className="text-sm text-stone-600 line-clamp-2 leading-relaxed">{service.description}</p>
                                        </div>
                                        <div className={clsx(
                                            "w-6 h-6 rounded-full border-2 flex items-center justify-center mt-1 transition-all z-10",
                                            selectedService?.id === service.id ? "border-rose-500 bg-rose-500" : "border-stone-300 group-hover:border-rose-400"
                                        )}>
                                            {selectedService?.id === service.id && <CheckCircle size={14} className="text-white" />}
                                        </div>
                                    </div>
                                ))}
                            </motion.div>
                        )}

                        {/* STEP 2: CHỌN NGÀY GIỜ */}
                        {step === 2 && (
                            <motion.div
                                key="step2"
                                initial={{ opacity: 0, x: 20 }}
                                animate={{ opacity: 1, x: 0 }}
                                exit={{ opacity: 0, x: -20 }}
                            >
                                <h4 className="font-bold mb-4 text-stone-800 flex items-center gap-2"><Calendar size={18} className="text-rose-500" /> Chọn Ngày</h4>
                                <div className="flex gap-3 overflow-x-auto pb-4 mb-8 no-scrollbar px-1">
                                    {getNext7Days().map((item) => (
                                        <button
                                            key={item.date}
                                            onClick={() => setSelectedDate(item.date)}
                                            className={clsx(
                                                "min-w-[80px] p-4 rounded-2xl flex flex-col items-center border transition-all shadow-sm",
                                                selectedDate === item.date
                                                    ? "bg-rose-600 text-white border-rose-600 shadow-rose-200 shadow-lg transform -translate-y-1"
                                                    : "border-stone-200 text-stone-600 hover:border-rose-300 hover:bg-rose-50"
                                            )}
                                        >
                                            <span className="text-xs uppercase font-bold mb-2 opacity-80">{item.dayName}</span>
                                            <span className="text-2xl font-serif font-bold">{item.dayNum}</span>
                                        </button>
                                    ))}
                                </div>

                                <h4 className="font-bold mb-4 text-stone-800 flex items-center gap-2"><Clock size={18} className="text-rose-500" /> Chọn Khung Giờ</h4>
                                <div className="grid grid-cols-3 sm:grid-cols-4 gap-4">
                                    {TIME_SLOTS.map((time) => (
                                        <button
                                            key={time}
                                            onClick={() => setSelectedTime(time)}
                                            className={clsx(
                                                "py-3 px-4 rounded-xl text-sm font-bold border transition-all shadow-sm",
                                                selectedTime === time
                                                    ? "bg-rose-600 text-white border-rose-600 shadow-md"
                                                    : "border-stone-200 text-stone-700 hover:border-rose-300 hover:bg-rose-50 hover:text-rose-700"
                                            )}
                                        >
                                            {time}
                                        </button>
                                    ))}
                                </div>

                                <div className="mt-8 bg-blue-50/50 p-4 rounded-xl flex items-start gap-3 text-sm text-blue-800 border border-blue-100">
                                    <ShieldCheck size={20} className="mt-0.5 shrink-0 text-blue-600" />
                                    <p className="leading-relaxed">Thiara cam kết chỉ nhận <strong>1 khách hàng</strong> trong khung giờ này để đảm bảo sự riêng tư tuyệt đối cho bạn.</p>
                                </div>
                            </motion.div>
                        )}

                        {/* STEP 3: THÔNG TIN */}
                        {step === 3 && (
                            <motion.div
                                key="step3"
                                initial={{ opacity: 0, x: 20 }}
                                animate={{ opacity: 1, x: 0 }}
                                exit={{ opacity: 0, x: -20 }}
                            >
                                <div className="bg-rose-50 p-6 rounded-2xl mb-6 border border-rose-100">
                                    <h4 className="font-bold text-rose-900 mb-3 font-serif text-lg">Xác nhận dịch vụ:</h4>
                                    <p className="text-stone-800 font-bold text-lg mb-2">{selectedService?.name}</p>
                                    <div className="flex flex-wrap gap-4 text-sm text-stone-600 mt-3 pt-3 border-t border-rose-200/50">
                                        <span className="flex items-center gap-2 bg-white px-3 py-1 rounded-full shadow-sm"><Calendar size={14} className="text-rose-500" /> {selectedDate}</span>
                                        <span className="flex items-center gap-2 bg-white px-3 py-1 rounded-full shadow-sm"><Clock size={14} className="text-rose-500" /> {selectedTime}</span>
                                        <span className="flex items-center gap-2 bg-white px-3 py-1 rounded-full shadow-sm font-bold text-rose-600">{selectedService?.price}</span>
                                    </div>
                                </div>

                                <form onSubmit={handleSubmit} className="space-y-5">
                                    <div>
                                        <label className="block text-sm font-bold text-stone-700 mb-2">Họ và tên *</label>
                                        <input
                                            type="text"
                                            required
                                            className="w-full p-4 border border-stone-200 rounded-xl focus:ring-2 focus:ring-rose-500 focus:border-rose-500 outline-none transition-all bg-stone-50 focus:bg-white"
                                            placeholder="Nhập tên của bạn"
                                            value={customerInfo.name}
                                            onChange={(e) => setCustomerInfo({ ...customerInfo, name: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-stone-700 mb-2">Số điện thoại (Zalo) *</label>
                                        <input
                                            type="tel"
                                            required
                                            className="w-full p-4 border border-stone-200 rounded-xl focus:ring-2 focus:ring-rose-500 focus:border-rose-500 outline-none transition-all bg-stone-50 focus:bg-white"
                                            placeholder="Để Thiara gửi nhắc hẹn cho bạn"
                                            value={customerInfo.phone}
                                            onChange={(e) => setCustomerInfo({ ...customerInfo, phone: e.target.value })}
                                        />
                                    </div>
                                    <div>
                                        <label className="block text-sm font-bold text-stone-700 mb-2">Ghi chú về da (Nếu có)</label>
                                        <textarea
                                            className="w-full p-4 border border-stone-200 rounded-xl focus:ring-2 focus:ring-rose-500 focus:border-rose-500 outline-none transition-all bg-stone-50 focus:bg-white h-32 resize-none"
                                            placeholder="Ví dụ: Da đang bị mụn ẩn, dị ứng tôm..."
                                            value={customerInfo.notes}
                                            onChange={(e) => setCustomerInfo({ ...customerInfo, notes: e.target.value })}
                                        ></textarea>
                                    </div>
                                </form>
                            </motion.div>
                        )}

                        {/* STEP 4: SUCCESS */}
                        {step === 4 && (
                            <motion.div
                                key="step4"
                                initial={{ opacity: 0, scale: 0.9 }}
                                animate={{ opacity: 1, scale: 1 }}
                                className="text-center py-8"
                            >
                                <div className="w-24 h-24 bg-green-100 rounded-full flex items-center justify-center mx-auto mb-8 animate-bounce-slow">
                                    <CheckCircle size={48} className="text-green-600" />
                                </div>
                                <h3 className="text-3xl font-bold text-stone-900 mb-4 font-serif">Đặt Lịch Thành Công!</h3>
                                <p className="text-stone-600 mb-8 text-lg max-w-md mx-auto leading-relaxed">
                                    Cảm ơn <strong>{customerInfo.name}</strong> đã tin tưởng Thiara. <br />
                                    Chúng mình đã gửi thông tin xác nhận qua Zalo SĐT <strong>{customerInfo.phone}</strong>.
                                </p>
                                <div className="bg-rose-50 p-6 rounded-2xl inline-block text-left mb-8 border border-rose-100 max-w-md w-full">
                                    <p className="text-base text-rose-800 font-bold mb-3 flex items-center gap-2"><ShieldCheck size={18} /> Lưu ý nhỏ:</p>
                                    <ul className="text-sm text-stone-600 space-y-2 list-disc list-inside marker:text-rose-400">
                                        <li>Vui lòng đến đúng giờ để đảm bảo đủ thời gian liệu trình.</li>
                                        <li>Nếu cần đổi lịch, nhắn tin cho Thiara trước 2 tiếng nhé.</li>
                                    </ul>
                                </div>
                                <button
                                    onClick={onClose}
                                    className="block w-full bg-stone-900 text-white py-4 rounded-xl font-bold hover:bg-stone-800 transition-all shadow-lg"
                                >
                                    Đóng cửa sổ
                                </button>
                            </motion.div>
                        )}
                    </AnimatePresence>
                </div>

                {/* Footer Modal (Buttons) */}
                {step < 4 && (
                    <div className="p-5 border-t border-stone-100 flex gap-4 bg-white z-20">
                        {step > 1 && (
                            <button
                                onClick={handleBack}
                                className="px-6 py-3 border border-stone-300 rounded-xl font-bold text-stone-600 hover:bg-stone-50 transition-colors flex items-center gap-2"
                            >
                                <ChevronLeft size={20} /> Quay lại
                            </button>
                        )}
                        <button
                            onClick={step === 3 ? handleSubmit : handleNext}
                            disabled={(step === 1 && !selectedService) || (step === 2 && (!selectedDate || !selectedTime)) || isSubmitting}
                            className={clsx(
                                "flex-1 py-3 rounded-xl font-bold text-white flex items-center justify-center gap-2 transition-all shadow-lg",
                                ((step === 1 && !selectedService) || (step === 2 && (!selectedDate || !selectedTime)) || isSubmitting)
                                    ? "bg-stone-300 cursor-not-allowed shadow-none"
                                    : "bg-rose-600 hover:bg-rose-700 shadow-rose-200 hover:shadow-rose-300 hover:-translate-y-0.5"
                            )}
                        >
                            {isSubmitting ? 'Đang xử lý...' : (step === 3 ? 'Xác Nhận Đặt Lịch' : 'Tiếp Theo')}
                            {!isSubmitting && <ChevronRight size={20} />}
                        </button>
                    </div>
                )}
            </motion.div>
        </div>
    );
};

export default BookingModal;
