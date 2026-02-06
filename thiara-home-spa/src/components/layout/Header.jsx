import React, { useState, useEffect } from 'react';
import { Menu, X } from 'lucide-react';
import { motion, AnimatePresence } from 'framer-motion';
import clsx from 'clsx';

const Header = ({ onBookNow }) => {
    const [isScrolled, setIsScrolled] = useState(false);
    const [isMobileMenuOpen, setIsMobileMenuOpen] = useState(false);

    useEffect(() => {
        const handleScroll = () => setIsScrolled(window.scrollY > 20);
        window.addEventListener('scroll', handleScroll);
        return () => window.removeEventListener('scroll', handleScroll);
    }, []);

    const navLinks = [
        { name: 'Về Thiara', href: '#about' },
        { name: 'Dịch Vụ', href: '#services' },
        { name: 'Khách Hàng', href: '#reviews' },
    ];

    return (
        <header
            className={clsx(
                "fixed top-0 w-full z-50 transition-all duration-300 border-b border-transparent",
                isScrolled
                    ? "bg-white/80 backdrop-blur-md py-3 shadow-sm border-rose-100/50"
                    : "bg-transparent py-5"
            )}
        >
            <div className="container mx-auto px-4 flex justify-between items-center">
                <div className={clsx(
                    "font-serif text-2xl font-bold tracking-wider transition-colors",
                    isScrolled ? "text-rose-900" : "text-rose-900 md:text-white"
                )}>
                    THIARA <span className="text-sm font-sans font-light block md:inline opacity-90">Home Spa</span>
                </div>

                {/* Desktop Nav */}
                <nav className="hidden md:flex items-center space-x-8">
                    {navLinks.map((link) => (
                        <a
                            key={link.name}
                            href={link.href}
                            className={clsx(
                                "text-sm font-medium tracking-wide transition-colors hover:text-rose-500",
                                isScrolled ? "text-stone-600" : "text-white/90 hover:text-white"
                            )}
                        >
                            {link.name}
                        </a>
                    ))}
                    <button
                        onClick={onBookNow}
                        className={clsx(
                            "px-6 py-2 rounded-full font-medium transition-all transform hover:scale-105 shadow-lg",
                            isScrolled
                                ? "bg-rose-600 text-white hover:bg-rose-700 shadow-rose-200"
                                : "bg-white text-rose-900 hover:bg-rose-50 shadow-black/10"
                        )}
                    >
                        Đặt Lịch Ngay
                    </button>
                </nav>

                {/* Mobile Menu Button */}
                <button
                    className="md:hidden text-rose-900 z-50 relative"
                    onClick={() => setIsMobileMenuOpen(!isMobileMenuOpen)}
                >
                    {isMobileMenuOpen ? <X size={28} className="text-stone-800" /> : <Menu size={28} className={isScrolled ? 'text-rose-900' : 'text-rose-900 md:text-white'} />}
                </button>

                {/* Mobile Nav Overlay */}
                <AnimatePresence>
                    {isMobileMenuOpen && (
                        <motion.div
                            initial={{ opacity: 0, y: -20 }}
                            animate={{ opacity: 1, y: 0 }}
                            exit={{ opacity: 0, y: -20 }}
                            className="absolute top-0 left-0 w-full bg-white shadow-xl pt-24 pb-8 px-6 flex flex-col space-y-4 md:hidden border-b border-rose-100"
                        >
                            {navLinks.map((link) => (
                                <a
                                    key={link.name}
                                    href={link.href}
                                    className="text-stone-800 text-lg font-medium py-3 border-b border-stone-100"
                                    onClick={() => setIsMobileMenuOpen(false)}
                                >
                                    {link.name}
                                </a>
                            ))}
                            <button
                                onClick={() => { onBookNow(); setIsMobileMenuOpen(false); }}
                                className="bg-rose-600 text-white w-full py-4 rounded-xl font-bold text-lg shadow-lg shadow-rose-200 mt-4"
                            >
                                Đặt Lịch Ngay
                            </button>
                        </motion.div>
                    )}
                </AnimatePresence>
            </div>
        </header>
    );
};

export default Header;
