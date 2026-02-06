import React from 'react';
import Header from './Header';
import Footer from './Footer';

const Layout = ({ children, onBookNow }) => {
    return (
        <div className="min-h-screen flex flex-col bg-stone-50 font-sans text-stone-800">
            <Header onBookNow={onBookNow} />
            <main className="flex-grow">
                {children}
            </main>
            <Footer />
        </div>
    );
};

export default Layout;
