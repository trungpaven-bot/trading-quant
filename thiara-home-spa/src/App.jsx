import React, { useState } from 'react';
import Layout from './components/layout/Layout';
import Hero from './components/home/Hero';
import Features from './components/home/Features';
import Services from './components/home/Services';
import Testimonials from './components/home/Testimonials';
import BookingModal from './components/booking/BookingModal';
import { AnimatePresence } from 'framer-motion';

function App() {
  const [isBookingOpen, setIsBookingOpen] = useState(false);

  const handleBookNow = () => {
    setIsBookingOpen(true);
  };

  return (
    <Layout onBookNow={handleBookNow}>
      <Hero onBookNow={handleBookNow} />
      <Features />
      <Services onBookNow={handleBookNow} />
      <Testimonials />

      <AnimatePresence>
        {isBookingOpen && (
          <BookingModal
            isOpen={isBookingOpen}
            onClose={() => setIsBookingOpen(false)}
          />
        )}
      </AnimatePresence>
    </Layout>
  );
}

export default App;
