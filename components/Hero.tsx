"use client";

import { useState, useEffect } from "react";

export default function Hero() {
  // Simple subtle animation on mount
  const [mounted, setMounted] = useState(false);
  useEffect(() => setMounted(true), []);

  return (
    <section className="text-center py-20 bg-gradient-to-r from-primary to-accent rounded-2xl shadow-xl transform transition-transform duration-700 ease-out">
      <h1
        className={`text-5xl font-extrabold mb-4 text-white transition-opacity duration-700 ${
          mounted ? "opacity-100" : "opacity-0"
        }`}
      >
        Create Stunning Presentations in Minutes
      </h1>
      <p className="text-lg text-gray-200 mb-8">
        PowerPoint Clone – modern UI, lightning‑fast collaboration, and cloud storage.
      </p>
      <a
        href="#"
        className="inline-block px-8 py-3 bg-white text-primary font-semibold rounded-full shadow-md hover:bg-gray-100 transition-colors"
      >
        Get Started
      </a>
    </section>
  );
}