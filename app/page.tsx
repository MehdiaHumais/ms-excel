"use client";

import { useEffect, useState } from "react";
import HeroSection from "@/components/HeroSection";
import FeatureCard from "@/components/FeatureCard";
import LoadingSkeleton from "@/components/LoadingSkeleton";
import ErrorAlert from "@/components/ErrorAlert";

type Feature = {
  title: string;
  description: string;
  icon: React.ReactNode;
};

export default function HomePage() {
  const [features, setFeatures] = useState<Feature[] | null>(null);
  const [loading, setLoading] = useState<boolean>(true);
  const [error, setError] = useState<string | null>(null);

  // Simulate data fetching
  useEffect(() => {
    const timer = setTimeout(() => {
      try {
        // Mock data
        const data: Feature[] = [
          {
            title: "Slide Templates",
            description: "Choose from a library of professionally designed slide templates.",
            icon: (
              <svg
                className="w-8 h-8 text-primary"
                fill="none"
                stroke="currentColor"
                strokeWidth={2}
                viewBox="0 0 24 24"
                aria-hidden="true"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M4 6h16M4 12h16M4 18h16"
                />
              </svg>
            ),
          },
          {
            title: "Real‑time Collaboration",
            description: "Edit presentations together with your team in real time.",
            icon: (
              <svg
                className="w-8 h-8 text-primary"
                fill="none"
                stroke="currentColor"
                strokeWidth={2}
                viewBox="0 0 24 24"
                aria-hidden="true"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M8 10h.M12 10h8M8 14h8M12 14h8"
                />
              </svg>
            ),
          },
          {
            title: "Export Options",
            description: "Export your slides as PDF, PPTX, or images with a single click.",
            icon: (
              <svg
                className="w-8 h-8 text-primary"
                fill="none"
                stroke="currentColor"
                strokeWidth={2}
                viewBox="0 0 24 24"
                aria-hidden="true"
              >
                <path
                  strokeLinecap="round"
                  strokeLinejoin="round"
                  d="M4 4h16v16H4V4z"
                />
              </svg>
            ),
          },
        ];
        setFeatures(data);
        setLoading(false);
      } catch (e) {
        setError("Failed to load features.");
        setLoading(false);
      }
    }, 1200);
    return () => clearTimeout(timer);
  }, []);

  return (
    <main className="w-full max-w-6xl mx-auto px-4 py-12 space-y-16">
      <HeroSection />
      {loading && (
        <div className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {[...Array(3)].map((_, i) => (
            <LoadingSkeleton key={i} />
          ))}
        </div>
      )}
      {error && <ErrorAlert message={error} />}
      {features && features.length === 0 && (
        <p className="text-center text-text-secondary">No features to display.</p>
      )}
      {features && features.length > 0 && (
        <section className="grid grid-cols-1 md:grid-cols-3 gap-6">
          {features.map((f, idx) => (
            <FeatureCard
              key={idx}
              title={f.title}
              description={f.description}
              icon={f.icon}
            />
          ))}
        </section>
      )}
    </main>
  );
}