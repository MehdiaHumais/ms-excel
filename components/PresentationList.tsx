"use client";

import { useEffect, useState } from "react";
import type { Presentation } from "@/lib/types";
import { fetchPresentations } from "@/lib/api";
import PresentationCard from "./PresentationCard";

export default function PresentationList() {
  const [presentations, setPresentations] = useState<Presentation[] | null>(null);
  const [error, setError] = useState<string | null>(null);

  useEffect(() => {
    fetchPresentations()
      .then(setPresentations)
      .catch((e) => setError(e.message));
  }, []);

  if (error) {
    return (
      <div className="p-6 bg-red-900/50 text-red-100 rounded-2xl text-center">
        <p className="font-medium">Failed to load presentations.</p>
        <p>{error}</p>
      </div>
    );
  }

  if (presentations === null) {
    // Loading skeleton
    return (
      <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
        {[...Array(3)].map((_, i) => (
          <div
            key={i}
            className="bg-gray-800 rounded-2xl p-4 animate-pulse"
          >
            <div className="h-48 bg-gray-700 rounded mb-4"></div>
            <div className="h-5 bg-gray-700 rounded w-3/4 mb-2"></div>
            <div className="h-4 bg-gray-700 rounded w-1/2"></div>
          </div>
        ))}
      </div>
    );
  }

  if (presentations.length === 0) {
    return (
      <div className="p-6 bg-gray-800/50 text-gray-300 rounded-2xl text-center">
        <p className="font-medium">No presentations found.</p>
      </div>
    );
  }

  return (
    <div className="grid gap-6 sm:grid-cols-2 lg:grid-cols-3">
      {presentations.map((p) => (
        <PresentationCard key={p.id} presentation={p} />
      ))}
    </div>
  );
}