"use client";

import type { Presentation } from "@/lib/types";

export default function PresentationCard({
  presentation,
}: {
  presentation: Presentation;
}) {
  return (
    <div className="bg-gray-800 rounded-2xl overflow-hidden shadow-lg transform hover:scale-105 transition-transform duration-300">
      <img
        src={presentation.thumbnailUrl}
        alt={presentation.title}
        className="w-full h-48 object-cover"
      />
      <div className="p-4">
        <h3 className="text-xl font-semibold text-primary mb-2">
          {presentation.title}
        </h3>
        <p className="text-sm text-gray-300 mb-4">
          {presentation.description}
        </p>
        <div className="flex items-center justify-between">
          <span className="text-xs text-gray-400">
            {presentation.slides} slides
          </span>
          <a
            href="#"
            className="px-4 py-2 bg-accent text-white rounded-full hover:bg-accent-dark transition-colors"
          >
            Open
          </a>
        </div>
      </div>
    </div>
  );
}