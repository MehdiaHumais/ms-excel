import type { Presentation } from "./types";

export async function fetchPresentations(): Promise<Presentation[]> {
  // Mocked data – replace with real API call when available
  return new Promise((resolve) => {
    setTimeout(() => {
      resolve([
        {
          id: "1",
          title: "Quarterly Business Review",
          description: "A concise overview of Q3 performance.",
          slides: 12,
          thumbnailUrl:
            "https://images.unsplash.com/photo-1581091012184-7a2e6c6c6e2c?au,
        {
          id: "3",
          title: "Investor Pitch Deck",
          description: "Key metrics and growth strategy for investors.",
          slides: 15,
          thumbnailUrl:
            "https://images.unsplash.com/photo-1518770660439-4636190af475?au*/
module.exports = {
  content: [
    "./app/**/*.{js,ts,jsx,tsx}",
    "./components/**/*.{js,ts,jsx,tsx}",
    "./lib/**/*.{js,ts,jsx,tsx}",
  ],
  theme: {
    extend: {
      colors: {
        primary: "#3b82f6", // blue-500
        secondary: "#10b981", // emerald-500
        accent: "#f59e0b", // amber-500
        "accent-dark": "#b45309",
        background: "#111827", // gray-900
        surface: "#1f2937", // gray-800
        text: "#f3f4f6", // gray-100
      },
      fontFamily: {
        sans: [
          "system-ui",
          "-apple-system",
          "BlinkMacSystemFont",
          "Segoe UI",
          "Roboto",
          "Helvetica Neue",
          "Arial",
          "Noto Sans",
          "sans-serif",
        ],
      },
    },
  },
  plugins: [],
};