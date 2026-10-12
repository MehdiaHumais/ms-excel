import "./globals.css";
import type { Metadata } from "next";

export const metadata: Metadata = {
  title: "PowerPoint Clone",
  description: "A modern presentation builder built with Next.js and Tailwind CSS",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="en" className="h-full scroll-smooth">
      <body className="min-h-screen flex flex-col items-center justify-center bg-gradient-to-br from-primary to-secondary text-text-primary antialiased">
        {children}
      </body>
    </html>
  );
}