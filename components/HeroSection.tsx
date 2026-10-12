import Link from "next/link";

export default function HeroSection() {
  return (
    <section className="text-center space-y-6 animate-fadeIn">
      <h1 className="text-4xl md:text-5xl font-bold bg-clip-text text-transparent bg-gradient-to-r from-primary to-secondary">
        Build Stunning Presentations in Minutes
      </h1>
      <p className="text-lg text-text-secondary max-w-2xl mx-auto">
        PowerPoint Clone brings modern design, real‑time collaboration, and instant
        exporting to your fingertips. Create, share, and impress.
      </p>
      <Link
        href="#features"
        className="inline-block mt-4 px-6 py-3 bg-primary text-white font-medium rounded-2xl shadow-button hover:bg-primary/90 focus:outline-none focus:ring-2 focus:ring-accent transition-colors duration-300"
      >
        Explore Features
      </Link>
    </section>
  );
}