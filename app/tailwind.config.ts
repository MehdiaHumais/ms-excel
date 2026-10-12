import type { Config } from 'tailwindcss'

const config: Config = {
  content: [
    './app/**/*.{js,ts,jsx,tsx}',
    './components/**/*.{js,ts,jsx,tsx}',
  ],
  theme: {
    extend: {
      colors: {
        primary: '#0ea5e9', // sky-500
        secondary: '#6366f1', // indigo-500
        accent: '#f59e0b', // amber-500
        background: '#f3f4f6', // gray-100
        surface: '#ffffff',
        muted: '#9ca3af', // gray-400
      },
      fontFamily: {
        sans: ['system-ui', 'sans-serif', 'Apple Color Emoji', 'Segoe UI Emoji'],
      },
      borderRadius: {
        xl: '1rem',
        '2xl': '1.5rem',
      },
      boxShadow: {
        card: '0 4px 12px rgba(0,0,0,0.05)',
        glow: '0 0 12px rgba(14,165,233,0.3)',
      },
      animation: {
        fadeIn: 'fadeIn 0.5s ease-out forwards',
      },
      keyframes: {
        fadeIn: {
          '0%': { opacity: 0, transform: 'translateY(10px)' },
          '100%': { opacity: 1, transform: 'translateY(0)' },
        },
      },
    },
  },
  plugins: [],
}

export default config