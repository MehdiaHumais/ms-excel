/** @type {import('tailwindcss').Config} */
module.exports = {
  darkMode: 'class',
  content: ['./index.html', './src/**/*.{ts,tsx}'],
  theme: {
    extend: {
      colors: {
        brand: {
          50: '#eef4ff',
          100: '#dfe9ff',
          200: '#c0d4fe',
          300: '#97b8fd',
          400: '#6d94f9',
          500: '#4a72f1',
          600: '#3455dd',
          700: '#2a43bb',
          800: '#263b97',
          900: '#233477'
        },
        panel: 'rgb(var(--panel) / <alpha-value>)',
        cell: 'rgb(var(--cell) / <alpha-value>)'
      },
      fontSize: {
        '2xs': ['12px', '18px']
      }
    }
  },
  plugins: []
};