/** @type {import('tailwindcss').Config} */
export default {
  content: [
    "./index.html",
    "./src/**/*.{js,ts,jsx,tsx}",
    "./*.{js,ts,jsx,tsx}"
  ],
  theme: {
    extend: {
      colors: {
        brand: {
          primary: 'var(--brand-primary)',
          accent: 'var(--brand-accent)',
          ink: 'var(--brand-ink)',
          'ink-soft': 'var(--brand-ink-soft)',
          gold: 'var(--brand-gold)',
        },
        surface: {
          app: 'var(--bg-app-logged-in)',
          offwhite: 'var(--bg-offwhite)',
          DEFAULT: 'var(--bg-surface)',
          cream: 'var(--bg-cream)',
          'ice-blue': 'var(--bg-ice-blue)',
        },
        borderToken: 'var(--border)',
        dangerToken: 'var(--danger)',
      }
    },
  },
  plugins: [],
}

