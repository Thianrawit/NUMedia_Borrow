/** @type {import('tailwindcss').Config} */
module.exports = {
  content: [
    "./*.html",
    "./*.js"
  ],
  theme: {
    extend: {
      fontFamily: {
        sans: ['Inter', 'Noto Sans Thai', 'sans-serif'],
      },
      colors: {
        brand: {
          50: '#fff7ed',
          100: '#ffedd5',
          200: '#fed7aa',
          300: '#fdba74',
          400: '#fb923c',
          500: '#f97316',
          600: '#ea580c',
          700: '#c2410c',
          800: '#9a3412',
          900: '#7c2d12'
        }
      }
    },
  },
  plugins: [],
  safelist: [
    'border-red-300', 'bg-red-50', 'text-red-600', 'text-red-500',
    'border-amber-300', 'bg-amber-50', 'text-amber-600',
    'border-emerald-300', 'bg-emerald-50', 'text-emerald-600',
    'border-[#f97316]', 'bg-orange-50/70', 'scale-[1.01]'
  ]
}
