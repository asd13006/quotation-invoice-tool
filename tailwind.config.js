/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./templates/**/*.html"],
  darkMode: 'class',
  theme: {
    extend: {
      fontFamily: {
        sans: ["Microsoft JhengHei", "微軟正黑體", "system-ui", "sans-serif"],
        mono: ["Cascadia Code", "Fira Code", "Consolas", "monospace"],
      },
    },
  },
  plugins: [],
}
