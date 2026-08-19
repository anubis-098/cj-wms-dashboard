/** @type {import('tailwindcss').Config} */
export default {
  darkMode: "class",
  content: ["./index.html", "./src/**/*.{js,ts,jsx,tsx}"],
  theme: {
    extend: {
      colors: {
        "cj-red": "#e42f44",
        "cj-blue": "#0080c6",
        "cj-navy": "#122033",
        "cj-yellow": "#ec8922",
        "screen-bg": "#f2f5f7",
      },
      boxShadow: {
        panel: "0 12px 30px rgba(18, 32, 51, 0.08)",
      },
    },
  },
  plugins: [],
};
