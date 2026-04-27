import type { Config } from "tailwindcss";

const config: Config = {
  darkMode: ["class"],
  content: ["./app/**/*.{ts,tsx}", "./components/**/*.{ts,tsx}"],
  theme: {
    extend: {
      colors: {
        cream: {
          50: "#fffdf8",
          100: "#faf6ef",
          200: "#f0e8d8",
          300: "#e5d9c4",
        },
        terracotta: {
          DEFAULT: "#c45c3e",
          dark: "#a34a30",
          light: "#e0785c",
        },
        ink: {
          DEFAULT: "#3d2f2f",
          muted: "#6b5e5c",
        },
      },
      borderRadius: {
        xl: "0.75rem",
        "2xl": "1rem",
      },
      fontFamily: {
        sans: ["var(--font-inter)", "Hiragino Sans", "Yu Gothic UI", "system-ui", "sans-serif"],
      },
    },
  },
  plugins: [],
};

export default config;
