import type { Config } from "tailwindcss";

export default {
  content: [
    "./src/pages/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/components/**/*.{js,ts,jsx,tsx,mdx}",
    "./src/app/**/*.{js,ts,jsx,tsx,mdx}",
  ],
  theme: {
    extend: {
      colors: {
        background: "var(--background)",
        foreground: "var(--foreground)",
        cream: "#FAF7F4",
        ink: { DEFAULT: "#4A3F3A", muted: "#8A7A74", dark: "#2C2420" },
        gold: { DEFAULT: "#C9A84C", light: "#E8C97A", dark: "#A07830" },
        rose: { DEFAULT: "#F2D9DF", mid: "#E8B8C4" },
        lilac: { DEFAULT: "#E8D9F0", mid: "#C9A8D8" },
        footer: "#1C1410",
      },
      fontFamily: {
        serif: ["var(--font-cormorant)", "Georgia", "serif"],
        sans: ["var(--font-raleway)", "system-ui", "sans-serif"],
      },
    },
  },
  plugins: [],
} satisfies Config;
