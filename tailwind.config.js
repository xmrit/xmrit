/** @type {import('tailwindcss').Config} */
module.exports = {
  content: ["./css/**/*.css", "./js/**/*.{js,ts}", "./*.html"],
  theme: {
    extend: {
      // Add inter var font
      fontFamily: {
        sans: ["Inter var", "sans-serif"],
      },
    },
  },
  plugins: [require("@tailwindcss/typography"), require("@tailwindcss/forms")],
};
