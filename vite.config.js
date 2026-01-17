import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";

export default defineConfig({
  plugins: [react()],
  base: "/Paraglider-trim-tuning/",
  build: {
    target: "es2018"
  }
});
