import { defineConfig } from "vite";
import react from "@vitejs/plugin-react";
import darkThemePlugin from "./vite-dark-theme.js";

export default defineConfig({
  plugins: [react(), darkThemePlugin()],
  server: {
    host: "127.0.0.1",
    port: 5174,
    proxy: {
      "/api": "http://127.0.0.1:4174",
    },
  },
});
