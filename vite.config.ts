import { defineConfig } from "vite";
import path from "node:path";
import electron from "vite-plugin-electron";
import react from "@vitejs/plugin-react";

export default defineConfig({
  resolve: {
    alias: {
      "@": path.resolve(__dirname, "src"),
    },
  },
  plugins: [
    react(),
    electron([
      {
        entry: "electron/main.ts",
      },
      {
        entry: path.resolve(__dirname, "electron/preload.ts"),
      },
    ]),
  ],
});
