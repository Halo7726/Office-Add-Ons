import { defineConfig } from "vite";
import basicSsl from "@vitejs/plugin-basic-ssl";

export default defineConfig({
  plugins: [basicSsl()],
  server: {
    host: "localhost",
    port: 3000,
    https: true,
  },
  preview: {
    host: "localhost",
    port: 3000,
    https: true,
  },
});
