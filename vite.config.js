import { defineConfig } from 'vite'
import fs from 'fs'
import path from 'path'
import os from 'os'

const certDir = path.join(os.homedir(), '.office-addin-dev-certs')

export default defineConfig({
  plugins: [],
  server: {
    https: {
      cert: fs.readFileSync(path.join(certDir, 'localhost.crt')),
      key: fs.readFileSync(path.join(certDir, 'localhost.key')),
    },
    port: 3000
  },
  build: {
    rollupOptions: {
      input: {
        main: path.resolve(__dirname, 'index.html'),
        authDialog: path.resolve(__dirname, 'auth-dialog.html'),
      },
    },
  },
})