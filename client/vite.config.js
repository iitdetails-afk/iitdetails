import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import path from 'path';
import fs from 'fs';

const backendPortFile = path.resolve(__dirname, '..', '.backend-port');

function getBackendPort() {
  try {
    if (fs.existsSync(backendPortFile)) {
      const port = Number(fs.readFileSync(backendPortFile, 'utf8').trim());
      if (!Number.isNaN(port) && port > 0) {
        return port;
      }
    }
  } catch {
    // fall back to the default port if the file is unavailable
  }

  return Number(process.env.PORT || 5000);
}

export default defineConfig({
  plugins: [react()],
  root: path.resolve(__dirname, '..'),
  publicDir: path.resolve(__dirname, '..'),
  server: {
    host: '127.0.0.1',
    port: 5174,
    strictPort: false,
    proxy: {
      '/api': {
        target: `http://127.0.0.1:${getBackendPort()}`,
        changeOrigin: true
      }
    }
  }
});
