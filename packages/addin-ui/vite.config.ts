import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

export default defineConfig({
  plugins: [react(), tailwindcss()],
  server: {
    host: '0.0.0.0',
    port: 5173,
    proxy: {
      '/api': 'http://127.0.0.1:18080',
      '/office': 'http://127.0.0.1:18080',
      '/wps': 'http://127.0.0.1:18080'
    }
  },
  build: {
    outDir: 'dist',
    emptyOutDir: true
  }
});
