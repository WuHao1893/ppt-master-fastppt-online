import path from 'node:path';
import { fileURLToPath, URL } from 'node:url';
import react from '@vitejs/plugin-react';
import { defineConfig } from 'vite';

const root = path.dirname(fileURLToPath(import.meta.url));

export default defineConfig({
  root,
  plugins: [react()],
  define: { 'import.meta.env.VITE_WS_PORT': JSON.stringify(process.env.WS_PORT || '8788') },
  resolve: { alias: { '@shared': new URL('../shared', import.meta.url).pathname } },
  server: {
    port: Number(process.env.WEB_PORT || 5173),
    proxy: {
      '/api': { target: `http://127.0.0.1:${process.env.PORT || 8787}`, changeOrigin: true },
    },
  },
  build: { outDir: path.resolve(root, 'dist'), emptyOutDir: true },
});
