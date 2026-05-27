import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import mkcert from 'vite-plugin-mkcert';

export default defineConfig({
  base: '/gnattHelper/',
  plugins: [
    react(),
    mkcert() // Automatically handles HTTPS certificates
  ],
  server: {
    port: 3000,
    https: true, // critical for Office
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        main: 'index.html', // We will rename taskpane.html to index.html
      },
    },
  },
});
