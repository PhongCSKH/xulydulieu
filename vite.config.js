import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import tailwindcss from '@tailwindcss/vite';

// https://vitejs.dev/config/
export default defineConfig({
  plugins: [
    react(),
    tailwindcss()
  ],
  base: '/', // Phục vụ tại tên miền gốc https://xulydulieu.site
  server: {
    port: 3000,
    open: true
  }
});
