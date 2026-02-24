import { defineConfig } from 'vite';
import fs from 'fs';
import path from 'path';

export default defineConfig({
  server: {
    port: 3000,
    https: fs.existsSync('./localhost.pem') ? {
      key: fs.readFileSync('./localhost-key.pem'),
      cert: fs.readFileSync('./localhost.pem'),
    } : undefined,
  },
  build: {
    outDir: 'dist',
    rollupOptions: {
      input: {
        taskpane: path.resolve(__dirname, 'taskpane.html'),
      },
    },
  },
});
