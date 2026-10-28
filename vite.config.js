import { defineConfig } from 'vite';
import { resolve } from 'node:path';

// One build for both programs: dist/word/index.html and dist/excel/index.html.
export default defineConfig({
  root: 'src',
  base: './',
  build: {
    outDir: '../dist',
    emptyOutDir: true,
    chunkSizeWarningLimit: 3000,
    rollupOptions: {
      input: {
        word: resolve(import.meta.dirname, 'src/word/index.html'),
        excel: resolve(import.meta.dirname, 'src/excel/index.html'),
      },
    },
  },
});
