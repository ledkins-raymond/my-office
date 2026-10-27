import { defineConfig } from 'vite';

export default defineConfig({
  root: 'src/word',
  base: './',
  build: {
    outDir: '../../dist/word',
    emptyOutDir: true,
    chunkSizeWarningLimit: 2000,
  },
});
