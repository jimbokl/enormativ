import { defineConfig } from 'vite'
import { fileURLToPath } from 'node:url'

export default defineConfig({
  base: './',
  envDir: '..',
  build: {
    outDir: 'dist',
    emptyOutDir: true,
    rollupOptions: {
      input: {
        portal: fileURLToPath(new URL('./index.html', import.meta.url)),
        check: fileURLToPath(new URL('./tools/check/index.html', import.meta.url)),
        compare: fileURLToPath(new URL('./tools/compare/index.html', import.meta.url)),
        sourcesExample: fileURLToPath(new URL('./examples/source-inventory/index.html', import.meta.url)),
      },
    },
  },
})
