import { fileURLToPath } from 'node:url';
import { defineConfig } from 'vite';

const here = fileURLToPath(new URL('.', import.meta.url));
const fromHere = (p: string): string => fileURLToPath(new URL(p, import.meta.url));

export default defineConfig({
  root: here,
  resolve: {
    // The demo runs against the library SOURCE, not dist/ — edits in src/ hot-reload here.
    alias: { 'highcharts-editable-excel': fromHere('../src/index.ts') },
  },
  server: { host: '127.0.0.1', port: 4173, strictPort: true },
  preview: { host: '127.0.0.1', port: 4173, strictPort: true },
  // Highcharts alone is ~500 kB minified; a single chunk is fine for a demo.
  build: { outDir: fromHere('../dist-demo'), emptyOutDir: true, chunkSizeWarningLimit: 1500 },
  // Highcharts ships UMD/CJS bundles; pre-bundle them so the side-effect module imports work in dev.
  optimizeDeps: { include: ['highcharts', 'highcharts/modules/exporting', 'highcharts/highcharts-more'] },
});
