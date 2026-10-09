import { defineConfig } from 'tsup';

/**
 * fflate is bundled; its MIT notice must ship with the code. fflate's sources carry no license
 * comment, so the notice is a preserved banner (`/*!`) plus THIRD_PARTY_LICENSES.md in the package.
 */
const LICENSE_BANNER =
  '/*! highcharts-editable-excel | MIT License | bundles fflate (https://github.com/101arrowz/fflate), ' +
  'MIT License, Copyright (c) 2026 Arjun Barrett | full license texts: THIRD_PARTY_LICENSES.md */';

export default defineConfig({
  // `index` is the stable entry, `internals` the experimental pipeline, `augment` the opt-in
  // TypeScript augmentation of `Highcharts.ExportingOptions` (an empty runtime module).
  entry: {
    index: 'src/index.ts',
    internals: 'src/internals.ts',
    augment: 'src/augment.ts',
    chartjs: 'src/chartjs/index.ts',
  },
  // ESM (*.js) for bundlers and Node `import`; CJS (*.cjs) for `require`.
  format: ['esm', 'cjs'],
  dts: true,
  banner: { js: LICENSE_BANNER },
  esbuildOptions(options) {
    // Keep `/*!` / `@license` comments (of any bundled dependency) at the end of the output.
    options.legalComments = 'eof';
  },
  sourcemap: true,
  clean: true,
  target: 'es2022',
  platform: 'browser',
  treeshake: true,
  external: ['highcharts', 'chart.js'],
});
