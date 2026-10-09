/**
 * Packs the library with `pnpm pack`, installs the tarball into a temporary project and verifies
 * the published entry points:
 * - `.`: ESM import and CJS require work, every runtime export of src/index.ts is exported, and
 *   none of the experimental pipeline functions leak into it;
 * - `./internals`: every runtime export of src/internals.ts resolves from ESM and CJS;
 * - `./augment`: a no-op runtime module (ESM and CJS) whose declarations add
 *   `exporting.editableExcel` to `Highcharts.Options` (typechecked against a real Highcharts);
 * - declarations exist and compile in a consumer, Highcharts is not bundled, and publint +
 *   are-the-types-wrong report no errors on the tarball.
 *
 * Builds first unless `--no-build` is passed AND dist/ already exists (CI builds in an earlier step).
 */
import { execSync } from 'node:child_process';
import { existsSync, mkdtempSync, readFileSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';

const root = resolve(import.meta.dirname ?? process.cwd(), '..');
const run = (cmd: string, cwd = root) => execSync(cmd, { cwd, stdio: 'pipe', encoding: 'utf8' });
const bin = (name: string): string => `"${join(root, 'node_modules', '.bin', name)}"`;

/** Fails with the tool's own output (stdout + stderr) instead of execSync's generic message. */
function runTool(label: string, cmd: string): string {
  try {
    return run(cmd);
  } catch (error) {
    const e = error as { stdout?: string; stderr?: string };
    throw new Error(`${label} failed:\n${e.stdout ?? ''}${e.stderr ?? ''}`);
  }
}

// The runtime exports of the source entry points are the contract: the package must expose them all.
const EXPECTED_EXPORTS = Object.keys(await import('../src/index')).sort();
if (EXPECTED_EXPORTS.length === 0) throw new Error('src/index.ts has no runtime exports');
const EXPECTED_INTERNALS = Object.keys(await import('../src/internals')).sort();
/** Experimental runtime exports: only `highcharts-editable-excel/internals` may expose them. */
const EXPERIMENTAL_ONLY = [
  'applyThemeOverrides',
  'buildCompatibilityReport',
  'createDefaultExcelWriter',
  'createDiagnostic',
  'DiagnosticCollector',
  'extractChartModel',
  'extractChartModelFromOptions',
  'resolveChartType',
  'translateChartModel',
];
const missingInternals = EXPERIMENTAL_ONLY.filter((k) => !EXPECTED_INTERNALS.includes(k));
if (missingInternals.length) throw new Error(`src/internals.ts does not export: ${missingInternals.join(', ')}`);
const leaked = EXPERIMENTAL_ONLY.filter((k) => EXPECTED_EXPORTS.includes(k));
if (leaked.length) throw new Error(`src/index.ts must not export experimental names: ${leaked.join(', ')}`);
// Same version as the devDependency, so the augmentation is checked against the Highcharts we test.
const HIGHCHARTS_VERSION = (
  JSON.parse(readFileSync(join(root, 'package.json'), 'utf8')) as { devDependencies: Record<string, string> }
).devDependencies.highcharts;

const noBuild = process.argv.includes('--no-build');
const ENTRIES = ['index', 'internals', 'augment', 'chartjs'];
const distReady = ENTRIES.every(
  (e) => existsSync(join(root, 'dist', `${e}.js`)) && existsSync(join(root, 'dist', `${e}.d.ts`)),
);

const work = mkdtempSync(join(tmpdir(), 'hc-xlsx-pack-'));
try {
  if (noBuild && distReady) console.log('using the existing dist/ (--no-build)');
  else run('pnpm build');
  const packOut = run(`pnpm pack --pack-destination "${work}"`);
  const tarball = readdirSync(work).find((f) => f.endsWith('.tgz'));
  if (!tarball) throw new Error(`pnpm pack produced no tarball: ${packOut}`);
  const tgz = join(work, tarball);

  // Package-shape linters on the exact tarball that would be published.
  console.log(runTool('publint', `${bin('publint')} run "${tgz}" --level warning`).trim() || 'publint ok');
  // Every resolution mode (node10 included) for the main entry; subpath exports need node16+/bundler.
  console.log(runTool('attw', `${bin('attw')} "${tgz}" --format ascii --entrypoints .`).trim());
  console.log(runTool('attw', `${bin('attw')} "${tgz}" --format ascii --profile node16`).trim());

  const app = join(work, 'app');
  writeFileSync(join(work, 'package.json'), '{}');
  run(`mkdir -p "${app}"`);
  writeFileSync(join(app, 'package.json'), JSON.stringify({ name: 'pack-check-app', private: true, type: 'module' }));
  run(`npm install --no-audit --no-fund --silent "${tgz}" "highcharts@${HIGHCHARTS_VERSION}"`, app);

  const pkgDir = join(app, 'node_modules', 'highcharts-editable-excel');
  const pkg = JSON.parse(readFileSync(join(pkgDir, 'package.json'), 'utf8')) as Record<string, unknown>;
  const dist = readdirSync(join(pkgDir, 'dist'));
  const missingDist = ENTRIES.flatMap((e) => [`${e}.js`, `${e}.cjs`, `${e}.d.ts`, `${e}.d.cts`]).filter(
    (f) => !dist.includes(f),
  );
  if (missingDist.length) throw new Error(`dist incomplete (missing ${missingDist.join(', ')}): ${dist.join(', ')}`);
  const exportsMap = pkg.exports as Record<string, unknown>;
  for (const sub of ['./internals', './augment', './chartjs'])
    if (!exportsMap?.[sub]) throw new Error(`package.json exports has no "${sub}"`);
  if (!(pkg.peerDependencies as Record<string, string>)?.highcharts)
    throw new Error('highcharts must be a peer dependency');
  if ((pkg.dependencies as Record<string, string> | undefined)?.highcharts)
    throw new Error('highcharts must not be a dependency');

  // Every emitted script (entries and shared chunks, ESM and CJS) is checked.
  for (const file of dist.filter((f) => /\.c?js$/.test(f))) {
    const bundle = readFileSync(join(pkgDir, 'dist', file), 'utf8');
    if (/Highcharts JS v\d/.test(bundle) || bundle.length > 600_000)
      throw new Error(`${file} appears to include Highcharts`);
    if (/from\s+["']node:/.test(bundle) || /require\(["']fs["']\)/.test(bundle))
      throw new Error(`${file} imports Node-only modules`);
  }

  const exportCheck = (lib: string, expected: string[], forbidden: string[]) => `{
  const expected = ${JSON.stringify(expected)};
  const forbidden = ${JSON.stringify(forbidden)};
  const missing = expected.filter((k) => !(k in ${lib}));
  if (missing.length) { console.error(${JSON.stringify(lib)} + ': exports missing from the package: ' + missing.join(', ')); process.exit(1); }
  const leaked = forbidden.filter((k) => k in ${lib});
  if (leaked.length) { console.error(${JSON.stringify(lib)} + ': must not export ' + leaked.join(', ')); process.exit(1); }
  const extra = Object.keys(${lib}).filter((k) => k !== 'default' && !expected.includes(k));
  console.log(${JSON.stringify(lib)} + ' exports ok: ' + expected.length + (extra.length ? ' (package-only: ' + extra.join(', ') + ')' : ''));
}`;
  const checks = [
    exportCheck('lib', EXPECTED_EXPORTS, EXPERIMENTAL_ONLY),
    exportCheck('internals', EXPECTED_INTERNALS, []),
    `if (typeof internals.translateChartModel !== 'function') { console.error('translateChartModel is not a function'); process.exit(1); }`,
  ].join('\n');
  writeFileSync(
    join(app, 'check.mjs'),
    `import * as lib from 'highcharts-editable-excel';
import * as internals from 'highcharts-editable-excel/internals';
import 'highcharts-editable-excel/augment';
${checks}
console.log('esm ok (., ./internals, ./augment)');`,
  );
  writeFileSync(
    join(app, 'check.cjs'),
    `const lib = require('highcharts-editable-excel');
const internals = require('highcharts-editable-excel/internals');
require('highcharts-editable-excel/augment');
${checks}
console.log('cjs ok (., ./internals, ./augment)');`,
  );
  console.log(run('node check.mjs', app).trim());
  console.log(run('node check.cjs', app).trim());

  // Declarations must compile in a consumer project, including the subpaths and the augmentation.
  writeFileSync(
    join(app, 'types-check.ts'),
    `import { exportHighchartsToXlsx, type ChartModel, type ExcelWriter, type ExportOptions } from 'highcharts-editable-excel';
import { DiagnosticCollector, translateChartModel, type WorkbookSpec } from 'highcharts-editable-excel/internals';
const o: ExportOptions = { filename: 'x.xlsx', fidelity: 'best-effort', dataMode: 'rendered' };
export const f = (c: unknown) => exportHighchartsToXlsx(c, o);
export const w: ExcelWriter = { name: 'noop', write: async (_wb: WorkbookSpec) => new Uint8Array() };
export const t = (m: ChartModel) => translateChartModel(m, { chartSheetName: 'C', dataSheetName: 'D', includeSourceData: true, fidelity: 'best-effort', diagnostics: new DiagnosticCollector() });`,
  );
  writeFileSync(
    join(app, 'augment-check.ts'),
    `import 'highcharts-editable-excel/augment';
import Highcharts from 'highcharts';
const o: Highcharts.Options = { exporting: { editableExcel: { enabled: false } } };
export default o;`,
  );
  // Negative control: without the augmentation the same option must be rejected.
  writeFileSync(
    join(app, 'augment-negative.ts'),
    `import Highcharts from 'highcharts';
const o: Highcharts.Options = { exporting: { editableExcel: { enabled: false } } };
export default o;`,
  );
  const tsconfig = (files: string[], moduleResolution: string, module: string) =>
    JSON.stringify({
      compilerOptions: {
        strict: true,
        module,
        moduleResolution,
        target: 'ES2022',
        noEmit: true,
        skipLibCheck: false,
        esModuleInterop: true,
        types: [],
      },
      files,
    });
  const tsc = (project: string) =>
    run(`node "${join(root, 'node_modules', 'typescript', 'bin', 'tsc')}" -p ${project}`, app);
  for (const [name, resolution, mod] of [
    ['tsconfig.bundler.json', 'Bundler', 'ESNext'],
    ['tsconfig.node16.json', 'Node16', 'Node16'],
  ] as const) {
    writeFileSync(join(app, name), tsconfig(['types-check.ts', 'augment-check.ts'], resolution, mod));
    try {
      tsc(name);
    } catch (error) {
      const e = error as { stdout?: string; stderr?: string };
      throw new Error(`consumer typecheck (${resolution}) failed:\n${e.stdout ?? ''}${e.stderr ?? ''}`);
    }
  }
  writeFileSync(join(app, 'tsconfig.negative.json'), tsconfig(['augment-negative.ts'], 'Bundler', 'ESNext'));
  let negativeFailed = false;
  try {
    tsc('tsconfig.negative.json');
  } catch (error) {
    negativeFailed = /editableExcel/.test((error as { stdout?: string }).stdout ?? '');
  }
  if (!negativeFailed) throw new Error('negative control: editableExcel typechecked without the augmentation');
  console.log('declarations ok (., ./internals, ./augment augments Highcharts.Options)');
  console.log(`pack-check passed (${tarball})`);
} finally {
  rmSync(work, { recursive: true, force: true });
}
