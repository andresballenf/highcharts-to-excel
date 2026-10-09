/**
 * Packs the library with `pnpm pack`, installs the tarball into a temporary project and verifies
 * the published entry point: ESM import works, declarations exist, every runtime export of
 * src/index.ts is exported by the package, Highcharts is not bundled, and publint +
 * are-the-types-wrong report no errors on the tarball.
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

// The runtime exports of the source entry point are the contract: the package must expose them all.
const EXPECTED_EXPORTS = Object.keys(await import('../src/index')).sort();
if (EXPECTED_EXPORTS.length === 0) throw new Error('src/index.ts has no runtime exports');

const noBuild = process.argv.includes('--no-build');
const distReady = existsSync(join(root, 'dist', 'index.js')) && existsSync(join(root, 'dist', 'index.d.ts'));

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
  console.log(runTool('attw', `${bin('attw')} "${tgz}" --format ascii`).trim());

  const app = join(work, 'app');
  writeFileSync(join(work, 'package.json'), '{}');
  run(`mkdir -p "${app}"`);
  writeFileSync(join(app, 'package.json'), JSON.stringify({ name: 'pack-check-app', private: true, type: 'module' }));
  run(`npm install --no-audit --no-fund --silent "${tgz}"`, app);

  const pkgDir = join(app, 'node_modules', 'highcharts-editable-excel');
  const pkg = JSON.parse(readFileSync(join(pkgDir, 'package.json'), 'utf8')) as Record<string, unknown>;
  const dist = readdirSync(join(pkgDir, 'dist'));
  if (!dist.includes('index.js') || !dist.includes('index.d.ts'))
    throw new Error(`dist incomplete: ${dist.join(', ')}`);
  if (!(pkg.peerDependencies as Record<string, string>)?.highcharts)
    throw new Error('highcharts must be a peer dependency');
  if ((pkg.dependencies as Record<string, string> | undefined)?.highcharts)
    throw new Error('highcharts must not be a dependency');

  const bundle = readFileSync(join(pkgDir, 'dist', 'index.js'), 'utf8');
  if (/Highcharts JS v\d/.test(bundle) || bundle.length > 600_000)
    throw new Error('bundle appears to include Highcharts');
  if (/from\s+["']node:/.test(bundle) || /require\(["']fs["']\)/.test(bundle))
    throw new Error('bundle imports Node-only modules');

  writeFileSync(
    join(app, 'check.mjs'),
    `import * as lib from 'highcharts-editable-excel';
const expected = ${JSON.stringify(EXPECTED_EXPORTS)};
const missing = expected.filter((k) => !(k in lib));
if (missing.length) { console.error('exports in src/index.ts missing from the package: ' + missing.join(', ')); process.exit(1); }
const extra = Object.keys(lib).filter((k) => !expected.includes(k));
console.log('exports ok: ' + expected.length + (extra.length ? ' (package-only: ' + extra.join(', ') + ')' : ''));`,
  );
  console.log(run('node check.mjs', app).trim());

  // Declarations must compile in a consumer project.
  writeFileSync(
    join(app, 'types-check.ts'),
    `import { exportHighchartsToXlsx, type ExportOptions } from 'highcharts-editable-excel';
const o: ExportOptions = { filename: 'x.xlsx', fidelity: 'best-effort', dataMode: 'rendered' };
export const f = (c: unknown) => exportHighchartsToXlsx(c, o);`,
  );
  writeFileSync(
    join(app, 'tsconfig.json'),
    JSON.stringify({
      compilerOptions: {
        strict: true,
        module: 'ESNext',
        moduleResolution: 'Bundler',
        target: 'ES2022',
        noEmit: true,
        skipLibCheck: false,
        types: [],
      },
      files: ['types-check.ts'],
    }),
  );
  run(`node "${join(root, 'node_modules', 'typescript', 'bin', 'tsc')}" -p tsconfig.json`, app);
  console.log('declarations ok');
  console.log(`pack-check passed (${tarball})`);
} finally {
  rmSync(work, { recursive: true, force: true });
}
