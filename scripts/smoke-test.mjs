// Packs the package, installs the tarball with npm in a throwaway app and
// runs the CJS and ESM consumers against it.
import { spawnSync } from 'node:child_process';
import fs from 'node:fs';
import os from 'node:os';
import path from 'node:path';
import { fileURLToPath } from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const smokeDir = path.join(root, 'scripts', 'smoke');
const win = process.platform === 'win32';

// With shell: true Node joins args unquoted, so on Windows quote any arg that
// contains whitespace or shell metacharacters (e.g. C:\Users\John Doe\...).
// Windows paths cannot contain double quotes, so wrapping in "..." is safe.
const quote = (arg) => (win && /[\s&()^%!;,=]/.test(arg) ? `"${arg}"` : arg);

function run(cmd, args, cwd) {
  const res = spawnSync(quote(cmd), args.map(quote), { cwd, stdio: 'inherit', shell: win });
  if (res.status !== 0) {
    throw new Error(`Command failed (${res.status}): ${cmd} ${args.join(' ')}`);
  }
}

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), 'exceljs-smoke-'));
let failed = false;
try {
  // pack always builds first (via prepack), so dist/ is fresh
  run('pnpm', ['pack', '--pack-destination', tmp], root);
  const tarball = fs.readdirSync(tmp).find((f) => f.endsWith('.tgz'));
  if (!tarball) throw new Error('pnpm pack produced no tarball');

  const app = path.join(tmp, 'app');
  fs.mkdirSync(app);
  fs.writeFileSync(
    path.join(app, 'package.json'),
    JSON.stringify({ name: 'exceljs-smoke-app', version: '1.0.0', private: true }),
  );
  run('npm', ['install', '--no-audit', '--no-fund', path.join(tmp, tarball)], app);

  fs.copyFileSync(path.join(smokeDir, 'consumer.cjs'), path.join(app, 'consumer.cjs'));
  fs.copyFileSync(path.join(smokeDir, 'consumer.mjs'), path.join(app, 'consumer.mjs'));
  run(process.execPath, ['consumer.cjs'], app);
  run(process.execPath, ['consumer.mjs'], app);
  console.log('smoke test OK');
} catch (err) {
  failed = true;
  console.error(err.message);
} finally {
  fs.rmSync(tmp, { recursive: true, force: true });
}
process.exit(failed ? 1 : 0);
