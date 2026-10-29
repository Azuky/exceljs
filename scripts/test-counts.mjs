// Runs the vitest projects with the json reporter and records test counts.
// Usage: node scripts/test-counts.mjs --out <file.json>
import { spawnSync } from 'node:child_process';
import { mkdirSync, mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const SUITES = ['unit', 'integration', 'end-to-end'];

const outIdx = process.argv.indexOf('--out');
if (outIdx === -1 || !process.argv[outIdx + 1]) {
  console.error('Usage: node scripts/test-counts.mjs --out <file.json>');
  process.exit(2);
}
const outFile = resolve(process.argv[outIdx + 1]);

const env = { ...process.env };
delete env.EXCEL_BUILD;

const stderrTail = (res) => (res.stderr || '').trim().split('\n').slice(-20).join('\n');

const suites = {};
const tmpDir = mkdtempSync(join(tmpdir(), 'test-counts-'));
for (const name of SUITES) {
  const reportFile = join(tmpDir, `${name}.json`);
  const res = spawnSync(
    join(root, 'node_modules/.bin/vitest'),
    ['run', '--project', name, '--reporter=json', `--outputFile=${reportFile}`],
    {
      cwd: root,
      env,
      encoding: 'utf8',
      stdio: ['ignore', 'pipe', 'pipe'],
      maxBuffer: 1 << 29,
    },
  );
  if (res.error) {
    throw new Error(`suite ${name}: failed to run vitest: ${res.error.message}\n${stderrTail(res)}`);
  }
  let report;
  try {
    report = JSON.parse(readFileSync(reportFile, 'utf8'));
  } catch (err) {
    throw new Error(`suite ${name}: no parsable vitest JSON (exit ${res.status}): ${err.message}\n${stderrTail(res)}`);
  }
  const failingTitles = report.testResults.flatMap((file) =>
    file.assertionResults.filter((a) => a.status === 'failed').map((a) => a.fullName),
  );
  const tests = report.numTotalTests;
  const passes = report.numPassedTests;
  const failures = report.numFailedTests;
  const pending = report.numPendingTests + (report.numTodoTests || 0);
  suites[name] = { tests, passes, failures, pending, failingTitles };
  console.log(`${name}: tests=${tests} passes=${passes} failures=${failures} pending=${pending}`);
}
rmSync(tmpDir, { recursive: true, force: true });

mkdirSync(dirname(outFile), { recursive: true });
writeFileSync(outFile, `${JSON.stringify({ node: process.version, suites }, null, 2)}\n`);
