// Runs the mocha suites with the json reporter and records test counts.
// Usage: node scripts/test-counts.mjs --out <file.json>
import {spawnSync} from 'node:child_process';
import {mkdirSync, writeFileSync} from 'node:fs';
import {dirname, join, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '..');

const SUITES = {
  unit: ['--require', 'spec/config/setup', '--require', 'spec/config/setup-unit', 'spec/unit', '--recursive'],
  integration: ['--require', 'spec/config/setup', 'spec/integration', '--recursive'],
  'end-to-end': ['--require', 'spec/config/setup', 'spec/end-to-end', '--recursive'],
};

const outIdx = process.argv.indexOf('--out');
if (outIdx === -1 || !process.argv[outIdx + 1]) {
  console.error('Usage: node scripts/test-counts.mjs --out <file.json>');
  process.exit(2);
}
const outFile = resolve(process.argv[outIdx + 1]);

const env = {...process.env};
delete env.EXCEL_BUILD;

const stderrTail = res => (res.stderr || '').trim().split('\n').slice(-20).join('\n');

const suites = {};
for (const [name, args] of Object.entries(SUITES)) {
  const res = spawnSync(
    join(root, 'node_modules/.bin/mocha'),
    [...args, '--reporter', 'json'],
    {cwd: root, env, encoding: 'utf8', stdio: ['ignore', 'pipe', 'pipe'], maxBuffer: 1 << 29}
  );
  if (res.error) {
    throw new Error(`suite ${name}: failed to run mocha: ${res.error.message}\n${stderrTail(res)}`);
  }
  // mocha 7 json reporter has no output option and specs may log to stdout: the report is the
  // last thing printed, so parse from the last top-level `{ "stats"` onwards.
  const start = res.stdout.lastIndexOf('{\n  "stats"');
  let report;
  try {
    if (start === -1) throw new Error('report marker not found in stdout');
    report = JSON.parse(res.stdout.slice(start));
  } catch (err) {
    throw new Error(
      `suite ${name}: no parsable mocha JSON (exit ${res.status}): ${err.message}\n${stderrTail(res)}`
    );
  }
  const {tests, passes, failures, pending} = report.stats;
  suites[name] = {
    tests,
    passes,
    failures,
    pending,
    failingTitles: report.failures.map(f => f.fullTitle),
  };
  console.log(`${name}: tests=${tests} passes=${passes} failures=${failures} pending=${pending}`);
}

mkdirSync(dirname(outFile), {recursive: true});
writeFileSync(outFile, `${JSON.stringify({node: process.version, suites}, null, 2)}\n`);
