import { defineConfig } from 'tsdown';

const shared = {
  platform: 'node',
  target: 'node22',
  sourcemap: true,
  dts: false,
  // dist/ is emptied by the `build` script (pnpm run clean) before tsdown runs
  clean: false,
} as const;

export default defineConfig([
  {
    ...shared,
    entry: { index: 'lib/exceljs.nodejs.js' },
    format: 'cjs',
  },
  {
    // thin ESM wrapper over dist/index.cjs: keep the sibling CJS bundle external
    ...shared,
    entry: { index: 'scripts/build/esm-wrapper.mjs' },
    format: 'esm',
    plugins: [
      {
        name: 'external-cjs-bundle',
        resolveId(id) {
          if (id === './index.cjs') return { id, external: true };
        },
      },
    ],
  },
]);
