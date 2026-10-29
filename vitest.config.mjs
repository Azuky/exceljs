import { configDefaults, defineConfig } from 'vitest/config';

const shared = {
  globals: true,
  fileParallelism: false,
  pool: 'forks',
  testTimeout: 2000,
  hookTimeout: 2000,
  setupFiles: ['spec/config/setup.mjs'],
};

export default defineConfig({
  test: {
    projects: [
      {
        test: {
          ...shared,
          name: 'unit',
          include: ['spec/unit/**/*.spec.js'],
          // zero-byte placeholder: mocha tolerates an empty file, vitest fails it with "No test suite found"
          exclude: [...configDefaults.exclude, 'spec/unit/utils/encryptor.spec.js'],
        },
      },
      { test: { ...shared, name: 'integration', include: ['spec/integration/**/*.spec.js'] } },
      { test: { ...shared, name: 'end-to-end', include: ['spec/end-to-end/**/*.spec.js'] } },
    ],
  },
});
