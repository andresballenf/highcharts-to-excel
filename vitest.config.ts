import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    setupFiles: ['tests/setup/jsdom-stubs.ts'],
    testTimeout: 30_000,
    hookTimeout: 30_000,
    // Every vi.spyOn is restored before each test; vi.fn() implementations are left alone.
    restoreMocks: true,
    mockReset: false,
    coverage: {
      provider: 'v8',
      include: ['src/**'],
      reporter: ['text-summary', 'html', 'json-summary'],
      reportsDirectory: 'tests/output/coverage',
      // Floors just under the measured totals (lines 91.8, functions 93.8, branches 79.9, statements 88.7).
      thresholds: { lines: 90, functions: 92, branches: 78, statements: 87 },
    },
  },
});
