import { defineConfig } from 'vitest/config';

export default defineConfig({
  test: {
    include: ['tests/**/*.test.ts'],
    environment: 'jsdom',
    setupFiles: ['tests/setup/jsdom-stubs.ts'],
    testTimeout: 30_000,
    hookTimeout: 30_000,
  },
});
