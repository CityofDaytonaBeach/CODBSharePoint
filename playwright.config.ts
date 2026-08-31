import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests/browser-e2e',
  timeout: 120000,
  use: {
    headless: true,
    browserName: 'chromium'
  }
});
