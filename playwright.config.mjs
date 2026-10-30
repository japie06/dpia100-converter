// UI tests for dpia100.html: npm run test:ui. Uses the installed Microsoft Edge (channel 'msedge'):
// never run 'npx playwright install', no browser is downloaded. Edge is preinstalled on windows-latest.
import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: 'tests/web',
  testMatch: '*.spec.mjs',
  outputDir: 'tests/out/playwright',
  forbidOnly: !!process.env.CI,
  reporter: process.env.CI
    ? [['list'], ['html', { outputFolder: 'tests/out/playwright-report', open: 'never' }]]
    : [['list']],
  use: {
    channel: 'msedge',
    headless: true,
    acceptDownloads: true,
    trace: 'retain-on-failure',
    screenshot: 'only-on-failure'
  }
});
