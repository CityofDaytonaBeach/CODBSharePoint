import { test, expect } from '@playwright/test';
import * as http from 'http';
import * as fs from 'fs';
import * as path from 'path';

/**
 * Simple static file server for the built dist + test-browser.html
 * Serves the project root over HTTP so the browser can load ES modules.
 */
function createStaticServer(projectRoot: string): http.Server {
  const types: Record<string, string> = {
    '.html': 'text/html',
    '.js': 'text/javascript',
    '.mjs': 'text/javascript',
    '.css': 'text/css',
    '.wasm': 'application/wasm',
    '.json': 'application/json'
  };

  return http.createServer((req, res) => {
    const url = new URL(req.url || '/', `http://${req.headers.host}`);
    let filePath = path.join(projectRoot, decodeURIComponent(url.pathname));
    if (filePath.endsWith('/') || !fs.existsSync(filePath)) {
      filePath = path.join(filePath, 'index.html');
    }
    if (!fs.existsSync(filePath)) {
      res.writeHead(404);
      res.end('Not found');
      return;
    }
    const ext = path.extname(filePath);
    res.writeHead(200, { 'Content-Type': types[ext] || 'application/octet-stream' });
    fs.createReadStream(filePath).pipe(res);
  });
}

test.describe.configure({ mode: 'serial' });

test('full pipeline runs entirely in the browser', async ({ browser }) => {
  const projectRoot = path.resolve(__dirname, '../..');
  const server = createStaticServer(projectRoot);
  await new Promise<void>(resolve => server.listen(0, () => resolve()));
  const address = server.address() as { port: number };
  const baseUrl = `http://localhost:${address.port}`;

  try {
    const page = await browser.newPage();

    // Capture console messages and page errors
    const consoleMessages: string[] = [];
    const pageErrors: string[] = [];
    page.on('console', msg => consoleMessages.push(`${msg.type()}: ${msg.text()}`));
    page.on('pageerror', err => pageErrors.push(err.message));

    // Load test-browser.html
    await page.goto(`${baseUrl}/test-browser.html`, { waitUntil: 'networkidle' });

    // Wait for SDK to load (button enabled)
    await page.waitForSelector('#btn-run:not([disabled])', { timeout: 60000 });

    // Click run
    await page.click('#btn-run');

    // Wait for the test to complete (final summary line in the log).
    // The status bar's single dot is only green at completion, but it also turns
    // green earlier (SDK load). So wait on the terminal log summary instead.
    await page.waitForFunction(() => {
      const log = document.getElementById('log');
      return log && log.textContent && log.textContent.includes('=== Test Complete in');
    }, undefined, { timeout: 120000 });

    // Wait a tick for the final status to render
    await page.waitForSelector('.status-row .dot.green', { timeout: 10000 });

    // Check status
    const status = await page.textContent('#runtime-status');
    expect(status).toContain('ALL TESTS PASSED');

    // Log output should contain key steps
    const logText = await page.textContent('#log');
    expect(logText).toContain('Knowledge');
    expect(logText).toContain('AI Contract');
    expect(logText).toContain('Full Build Pipeline');
    expect(logText).toContain('SPPKG Package Validation');
    expect(logText).toContain('Browser Production Smoke Test');
    expect(logText).toContain('ALL PASSED');

    // Verify SPPKG produced
    expect(logText).toContain('SPPKG Package Validation');
    expect(logText).toContain('Is ZIP: true');
    expect(logText).toContain('Package size:');

    // Download button should be enabled after successful build
    const downloadDisabled = await page.isDisabled('#btn-download');
    expect(downloadDisabled).toBe(false);

    // Download the .sppkg
    const [download] = await Promise.all([
      page.waitForEvent('download'),
      page.click('#btn-download')
    ]);
    const suggestedFilename = download.suggestedFilename();
    expect(suggestedFilename).toMatch(/\.sppkg$/);
    const downloadPath = await download.path();
    const bytes = fs.readFileSync(downloadPath!);
    expect(bytes.length).toBeGreaterThan(0);
    // Verify ZIP signature
    expect(bytes[0]).toBe(0x50);
    expect(bytes[1]).toBe(0x4B);

    // No fatal page errors
    expect(pageErrors.filter(e => !e.includes('esbuild-wasm'))).toEqual([]);
  } finally {
    server.close();
  }
});

test('builds a multi-component solution in-browser', async ({ browser }) => {
  const projectRoot = path.resolve(__dirname, '../..');
  const server = createStaticServer(projectRoot);
  await new Promise<void>(resolve => server.listen(0, () => resolve()));
  const address = server.address() as { port: number };
  const baseUrl = `http://localhost:${address.port}`;

  try {
    const page = await browser.newPage();
    await page.goto(`${baseUrl}/test-browser.html`, { waitUntil: 'networkidle' });
    await page.waitForSelector('#btn-run:not([disabled])', { timeout: 60000 });
    await page.click('#btn-run');
    await page.waitForFunction(() => {
      const log = document.getElementById('log');
      return log && log.textContent && log.textContent.includes('=== Test Complete in');
    }, undefined, { timeout: 120000 });

    const logText = await page.textContent('#log');
    expect(logText).toContain('ALL PASSED');
  } finally {
    server.close();
  }
});
