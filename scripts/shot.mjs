// Developer tool: opens one file in real Chrome/Edge with the built package and
// saves screenshots, for comparing a rendering with the original application.
//
//   node scripts/shot.mjs <file> [outPrefix] [--scroll=0,900,1800] [--sheet=1] [--width=1300] [--height=900] [--eval=js] [--gridlines=0]
import { createReadStream, existsSync, mkdirSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { basename, dirname, extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

const root = resolve(fileURLToPath(import.meta.url), '../..');
const args = process.argv.slice(2);
const flags = Object.fromEntries(
  args.filter((a) => a.startsWith('--')).map((a) => a.slice(2).split(/=(.*)/s).slice(0, 2)),
);
const [file, prefixArg] = args.filter((a) => !a.startsWith('--'));
if (!file || !existsSync(file)) {
  console.error('Usage: node scripts/shot.mjs <file> [outPrefix]');
  process.exit(1);
}
const target = resolve(file);
const prefix = resolve(prefixArg ?? join(root, 'tests/.generated/shots', basename(file)));
mkdirSync(dirname(prefix), { recursive: true });

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.pdf': 'application/pdf',
};
const server = createServer((request, response) => {
  const path = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
  const local = path.startsWith('/__file/')
    ? target
    : path.startsWith('/dist/') && flags.dist
      ? normalize(join(root, flags.dist, path.slice(6)))
      : normalize(join(root, path));
  if (!existsSync(local) || !statSync(local).isFile()) {
    response.writeHead(404).end('not found');
    return;
  }
  response.writeHead(200, { 'content-type': MIME[extname(local)] ?? 'application/octet-stream' });
  createReadStream(local).pipe(response);
});
await new Promise((done) => server.listen(0, '127.0.0.1', done));
const origin = `http://127.0.0.1:${server.address().port}`;

let browser;
for (const channel of ['chrome', 'msedge']) {
  try {
    browser = await chromium.launch({ channel, headless: true });
    break;
  } catch {
    // Try the next installed browser.
  }
}
const page = await browser.newPage({
  viewport: { width: Number(flags.width ?? 1300), height: Number(flags.height ?? 900) },
});
page.on('console', (message) => {
  if (message.type() === 'error') console.log('console error:', message.text());
});
page.on('pageerror', (error) => console.log('page error:', error.message));
await page.goto(`${origin}/tests/browser/harness.html`);
await page.waitForFunction('window.__ready === true');
const result = await page.evaluate(
  ([url, theme, pdfjs, gridlines]) => window.__open(url, { theme, pdfjs, gridlines }),
  [
    `/__file/${encodeURIComponent(basename(file))}`,
    flags.theme ?? 'light',
    Boolean(flags.pdfjs),
    flags.gridlines === undefined ? null : flags.gridlines !== '0',
  ],
);
console.log(JSON.stringify(result));
if (flags.sheet) {
  await page.locator('.av-tab').nth(Number(flags.sheet)).click();
  await page.waitForTimeout(400);
}
await page.waitForTimeout(600);
const scrolls = String(flags.scroll ?? '0')
  .split(',')
  .map(Number);
for (const top of scrolls) {
  await page.evaluate((y) => {
    const scroller = document.querySelector('.av-pages, .av-grid, .av-scroll');
    if (scroller) scroller.scrollTop = y;
  }, top);
  await page.waitForTimeout(300);
  const path = `${prefix}-${top}.png`;
  await page.screenshot({ path });
  console.log(path);
}
if (flags.eval) console.log(JSON.stringify(await page.evaluate(flags.eval), null, 1));
await browser.close();
server.close();
