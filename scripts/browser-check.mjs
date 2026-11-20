// Opens every sample file in a real browser (installed Chrome or Edge, driven by
// playwright-core) using the built package from dist/, checks what was rendered,
// verifies that scripts embedded in files never run, and saves screenshots.
//
//   npm run test:browser
//
// Sample files come from the test suite's global setup, so run `npm test` once first.
import { createReadStream, existsSync, mkdirSync, statSync } from 'node:fs';
import { createServer } from 'node:http';
import { extname, join, normalize, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

const root = resolve(fileURLToPath(import.meta.url), '../..');
const samples = join(root, 'tests/.generated');
const screenshots = join(samples, 'screenshots');
const systemFont = 'C:/Windows/Fonts/arial.ttf';

if (!existsSync(join(root, 'dist/index.js'))) {
  console.error('dist/ is missing. Run "npm run build" first.');
  process.exit(1);
}
if (!existsSync(join(samples, 'sample.docx'))) {
  console.error('Sample files are missing. Run "npm test" once to generate them.');
  process.exit(1);
}
mkdirSync(screenshots, { recursive: true });

const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
  '.json': 'application/json',
  '.map': 'application/json',
  '.css': 'text/css',
  '.svg': 'image/svg+xml',
  '.png': 'image/png',
  '.pdf': 'application/pdf',
  '.wav': 'audio/wav',
  '.wasm': 'application/wasm',
};

const server = createServer((request, response) => {
  const path = decodeURIComponent(new URL(request.url, 'http://localhost').pathname);
  const file = path === '/system-font.ttf' ? systemFont : normalize(join(root, path));
  if (path !== '/system-font.ttf' && !file.startsWith(root)) {
    response.writeHead(403).end();
    return;
  }
  if (!existsSync(file) || !statSync(file).isFile()) {
    response.writeHead(404).end('not found');
    return;
  }
  response.writeHead(200, { 'content-type': MIME[extname(file)] ?? 'application/octet-stream' });
  createReadStream(file).pipe(response);
});
await new Promise((done) => server.listen(0, '127.0.0.1', done));
const origin = `http://127.0.0.1:${server.address().port}`;

async function launch() {
  for (const channel of ['chrome', 'msedge']) {
    try {
      return await chromium.launch({ channel, headless: true });
    } catch {
      // Try the next installed browser.
    }
  }
  throw new Error('Neither Chrome nor Edge could be started.');
}

/** Each check opens one file and returns a list of problems (empty when all is well). */
const text = (selector = '.av-content') =>
  `(document.querySelector(${JSON.stringify(selector)})?.textContent ?? '')`;
const count = (selector) => `document.querySelectorAll(${JSON.stringify(selector)}).length`;

const CHECKS = [
  {
    file: 'sample.txt',
    format: 'text',
    expect: [[`${text()}.includes('Plain text sample.')`, 'text shown']],
  },
  {
    file: 'sample.ts',
    format: 'text',
    expect: [[`${count('.av-tok-keyword')} >= 3`, 'keywords highlighted']],
  },
  {
    file: 'sample.json',
    format: 'json',
    expect: [[`${text()}.includes('"count": 3')`, 'JSON pretty-printed']],
  },
  {
    file: 'sample.xml',
    format: 'xml',
    expect: [[`${count('.av-tok-tag')} > 0`, 'tags highlighted']],
  },
  {
    file: 'sample.md',
    format: 'markdown',
    expect: [
      [`${text('.av-prose h1')} === 'Markdown sample'`, 'heading'],
      [`${count('.av-prose table td')} === 2`, 'table'],
      [`${count('.av-prose li.av-task input')} === 2`, 'task list'],
    ],
  },
  {
    file: 'sample.csv',
    format: 'csv',
    expect: [
      [
        `[...document.querySelectorAll('.av-grid-c')].some((c) => c.textContent === 'Gadget, large')`,
        'quoted cell',
      ],
      [`document.querySelector('.av-grid').scrollHeight > 0`, 'grid has layout'],
    ],
  },
  {
    file: 'sample.html',
    format: 'html',
    wait: 400,
    expect: [
      [`document.querySelector('iframe').getAttribute('sandbox') === ''`, 'sandboxed frame'],
      [`!window.__messages.includes('html-script-ran')`, 'embedded script did not run'],
    ],
  },
  {
    file: 'sample.svg',
    format: 'svg',
    wait: 300,
    expect: [
      [`document.querySelector('.av-image img').naturalWidth === 200`, 'SVG decoded'],
      [`!window.__messages.includes('svg-script-ran')`, 'SVG script did not run'],
    ],
  },
  {
    file: 'sample.png',
    format: 'image',
    expect: [[`document.querySelector('.av-image img').naturalWidth === 1`, 'PNG decoded']],
  },
  {
    file: 'sample.wav',
    format: 'audio',
    expect: [[`Math.round(document.querySelector('audio').duration) === 1`, 'audio metadata read']],
  },
  {
    file: 'sample.pdf',
    format: 'pdf',
    label: 'pdf-native',
    expect: [
      [
        `document.querySelector('iframe.av-frame').src.startsWith('blob:')`,
        'PDF framed from a blob URL',
      ],
    ],
  },
  {
    file: 'sample.pdf',
    format: 'pdf',
    label: 'pdf-pdfjs',
    options: { pdfjs: true },
    wait: 1500,
    expect: [
      [`${count('.av-page canvas')} === 1`, 'page drawn to canvas'],
      [
        `(() => { const c = document.querySelector('.av-page canvas'); const d = c.getContext('2d').getImageData(0, 0, c.width, c.height).data; for (let i = 0; i < d.length; i += 4) if (d[i] < 128) return true; return false; })()`,
        'canvas contains drawn text',
      ],
      [`window.__viewer.getPage().pageCount === 1`, 'page count'],
    ],
  },
  {
    file: 'sample.xlsx',
    format: 'xlsx',
    expect: [
      [
        `[...document.querySelectorAll('.av-grid-c')].some((c) => c.textContent === '26%')`,
        'number format (percent)',
      ],
      [
        `[...document.querySelectorAll('.av-grid-c')].some((c) => c.textContent === '2023-03-15')`,
        'date format',
      ],
      [`${count('.av-tab')} === 3`, 'sheet tabs'],
    ],
  },
  {
    file: 'real.xlsx',
    format: 'xlsx',
    expect: [
      [
        `[...document.querySelectorAll('.av-grid-c')].some((c) => c.textContent === '2024-01-31')`,
        'ExcelJS date',
      ],
      [
        `[...document.querySelectorAll('.av-grid-c')].some((c) => c.textContent === 'Merged note' && c.offsetWidth > 300)`,
        'merged cell spans columns',
      ],
    ],
  },
  {
    file: 'sample.ods',
    format: 'ods',
    expect: [
      [
        `[...document.querySelectorAll('.av-grid-c')].some((c) => c.textContent === 'Widget')`,
        'cell text',
      ],
    ],
  },
  {
    file: 'sample.docx',
    format: 'docx',
    wait: 300,
    expect: [
      [`${text('.av-paper h1')} === 'Annual Report'`, 'heading'],
      [`${count('.av-paper table')} === 2`, 'tables'],
      [
        `[...document.querySelectorAll('.av-paper img')].every((i) => i.naturalWidth > 0)`,
        'images decoded',
      ],
    ],
  },
  {
    file: 'real.docx',
    format: 'docx',
    wait: 300,
    expect: [
      [`${text('.av-paper h1')} === 'Project Report'`, 'heading'],
      [`document.querySelector('.av-paper img').naturalWidth > 0`, 'image decoded'],
      [`document.querySelector('.av-paper').offsetWidth > 600`, 'paper has page width'],
    ],
  },
  {
    file: 'sample.odt',
    format: 'odt',
    wait: 300,
    expect: [[`${text('.av-paper h2')} === 'Chapter'`, 'heading']],
  },
  {
    file: 'sample.rtf',
    format: 'rtf',
    wait: 300,
    expect: [
      [`${text('.av-paper')}.includes('Café €uro')`, 'code page and Unicode text'],
      [`document.querySelector('.av-paper img').naturalWidth === 1`, 'embedded picture decoded'],
    ],
  },
  {
    file: 'sample.pptx',
    format: 'pptx',
    wait: 800,
    expect: [
      [
        `${count('.av-slide')} >= 1 && ${count('.av-page')} === 4`,
        'visible slides built, one frame per slide',
      ],
      [`${text('.av-slide')}.includes('Quarterly Review')`, 'title text'],
      [`document.querySelector('.av-page').offsetWidth > 300`, 'slide scaled to width'],
    ],
  },
  {
    file: 'real.pptx',
    format: 'pptx',
    wait: 800,
    expect: [
      [`${text('.av-slide')}.includes('Quarterly Results')`, 'PptxGenJS title'],
      [
        `[...document.querySelectorAll('.av-slide img')].every((i) => i.naturalWidth > 0)`,
        'picture decoded',
      ],
      [`window.__viewer.getPage().pageCount === 2`, 'page count'],
    ],
  },
  {
    file: 'sample.odp',
    format: 'odp',
    wait: 800,
    expect: [[`${text('.av-slide')}.includes('Deck title')`, 'title text']],
  },
  {
    file: 'sample.epub',
    format: 'epub',
    wait: 400,
    expect: [
      [`document.querySelector('iframe').getAttribute('sandbox') === ''`, 'sandboxed frame'],
      [`window.__viewer.getPage().pageCount === 2`, 'chapter count'],
    ],
  },
  {
    file: 'sample.zip',
    format: 'zip',
    expect: [[`${count('.av-list tbody tr')} === 2`, 'entries listed']],
  },
  {
    file: 'sample.ipynb',
    format: 'ipynb',
    wait: 300,
    expect: [
      [`${text('.av-prose h1')} === 'Analysis'`, 'markdown cell'],
      [`document.querySelector('.av-nb-cell img').naturalWidth === 1`, 'image output decoded'],
    ],
  },
  {
    file: 'sample.eml',
    format: 'eml',
    wait: 300,
    expect: [
      [`${text('.av-mail-head')}.includes('Quarterly report')`, 'subject'],
      [
        `document.querySelector('.av-mail iframe').getAttribute('sandbox') === ''`,
        'sandboxed HTML body',
      ],
    ],
  },
  ...(existsSync(systemFont)
    ? [
        {
          file: '../../system-font.ttf',
          url: '/system-font.ttf',
          format: 'font',
          label: 'font',
          expect: [
            [`${count('.av-font p')} === 9`, 'specimen lines'],
            [
              `document.fonts.check('20px ' + document.querySelector('.av-font p').style.fontFamily)`,
              'font registered',
            ],
          ],
        },
      ]
    : []),
  {
    file: 'legacy.doc',
    error: 'UNSUPPORTED_FORMAT',
    expect: [[`${text('.av-status')}.includes('Preview is not available')`, 'fallback card shown']],
  },
  {
    file: 'missing-file.docx',
    error: 'FETCH_FAILED',
    expect: [[`${text('.av-status')}.includes('HTTP 404')`, 'error card shown']],
  },
];

const browser = await launch();
const page = await browser.newPage({ viewport: { width: 1100, height: 760 } });
const pageErrors = [];
let blockedScripts = 0;
page.on('pageerror', (error) => pageErrors.push(String(error)));
page.on('console', (message) => {
  // 404s are expected for the missing-file check; everything else is a real error.
  const text = message.text();
  // The sandbox reports each script it refuses to run: that is the protection working.
  if (text.includes('Blocked script execution')) blockedScripts++;
  else if (message.type() === 'error' && !text.includes('404')) pageErrors.push(text);
});

const failures = [];
const rows = [];
try {
  await page.goto(`${origin}/tests/browser/harness.html`);
  await page.waitForFunction('window.__ready === true');

  for (const check of CHECKS) {
    const label = check.label ?? check.file;
    const url = check.url ?? `/tests/.generated/${check.file}`;
    const problems = [];
    const result = await page.evaluate(
      ([target, options]) => window.__open(target, options),
      [url, check.options ?? {}],
    );
    if (check.error) {
      if (result.error?.code !== check.error) {
        problems.push(`expected error ${check.error}, got ${JSON.stringify(result)}`);
      }
    } else if (result.state !== 'ready') {
      problems.push(`did not load: ${JSON.stringify(result)}`);
    } else if (result.format !== check.format) {
      problems.push(`detected as "${result.format}", expected "${check.format}"`);
    }
    if (check.wait) await page.waitForTimeout(check.wait);
    if (problems.length === 0) {
      for (const [expression, description] of check.expect) {
        const ok = await page.evaluate(expression).catch((error) => `threw: ${error.message}`);
        if (ok !== true) problems.push(`${description} (${ok})`);
      }
    }
    await page.screenshot({ path: join(screenshots, `${label.replace(/[^\w.-]+/g, '_')}.png`) });
    rows.push({ file: label, result: problems.length ? `FAIL: ${problems.join('; ')}` : 'ok' });
    if (problems.length) failures.push(label);
  }

  // Zoom, paging and dark theme through the public API.
  const interactions = await page.evaluate(async () => {
    const opened = await window.__open('/tests/.generated/real.pptx', { theme: 'dark' });
    const viewer = window.__viewer;
    const before = viewer.getZoom();
    viewer.zoomIn();
    viewer.goToPage(2);
    return {
      state: opened.state,
      zoomChanged: viewer.getZoom() > before,
      page: viewer.getPage().page,
      theme: viewer.element.getAttribute('data-theme'),
      darkBackground: getComputedStyle(viewer.element).backgroundColor,
    };
  });
  await page.waitForTimeout(500);
  await page.screenshot({ path: join(screenshots, 'dark-theme-zoomed.png') });
  const interactionsOk =
    interactions.state === 'ready' &&
    interactions.zoomChanged &&
    interactions.page === 2 &&
    interactions.darkBackground === 'rgb(13, 17, 23)';
  rows.push({
    file: 'zoom, paging, dark theme',
    result: interactionsOk ? 'ok' : `FAIL: ${JSON.stringify(interactions)}`,
  });
  if (!interactionsOk) failures.push('interactions');

  // Nested preview: open a file inside the ZIP, then go back.
  const nested = await page.evaluate(async () => {
    await window.__open('/tests/.generated/sample.zip');
    [...document.querySelectorAll('.av-link')]
      .find((b) => b.textContent.endsWith('readme.md'))
      .click();
    for (let i = 0; i < 100 && !document.querySelector('.av-child h1'); i++) {
      await new Promise((r) => setTimeout(r, 20));
    }
    const heading = document.querySelector('.av-child h1')?.textContent;
    document.querySelector('.av-child-bar button').click();
    return { heading, closed: document.querySelector('.av-child') === null };
  });
  const nestedOk = nested.heading === 'Inside the archive' && nested.closed;
  rows.push({
    file: 'nested preview (zip entry)',
    result: nestedOk ? 'ok' : `FAIL: ${JSON.stringify(nested)}`,
  });
  if (!nestedOk) failures.push('nested');

  // The custom element, loaded from the built dist/element.js.
  await page.goto(`${origin}/tests/browser/element.html`);
  const elementOk = await page
    .waitForFunction(
      `document.querySelector('any-view h1')?.textContent === 'Markdown sample'`,
      null,
      { timeout: 5000 },
    )
    .then(
      () => true,
      () => false,
    );
  await page.screenshot({ path: join(screenshots, 'custom-element.png') });
  rows.push({
    file: '<any-view> custom element',
    result: elementOk ? 'ok' : 'FAIL: content did not render',
  });
  if (!elementOk) failures.push('custom element');

  // The runnable examples shipped in examples/.
  for (const [name, path, ready] of [
    [
      'example: vanilla',
      '/examples/vanilla/index.html?src=/tests/.generated/sample.md',
      `document.querySelector('#preview h1')?.textContent === 'Markdown sample' && document.getElementById('status').textContent.includes('Markdown')`,
    ],
    [
      'example: web component',
      '/examples/web-component/index.html?src=/tests/.generated/sample.csv',
      `document.querySelectorAll('any-view .av-grid-c').length > 0 && document.getElementById('status').textContent.includes('CSV')`,
    ],
  ]) {
    await page.goto(origin + path);
    const ok = await page.waitForFunction(ready, null, { timeout: 5000 }).then(
      () => true,
      () => false,
    );
    await page.screenshot({ path: join(screenshots, name.replace(/[^w.-]+/g, '_') + '.png') });
    rows.push({ file: name, result: ok ? 'ok' : 'FAIL: example did not render' });
    if (!ok) failures.push(name);
  }
} finally {
  await browser.close();
  server.close();
}

console.table(rows);
if (pageErrors.length) {
  console.error('Errors reported by the page:\n  ' + pageErrors.join('\n  '));
  failures.push('page errors');
}
console.log(`Browser: ${browser.version()}   Screenshots: ${screenshots}`);
console.log(`Scripts inside files that the sandbox refused to run: ${blockedScripts}`);
if (failures.length) {
  console.error(`\n${failures.length} check(s) failed.`);
  process.exit(1);
}
console.log(`\nAll ${rows.length} browser checks passed.`);
