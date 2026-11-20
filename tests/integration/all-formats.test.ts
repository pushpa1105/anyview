import { afterEach, describe, expect, it, vi } from 'vitest';
import { createViewer, SUPPORTED_FORMAT_IDS } from '../../src';
import { BUILTIN_RENDERERS } from '../../src/core/registry';
import { binaryFile, mount, textFile, waitFor } from '../helpers/dom';
import {
  docxFixture,
  EML_FIXTURE,
  epubFixture,
  NOTEBOOK_FIXTURE,
  odpFixture,
  odsFixture,
  odtFixture,
  pptxFixture,
  RTF_FIXTURE,
  xlsxFixture,
} from '../helpers/fixtures';
import { makeZip, PNG_1X1 } from '../helpers/zip';

afterEach(() => {
  vi.unstubAllGlobals();
  document.body.textContent = '';
});

interface Case {
  format: string;
  file: () => File;
  /** Text expected somewhere in the rendered content. */
  shows?: string;
  /** Element expected in the rendered content. */
  selector?: string;
}

const CASES: Case[] = [
  { format: 'text', file: () => textFile('plain words', 'notes.txt'), shows: 'plain words' },
  { format: 'json', file: () => textFile('{"a":1}', 'data.json'), shows: '"a": 1' },
  { format: 'xml', file: () => textFile('<a><b/></a>', 'feed.xml'), selector: '.av-tok-tag' },
  { format: 'csv', file: () => textFile('a,b\n1,2', 'table.csv'), selector: '.av-grid' },
  { format: 'tsv', file: () => textFile('a\tb\n1\t2', 'table.tsv'), selector: '.av-grid' },
  { format: 'markdown', file: () => textFile('# Heading', 'readme.md'), selector: 'h1' },
  { format: 'html', file: () => textFile('<p>page</p>', 'page.html'), selector: 'iframe[sandbox]' },
  { format: 'pdf', file: () => textFile('%PDF-1.7', 'doc.pdf'), selector: 'iframe' },
  { format: 'xlsx', file: () => binaryFile(xlsxFixture(), 'book.xlsx'), selector: '.av-tabs' },
  { format: 'ods', file: () => binaryFile(odsFixture(), 'book.ods'), selector: '.av-tabs' },
  { format: 'docx', file: () => binaryFile(docxFixture(), 'report.docx'), shows: 'Annual Report' },
  { format: 'odt', file: () => binaryFile(odtFixture(), 'report.odt'), shows: 'Chapter' },
  { format: 'pptx', file: () => binaryFile(pptxFixture(), 'deck.pptx'), selector: '.av-page' },
  { format: 'odp', file: () => binaryFile(odpFixture(), 'deck.odp'), selector: '.av-page' },
  { format: 'rtf', file: () => textFile(RTF_FIXTURE, 'letter.rtf'), shows: 'Heading' },
  {
    format: 'epub',
    file: () => binaryFile(epubFixture(), 'book.epub'),
    selector: 'iframe[sandbox]',
  },
  { format: 'zip', file: () => binaryFile(makeZip({ 'a.txt': 'x' }), 'files.zip'), shows: 'a.txt' },
  { format: 'ipynb', file: () => textFile(NOTEBOOK_FIXTURE, 'nb.ipynb'), shows: 'Analysis' },
  { format: 'eml', file: () => textFile(EML_FIXTURE, 'mail.eml'), shows: 'Quarterly report' },
];

describe('every supported format opens through createViewer', () => {
  it.each(CASES)('$format', async ({ format, file, shows, selector }) => {
    const container = mount();
    const viewer = createViewer(container);
    const loaded = vi.fn();
    viewer.on('load', loaded);
    await viewer.load(file());
    expect(viewer.state).toBe('ready');
    expect(viewer.format?.id).toBe(format);
    expect(loaded).toHaveBeenCalledTimes(1);
    const content = container.querySelector('.av-content')!;
    if (shows) await waitFor(() => content.textContent?.includes(shows));
    if (selector) expect(content.querySelector(selector)).not.toBeNull();
    viewer.destroy();
    expect(container.childNodes).toHaveLength(0);
  });

  it('image and svg (which wait for the browser to decode them)', async () => {
    for (const [name, format] of [
      ['photo.png', 'image'],
      ['logo.svg', 'svg'],
    ] as const) {
      const container = mount();
      const viewer = createViewer(container);
      const pending = viewer.load(binaryFile(PNG_1X1, name));
      const image = await waitFor(() => container.querySelector('img'));
      image.dispatchEvent(new Event('load'));
      await pending;
      expect(viewer.format?.id).toBe(format);
      expect(viewer.getZoom()).toBe(1);
      viewer.destroy();
    }
  });

  it('video and audio (which wait for metadata)', async () => {
    for (const [name, format, tag] of [
      ['clip.mp4', 'video', 'video'],
      ['song.mp3', 'audio', 'audio'],
    ] as const) {
      const container = mount();
      const viewer = createViewer(container);
      const pending = viewer.load(binaryFile(new Uint8Array(8), name));
      const player = await waitFor(() => container.querySelector<HTMLMediaElement>(tag));
      await waitFor(() => player.getAttribute('src'));
      player.pause = vi.fn();
      player.load = vi.fn();
      player.dispatchEvent(new Event('loadedmetadata'));
      await pending;
      expect(viewer.format?.id).toBe(format);
      viewer.destroy();
    }
  });

  it('font', async () => {
    class FakeFontFace {
      load = (): Promise<void> => Promise.resolve();
    }
    vi.stubGlobal('FontFace', FakeFontFace);
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: { add: vi.fn(), delete: vi.fn() },
    });
    const container = mount();
    const viewer = createViewer(container);
    await viewer.load(binaryFile(new Uint8Array(8), 'type.woff2'));
    expect(viewer.format?.id).toBe('font');
    expect(container.querySelector('.av-font')).not.toBeNull();
  });

  it('covers every format that has a built-in renderer', () => {
    const tested = new Set([
      ...CASES.map((item) => item.format),
      'image',
      'svg',
      'video',
      'audio',
      'font',
    ]);
    expect([...SUPPORTED_FORMAT_IDS].sort()).toEqual([...tested].sort());
  });

  it('loads each built-in renderer module', async () => {
    for (const definition of BUILTIN_RENDERERS) {
      const renderer = await definition.load();
      expect(typeof renderer.render, definition.id).toBe('function');
    }
  });
});

describe('nested previews end to end', () => {
  it('opens a file inside a ZIP archive and returns to the listing', async () => {
    const container = mount();
    const viewer = createViewer(container);
    await viewer.load(
      binaryFile(makeZip({ 'docs/readme.md': '# Inside the archive' }), 'bundle.zip'),
    );
    container.querySelector<HTMLButtonElement>('.av-link')!.click();
    const heading = await waitFor(() => container.querySelector('.av-child h1'));
    expect(heading.textContent).toBe('Inside the archive');
    container.querySelector<HTMLButtonElement>('.av-child-bar button')!.click();
    expect(container.querySelector('.av-child')).toBeNull();
    expect(container.querySelector('.av-list')).not.toBeNull();
  });

  it('opens an email attachment', async () => {
    const container = mount();
    const viewer = createViewer(container);
    await viewer.load(textFile(EML_FIXTURE, 'mail.eml'));
    container.querySelector<HTMLButtonElement>('.av-mail-attachments button')!.click();
    await waitFor(() => container.querySelector('.av-child .av-grid'));
    expect(container.querySelector('.av-child .av-file-name')?.textContent).toBe('data.csv');
  });
});

describe('untrusted content stays inert', () => {
  it('never creates script elements or event handlers in the host page', async () => {
    const hostile = [
      textFile('<img src=x onerror="window.pwned=1"><script>window.pwned=1</script>', 'a.md'),
      textFile('<script>window.pwned=1</script>', 'a.txt'),
      textFile('<a onclick="window.pwned=1"><script>window.pwned=1</script></a>', 'a.xml'),
      textFile('"<script>window.pwned=1</script>",<img src=x onerror=alert(1)>', 'a.csv'),
      textFile(
        JSON.stringify({
          cells: [{ cell_type: 'markdown', source: '<script>window.pwned=1</script>' }],
        }),
        'a.ipynb',
      ),
      textFile('Subject: <script>window.pwned=1</script>\n\n<img src=x onerror=alert(1)>', 'a.eml'),
      textFile(String.raw`{\rtf1 <script>window.pwned=1</script>}`, 'a.rtf'),
    ];
    for (const file of hostile) {
      const container = mount();
      const viewer = createViewer(container);
      await viewer.load(file);
      expect(container.querySelector('script'), file.name).toBeNull();
      expect(container.querySelector('[onerror], [onclick], [onload]'), file.name).toBeNull();
      viewer.destroy();
    }
    expect((window as { pwned?: number }).pwned).toBeUndefined();
  });
});
