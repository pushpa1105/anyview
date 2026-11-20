import { afterEach, describe, expect, it, vi } from 'vitest';
import { AnyViewError } from '../../src';
import type { PdfJsLike, RenderContext } from '../../src';
import { parseEmail, decodeHeaderValue } from '../../src/formats/eml';
import { expandSelfClosing, openEpub } from '../../src/formats/epub';
import { renderer as archiveRenderer } from '../../src/renderers/archive';
import { renderer as emailRenderer } from '../../src/renderers/email';
import { renderer as epubRenderer } from '../../src/renderers/epub';
import { renderer as fontRenderer } from '../../src/renderers/font';
import { renderer as htmlRenderer } from '../../src/renderers/html';
import { renderer as imageRenderer } from '../../src/renderers/image';
import { renderer as markdownRenderer } from '../../src/renderers/markdown';
import { renderer as mediaRenderer } from '../../src/renderers/media';
import { renderer as notebookRenderer } from '../../src/renderers/notebook';
import { renderer as pdfRenderer } from '../../src/renderers/pdf';
import { prettyJson, renderer as textRenderer } from '../../src/renderers/text';
import { withBase } from '../../src/ui/frame';
import { printNode } from '../../src/ui/print';
import { openZip } from '../../src/util/zip';
import { instanceOf, renderContext, waitFor } from '../helpers/dom';
import { EML_FIXTURE, epubFixture, NOTEBOOK_FIXTURE } from '../helpers/fixtures';
import { makeZip, PNG_1X1 } from '../helpers/zip';

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  document.body.textContent = '';
});

const encode = (text: string): Uint8Array => new TextEncoder().encode(text);

describe('text renderer', () => {
  it('shows line numbers and highlights known languages', async () => {
    const { context, container } = renderContext('const a = 1;\nlet b = "x";\n', 'app.js', 'text', {
      language: 'javascript',
    });
    const instance = instanceOf(await textRenderer.render(context));
    expect(container.querySelector('.av-code-gutter')?.textContent).toBe('1\n2');
    expect(container.querySelector('.av-code-lines')?.textContent).toBe(
      'const a = 1;\nlet b = "x";\n',
    );
    expect(container.querySelectorAll('.av-tok-keyword')).toHaveLength(2);
    instance.zoom!.set(2);
    expect(instance.zoom!.get()).toBe(2);
    expect(container.querySelector<HTMLElement>('.av-code-table')!.style.fontSize).toBe('26px');
    expect(instance.printable!()).toBe(container.querySelector('.av-code-lines'));
  });

  it('shows unknown languages as plain text and normalises line endings', async () => {
    const { context, container } = renderContext('a\r\nb\rc', 'notes.txt', 'text');
    await textRenderer.render(context);
    expect(container.querySelector('.av-code-lines')?.textContent).toBe('a\nb\nc');
    expect(container.querySelector('.av-code-gutter')?.textContent).toBe('1\n2\n3');
    expect(container.querySelector('[class^=av-tok]')).toBeNull();
  });

  it('numbers an empty file as one line', async () => {
    const { context, container } = renderContext('', 'empty.txt', 'text');
    await textRenderer.render(context);
    expect(container.querySelector('.av-code-gutter')?.textContent).toBe('1');
  });

  it('pretty-prints minified JSON', async () => {
    const { context, container } = renderContext('{"a":[1,2],"b":{}}', 'data.json', 'json', {
      language: 'json',
    });
    await textRenderer.render(context);
    expect(container.querySelector('.av-code-lines')?.textContent).toBe(
      '{\n  "a": [\n    1,\n    2\n  ],\n  "b": {}\n}',
    );
    expect(container.querySelector('.av-tok-attr')?.textContent).toBe('"a"');
  });

  it('truncates very large files with a notice and skips highlighting', async () => {
    const { context, container } = renderContext('x'.repeat(5_000_100), 'big.js', 'text', {
      language: 'javascript',
    });
    await textRenderer.render(context);
    expect(container.querySelector('.av-notice')?.textContent).toContain('4.8 MB');
    expect(container.querySelector('.av-code-lines')?.textContent).toHaveLength(5_000_000);
    expect(container.querySelector('[class^=av-tok]')).toBeNull();
  });
});

describe('prettyJson', () => {
  it('keeps number text exactly as written', () => {
    expect(prettyJson('{"big":12345678901234567890,"f":1.50}')).toBe(
      '{\n  "big": 12345678901234567890,\n  "f": 1.50\n}',
    );
  });

  it('leaves punctuation inside strings alone', () => {
    expect(prettyJson('["a,b:{c}","q\\"uote"]')).toBe('[\n  "a,b:{c}",\n  "q\\"uote"\n]');
  });

  it('keeps empty containers compact and strips stray whitespace', () => {
    expect(prettyJson('{"a":[ ],"b":{ }, "c" : 1}')).toBe('{\n  "a": [],\n  "b": {},\n  "c": 1\n}');
  });

  it('does not touch multi-line or invalid input', () => {
    expect(prettyJson('{\n"a":1}')).toBe('{\n"a":1}');
    expect(prettyJson('{not json}')).toBe('{not json}');
  });
});

describe('markdown renderer', () => {
  it('renders prose and zooms by font size', async () => {
    const { context, container } = renderContext(
      '# Title\n\n![pic](img.png)',
      'readme.md',
      'markdown',
    );
    const instance = instanceOf(await markdownRenderer.render(context));
    const prose = container.querySelector<HTMLElement>('.av-prose')!;
    expect(prose.querySelector('h1')?.textContent).toBe('Title');
    instance.zoom!.set(1.5);
    expect(instance.zoom!.get()).toBe(1.5);
    expect(prose.style.fontSize).toBe('21px');
    expect(instance.printable!()).toBe(prose);
  });

  it('shows very large Markdown files as plain text instead of parsing them', async () => {
    const line = '# heading ';
    const { context, container } = renderContext(line.repeat(160_000), 'huge.md', 'markdown');
    await markdownRenderer.render(context);
    expect(container.querySelector('.av-prose')).toBeNull();
    expect(container.querySelector('.av-code-lines')?.textContent?.startsWith(line + line)).toBe(
      true,
    );
  });

  it('resolves relative images against the file URL', async () => {
    const { context, container } = renderContext('![pic](img/a.png)', 'readme.md', 'markdown');
    Object.assign(context.file, { url: 'https://example.com/docs/readme.md' });
    await markdownRenderer.render(context);
    expect(container.querySelector('img')?.getAttribute('src')).toBe(
      'https://example.com/docs/img/a.png',
    );
  });
});

describe('html renderer', () => {
  it('shows the page in a fully sandboxed frame', async () => {
    const { context, container } = renderContext(
      '<html><head><title>t</title></head><body><script>alert(1)</script>hi</body></html>',
      'page.html',
      'html',
    );
    await htmlRenderer.render(context);
    const frame = container.querySelector('iframe')!;
    expect(frame.getAttribute('sandbox')).toBe('');
    expect(frame.getAttribute('referrerpolicy')).toBe('no-referrer');
    expect(frame.title).toBe('page.html');
    expect(frame.srcdoc).toContain('<head><base href="about:invalid"><title>');
    // Nothing from the file is parsed into the host document.
    expect(container.querySelector('script')).toBeNull();
  });

  it('points relative URLs at the file location, escaping it safely', () => {
    expect(withBase('<p>x</p>', 'https://example.com/a/"><script>.html')).toBe(
      '<base href="https://example.com/a/&quot;>&lt;script>.html"><p>x</p>',
    );
    expect(withBase('<HEAD lang="en"><x></HEAD>', null)).toBe(
      '<HEAD lang="en"><base href="about:invalid"><x></HEAD>',
    );
  });
});

describe('image renderer', () => {
  async function render(name: string, formatId: string, overrides: Partial<RenderContext> = {}) {
    const test = renderContext(PNG_1X1, name, formatId, overrides);
    const pending = Promise.resolve(imageRenderer.render(test.context));
    const image = await waitFor(() => test.container.querySelector('img'));
    return { ...test, pending, image };
  }

  it('waits for the image, then offers zoom', async () => {
    const { pending, image } = await render('photo.png', 'image');
    expect(image.getAttribute('src')).toBe('blob:test/file');
    expect(image.alt).toBe('photo.png');
    image.dispatchEvent(new Event('load'));
    const instance = instanceOf(await pending);
    expect(instance.zoom!.get()).toBe(1);
    expect(image.classList.contains('av-fit')).toBe(true);
    instance.zoom!.set(2);
    expect(image.classList.contains('av-fit')).toBe(false);
    expect(image.style.width).toBe('1600px');
    expect(instance.printable!()).toBe(image);
  });

  it('reports images that cannot be decoded', async () => {
    const { pending, image } = await render('broken.png', 'image');
    image.dispatchEvent(new Event('error'));
    await expect(pending).rejects.toMatchObject({ code: 'RENDER_FAILED' });
  });

  it('stops waiting when the load is aborted', async () => {
    const { pending, abort } = await render('photo.png', 'image');
    abort();
    await expect(pending).rejects.toMatchObject({ code: 'ABORTED' });

    const second = renderContext(PNG_1X1, 'photo.png', 'image');
    second.abort();
    await expect(imageRenderer.render(second.context)).rejects.toMatchObject({ code: 'ABORTED' });
  });

  it('gives SVG blobs their MIME type so they render inside <img>', async () => {
    const { image, objectUrls, pending } = await render('logo.svg', 'svg');
    expect(objectUrls[0]!.type).toBe('image/svg+xml');
    expect(image.getAttribute('src')).toBe('blob:test/1');
    image.dispatchEvent(new Event('load'));
    await pending;
  });

  it('uses the original URL for directly loadable SVG files', async () => {
    const test = renderContext('<svg/>', 'logo.svg', 'svg', {
      fileUrl: () => Promise.resolve('https://example.com/logo.svg'),
    });
    Object.assign(test.context.file, { url: 'https://example.com/logo.svg', direct: true });
    const pending = Promise.resolve(imageRenderer.render(test.context));
    const image = await waitFor(() => test.container.querySelector('img'));
    expect(image.getAttribute('src')).toBe('https://example.com/logo.svg');
    expect(test.objectUrls).toHaveLength(0);
    image.dispatchEvent(new Event('load'));
    await pending;
  });
});

describe('media renderer', () => {
  it('plays video and releases it on destroy', async () => {
    const test = renderContext(new Uint8Array(4), 'clip.mp4', 'video');
    const pending = Promise.resolve(mediaRenderer.render(test.context));
    const player = await waitFor(() => test.container.querySelector('video'));
    await waitFor(() => player.getAttribute('src'));
    expect(player.hasAttribute('controls')).toBe(true);
    expect(player.hasAttribute('playsinline')).toBe(true);
    player.dispatchEvent(new Event('loadedmetadata'));
    const instance = instanceOf(await pending);
    const pause = vi.spyOn(player, 'pause').mockImplementation(() => undefined);
    const load = vi.spyOn(player, 'load').mockImplementation(() => undefined);
    instance.destroy!();
    expect(pause).toHaveBeenCalled();
    expect(load).toHaveBeenCalled();
    expect(player.hasAttribute('src')).toBe(false);
  });

  it('uses an audio element for audio and reports unsupported codecs', async () => {
    const test = renderContext(new Uint8Array(4), 'song.flac', 'audio');
    const pending = Promise.resolve(mediaRenderer.render(test.context));
    const player = await waitFor(() => test.container.querySelector('audio'));
    expect(test.container.querySelector('.av-media-audio')).not.toBeNull();
    await waitFor(() => player.getAttribute('src'));
    player.dispatchEvent(new Event('error'));
    await expect(pending).rejects.toMatchObject({
      code: 'RENDER_FAILED',
      message: expect.stringContaining('audio') as unknown,
    });
  });
});

describe('pdf renderer', () => {
  const PDF = encode('%PDF-1.7\n%%EOF');

  it('embeds valid PDFs with the browser viewer', async () => {
    const { context, container, objectUrls } = renderContext(PDF, 'doc.pdf', 'pdf');
    const instance = instanceOf(await pdfRenderer.render(context));
    const frame = container.querySelector('iframe')!;
    expect(frame.getAttribute('src')).toBe('blob:test/1');
    expect(frame.hasAttribute('sandbox')).toBe(false);
    expect(objectUrls[0]!.type).toBe('application/pdf');

    const opened = vi.fn();
    vi.stubGlobal('open', opened);
    const print = vi.fn();
    Object.defineProperty(frame, 'contentWindow', {
      configurable: true,
      value: { focus: vi.fn(), print },
    });
    instance.print!();
    expect(print).toHaveBeenCalledTimes(1);
    // If in-place printing is impossible the PDF opens in a new tab instead.
    Object.defineProperty(frame, 'contentWindow', { configurable: true, value: null });
    instance.print!();
    expect(opened).toHaveBeenCalledWith('blob:test/1', '_blank', 'noopener');
  });

  it('accepts a header that is not at the very start', async () => {
    const { context, container } = renderContext(encode('junk\n%PDF-1.4'), 'doc.pdf', 'pdf');
    await pdfRenderer.render(context);
    expect(container.querySelector('iframe')).not.toBeNull();
  });

  it('rejects content that is not a PDF', async () => {
    const { context } = renderContext('<html>not a pdf</html>', 'fake.pdf', 'pdf');
    await expect(pdfRenderer.render(context)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
    const tiny = renderContext('ab', 'tiny.pdf', 'pdf');
    await expect(pdfRenderer.render(tiny.context)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('frames a cross-origin PDF directly when it cannot be fetched', async () => {
    const blocked = new AnyViewError('FETCH_FAILED', 'blocked by CORS');
    const { context, container } = renderContext(PDF, 'doc.pdf', 'pdf');
    Object.assign(context.file, {
      url: 'https://other.example/doc.pdf',
      direct: true,
      bytes: () => Promise.reject(blocked),
    });
    const instance = instanceOf(await pdfRenderer.render(context));
    expect(container.querySelector('iframe')?.getAttribute('src')).toBe(
      'https://other.example/doc.pdf',
    );
    const opened = vi.fn();
    vi.stubGlobal('open', opened);
    instance.print!();
    expect(opened).toHaveBeenCalled();
  });

  it('does not fall back for same-origin or non-network failures', async () => {
    const blocked = new AnyViewError('FETCH_FAILED', 'HTTP 404');
    const sameOrigin = renderContext(PDF, 'doc.pdf', 'pdf');
    Object.assign(sameOrigin.context.file, {
      url: `${window.location.origin}/doc.pdf`,
      direct: true,
      bytes: () => Promise.reject(blocked),
    });
    await expect(pdfRenderer.render(sameOrigin.context)).rejects.toBe(blocked);

    const invalidUrl = renderContext(PDF, 'doc.pdf', 'pdf');
    Object.assign(invalidUrl.context.file, {
      url: 'not a url',
      direct: true,
      bytes: () => Promise.reject(blocked),
    });
    await expect(pdfRenderer.render(invalidUrl.context)).rejects.toBe(blocked);
  });

  function fakePdfJs(pageCount: number) {
    const destroy = vi.fn();
    const cancel = vi.fn();
    const rendered: number[] = [];
    const pdfjs: PdfJsLike = {
      getDocument: ({ data }) => ({
        promise: Promise.resolve({
          numPages: pageCount,
          destroy,
          getPage: (pageNumber: number) =>
            Promise.resolve({
              getViewport: ({ scale }: { scale: number }) => ({
                width: 300 * scale,
                height: 400 * scale,
              }),
              render: () => {
                rendered.push(pageNumber);
                return { promise: Promise.resolve(), cancel };
              },
            }),
          byteLength: data.length,
        }),
      }),
    };
    return { pdfjs, destroy, cancel, rendered };
  }

  function withPdfJs(pdfjs: unknown, name = 'doc.pdf') {
    const test = renderContext(PDF, name, 'pdf');
    Object.assign(test.context.options, { pdfjs });
    return test;
  }

  it('draws pages to canvases when PDF.js is supplied', async () => {
    vi.stubGlobal('IntersectionObserver', undefined);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({} as never);
    const { pdfjs, destroy, rendered } = fakePdfJs(3);
    const { context, container, notifications } = withPdfJs(() => Promise.resolve(pdfjs));
    const instance = instanceOf(await pdfRenderer.render(context));

    const holders = container.querySelectorAll<HTMLElement>('.av-page');
    expect(holders).toHaveLength(3);
    expect(holders[0]!.style.width).toBe('400px');
    await waitFor(() => container.querySelectorAll('canvas').length === 3);
    expect(rendered.sort()).toEqual([1, 2, 3]);

    expect(instance.pages!.count()).toBe(3);
    expect(instance.pages!.current()).toBe(1);
    instance.pages!.goTo(2);
    expect(instance.pages!.current()).toBe(2);
    instance.pages!.goTo(9);
    expect(instance.pages!.current()).toBe(2);
    container.querySelector('.av-pages')!.dispatchEvent(new Event('scroll'));
    await waitFor(() => notifications.includes('page'));

    instance.zoom!.set(2);
    expect(instance.zoom!.get()).toBe(2);
    expect(holders[0]!.style.width).toBe('800px');
    await waitFor(() => rendered.length === 6);
    expect(instance.printable!()).toBe(container.querySelector('.av-pages'));

    instance.destroy!();
    expect(destroy).toHaveBeenCalledTimes(1);
  });

  it('draws only pages that scroll into view', async () => {
    const observed: Element[] = [];
    let trigger: (entries: { isIntersecting: boolean; target: Element }[]) => void = () =>
      undefined;
    vi.stubGlobal(
      'IntersectionObserver',
      class {
        constructor(callback: typeof trigger) {
          trigger = callback;
        }
        observe(element: Element): void {
          observed.push(element);
        }
        unobserve(): void {}
        disconnect(): void {}
      },
    );
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue({} as never);
    const { pdfjs, rendered } = fakePdfJs(5);
    const { context, container } = withPdfJs(pdfjs);
    const instance = instanceOf(await pdfRenderer.render(context));
    expect(rendered).toEqual([]);
    trigger([
      { isIntersecting: true, target: observed[1]! },
      { isIntersecting: false, target: observed[2]! },
    ]);
    await waitFor(() => container.querySelectorAll('canvas').length === 1);
    expect(rendered).toEqual([2]);
    instance.zoom!.set(1.5);
    instance.destroy!();
  });

  it('skips drawing when no 2D context is available', async () => {
    vi.stubGlobal('IntersectionObserver', undefined);
    vi.spyOn(HTMLCanvasElement.prototype, 'getContext').mockReturnValue(null);
    const { pdfjs } = fakePdfJs(1);
    const { context, container } = withPdfJs(pdfjs);
    await pdfRenderer.render(context);
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(container.querySelectorAll('canvas')).toHaveLength(0);
  });

  it('reports bad PDF.js modules and unreadable documents', async () => {
    await expect(pdfRenderer.render(withPdfJs({ nope: true }).context)).rejects.toMatchObject({
      code: 'INVALID_OPTIONS',
    });
    await expect(
      pdfRenderer.render(withPdfJs(() => Promise.reject(new Error('chunk failed'))).context),
    ).rejects.toMatchObject({ code: 'RENDER_FAILED' });
    const broken: PdfJsLike = {
      getDocument: () => ({ promise: Promise.reject(new Error('bad xref')) }),
    };
    await expect(pdfRenderer.render(withPdfJs(broken).context)).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('releases the document if the load was aborted meanwhile', async () => {
    const { pdfjs, destroy } = fakePdfJs(1);
    const test = withPdfJs(pdfjs);
    test.abort();
    await expect(pdfRenderer.render(test.context)).rejects.toMatchObject({ code: 'ABORTED' });
    expect(destroy).toHaveBeenCalledTimes(1);
  });

  it('releases through the loading task, as PDF.js 5 and later require', async () => {
    vi.stubGlobal('IntersectionObserver', undefined);
    const { pdfjs } = fakePdfJs(1);
    const taskDestroy = vi.fn(() => Promise.reject(new Error('already destroyed')));
    const modern: PdfJsLike = {
      getDocument: (source) => {
        const task = pdfjs.getDocument(source);
        // Modern documents have no destroy() of their own.
        return {
          promise: task.promise.then((document) => ({ ...document, destroy: undefined })),
          destroy: taskDestroy,
        };
      },
    };
    const { context } = withPdfJs(modern);
    const instance = instanceOf(await pdfRenderer.render(context));
    instance.destroy!();
    expect(taskDestroy).toHaveBeenCalledTimes(1);
  });
});

describe('font renderer', () => {
  it('registers the font, shows a specimen and unregisters on destroy', async () => {
    const added = new Set<unknown>();
    const load = vi.fn(() => Promise.resolve());
    class FakeFontFace {
      constructor(
        readonly family: string,
        readonly source: ArrayBuffer,
      ) {}
      load = load;
    }
    vi.stubGlobal('FontFace', FakeFontFace);
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: {
        add: (face: unknown) => added.add(face),
        delete: (face: unknown) => added.delete(face),
      },
    });
    const { context, container } = renderContext(new Uint8Array(8), 'type.woff2', 'font');
    const instance = instanceOf(await fontRenderer.render(context));
    expect(added.size).toBe(1);
    const lines = container.querySelectorAll<HTMLElement>('.av-font p');
    expect(lines).toHaveLength(9);
    expect(lines[0]!.style.fontFamily).toMatch(/^anyview-font-\d+$/);
    expect(lines[8]!.textContent).toBe('72pxThe quick brown fox jumps over the lazy dog');
    expect(instance.printable!()).toBe(container.querySelector('.av-font'));
    instance.destroy!();
    expect(added.size).toBe(0);
  });

  it('reports unreadable fonts and browsers without the font API', async () => {
    class BrokenFontFace {
      load = (): Promise<void> => Promise.reject(new Error('OTS parsing error'));
    }
    vi.stubGlobal('FontFace', BrokenFontFace);
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: { add: vi.fn(), delete: vi.fn() },
    });
    const broken = renderContext(new Uint8Array(8), 'bad.ttf', 'font');
    await expect(fontRenderer.render(broken.context)).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });

    vi.stubGlobal('FontFace', undefined);
    const unsupported = renderContext(new Uint8Array(8), 'type.ttf', 'font');
    await expect(fontRenderer.render(unsupported.context)).rejects.toMatchObject({
      code: 'RENDER_FAILED',
    });
  });
});

describe('EPUB', () => {
  it('reads the title and reading order, skipping missing chapters', async () => {
    const book = await openEpub(openZip(epubFixture()));
    expect(book.title).toBe('The Test Book');
    expect(book.chapterCount).toBe(2);
  });

  it('rejects packages that are not EPUBs', async () => {
    await expect(openEpub(openZip(makeZip({ 'a.txt': 'x' })))).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
    const noPackage = makeZip({
      'META-INF/container.xml':
        '<container><rootfiles><rootfile full-path="missing.opf"/></rootfiles></container>',
    });
    await expect(openEpub(openZip(noPackage))).rejects.toMatchObject({ code: 'PARSE_FAILED' });
    const noChapters = makeZip({
      'META-INF/container.xml':
        '<container><rootfiles><rootfile full-path="c.opf"/></rootfiles></container>',
      'c.opf': '<package><manifest/><spine/></package>',
    });
    await expect(openEpub(openZip(noChapters))).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('produces self-contained chapters with all active content removed', async () => {
    const book = await openEpub(openZip(epubFixture()));
    const html = await book.chapterHtml(0);
    const parsed = new DOMParser().parseFromString(html, 'text/html');
    expect(parsed.querySelector('h1')?.textContent).toBe('Chapter One');
    for (const selector of ['script', 'iframe', 'form', 'base', 'meta[http-equiv]', 'link']) {
      expect(parsed.querySelector(selector), selector).toBeNull();
    }
    expect(html).not.toMatch(/onload|onclick|alert\(/);
    // Links would navigate the frame away from the book, so they lose their target.
    expect(parsed.querySelectorAll('a[href]')).toHaveLength(0);
    expect(parsed.querySelector('p')?.textContent).toContain('Hello next out');

    const images = parsed.querySelectorAll('img');
    expect(images[0]!.getAttribute('src')).toMatch(/^data:image\/png;base64,iVBOR/);
    expect(images[0]!.hasAttribute('srcset')).toBe(false);
    expect(images[1]!.hasAttribute('src')).toBe(false);
    expect(images[2]!.hasAttribute('src')).toBe(false);
    const vectorImages = parsed.querySelectorAll('svg image');
    expect(vectorImages[0]!.getAttribute('href')).toMatch(/^data:image\/png;base64,/);
    expect(vectorImages[1]!.hasAttribute('href')).toBe(false);

    const styles = Array.from(parsed.querySelectorAll('style'), (style) => style.textContent ?? '');
    expect(styles.some((css) => css.includes('h1 { color: navy }'))).toBe(true);
    expect(styles.join('')).not.toContain('@import');
  });

  it('applies zoom and reports missing chapters', async () => {
    const book = await openEpub(openZip(epubFixture()));
    expect(await book.chapterHtml(1, 1.5)).toContain('body{zoom:1.5}');
    await expect(book.chapterHtml(7)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('expands XHTML self-closing tags that HTML would leave open', () => {
    expect(expandSelfClosing('<title/><a id="x"/><br/><img src="a"/><svg:path d="M0"/>')).toBe(
      '<title></title><a id="x"></a><br/><img src="a"/><svg:path d="M0"></svg:path>',
    );
  });

  it('pages through chapters in a sandboxed frame', async () => {
    const { context, container } = renderContext(epubFixture(), 'book.epub', 'epub');
    const instance = instanceOf(await epubRenderer.render(context));
    const frame = (): HTMLIFrameElement => container.querySelector('iframe')!;
    expect(frame().getAttribute('sandbox')).toBe('');
    expect(frame().title).toBe('The Test Book');
    expect(frame().srcdoc).toContain('Chapter One');
    expect(instance.pages!.count()).toBe(2);
    expect(instance.pages!.current()).toBe(1);

    instance.pages!.goTo(2);
    await waitFor(() => frame().srcdoc.includes('Second chapter'));
    expect(instance.pages!.current()).toBe(2);
    instance.zoom!.set(2);
    expect(instance.zoom!.get()).toBe(2);
    await waitFor(() => frame().srcdoc.includes('body{zoom:2}'));

    instance.pages!.goTo(9);
    await waitFor(() => container.textContent?.includes('missing from this EPUB'));
    instance.destroy!();
  });
});

describe('archive renderer', () => {
  it('lists files by name and opens them in a nested viewer', async () => {
    const zip = makeZip({ 'docs/': '', 'docs/b.txt': 'bee', 'a.md': '# A' });
    const { context, container, children } = renderContext(zip, 'files.zip', 'zip');
    const instance = instanceOf(await archiveRenderer.render(context));
    const rows = container.querySelectorAll('tbody tr');
    expect(Array.from(rows, (row) => row.textContent)).toEqual(['a.md3 B', 'docs/b.txt3 B']);
    expect(instance.printable!()).toBe(container.querySelector('table'));

    rows[1]!.querySelector('button')!.click();
    await waitFor(() => children.length === 1);
    expect(children[0]!.name).toBe('b.txt');
    expect(await children[0]!.source.text()).toBe('bee');
  });

  it('explains when an entry cannot be extracted', async () => {
    const { context, container, children } = renderContext(
      makeZip({ 'secret.txt': 'x' }, { encrypted: true }),
      'locked.zip',
      'zip',
    );
    await archiveRenderer.render(context);
    container.querySelector<HTMLButtonElement>('tbody button')!.click();
    const status = await waitFor(() => {
      const element = container.querySelector<HTMLElement>('[role=alert]');
      return element && !element.hidden ? element : null;
    });
    expect(status.textContent).toContain('encrypted');
    expect(children).toHaveLength(0);
  });

  it('does not open an entry once the archive has been closed', async () => {
    const test = renderContext(makeZip({ 'a.txt': 'x' }), 'files.zip', 'zip');
    await archiveRenderer.render(test.context);
    test.abort();
    test.container.querySelector<HTMLButtonElement>('tbody button')!.click();
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(test.children).toHaveLength(0);
  });

  it('says so when the archive is empty', async () => {
    const { context, container } = renderContext(makeZip({}), 'empty.zip', 'zip');
    await archiveRenderer.render(context);
    expect(container.textContent).toBe('This archive is empty.');
  });

  it('truncates enormous listings', async () => {
    const files: Record<string, string> = {};
    for (let index = 0; index < 5005; index++) files[`f${index}.txt`] = '';
    const { context, container } = renderContext(
      makeZip(files, { compress: false }),
      'many.zip',
      'zip',
    );
    await archiveRenderer.render(context);
    expect(container.querySelector('.av-notice')?.textContent).toContain('5,005');
    expect(container.querySelectorAll('tbody tr')).toHaveLength(5000);
  });

  it('rejects files that are not ZIP archives', async () => {
    const { context } = renderContext('not a zip at all, just some text', 'fake.zip', 'zip');
    await expect(archiveRenderer.render(context)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });
});

describe('notebook renderer', () => {
  it('renders markdown, code and outputs', async () => {
    const { context, container } = renderContext(NOTEBOOK_FIXTURE, 'analysis.ipynb', 'ipynb');
    const instance = instanceOf(await notebookRenderer.render(context));
    expect(container.querySelector('h1')?.textContent).toBe('Analysis');
    expect(container.querySelector('em')?.textContent).toBe('notes');
    const code = container.querySelector('.av-nb-code')!;
    expect(code.textContent).toBe('import math\nprint(math.pi)');
    expect(code.querySelector('.av-tok-keyword')?.textContent).toBe('import');

    const outputs = Array.from(
      container.querySelectorAll('.av-nb-output'),
      (output) => output.textContent,
    );
    expect(outputs).toEqual([
      '3.14159\n',
      'warning',
      '42',
      'Traceback\nline 2',
      'KeyError: k',
      'raw cell text',
    ]);
    expect(container.querySelectorAll('.av-nb-error')).toHaveLength(3);
    expect(container.querySelector('img')?.getAttribute('src')).toMatch(
      /^data:image\/png;base64,iVBOR/,
    );
    expect(container.querySelector('strong')?.textContent).toBe('rendered');
    // HTML output is never interpreted.
    expect(container.querySelector('script')).toBeNull();
    expect(container.querySelector('b')).toBeNull();

    instance.zoom!.set(1.25);
    expect(instance.zoom!.get()).toBe(1.25);
    expect(instance.printable!()).toBe(container.querySelector('.av-prose'));
  });

  it('supports the older worksheet layout', async () => {
    const legacy = JSON.stringify({
      nbformat: 3,
      worksheets: [
        { cells: [{ cell_type: 'code', input: ['x = 1'], language: 'python' }] },
        'junk',
      ],
    });
    const { context, container } = renderContext(legacy, 'old.ipynb', 'ipynb');
    await notebookRenderer.render(context);
    expect(container.querySelector('.av-nb-code')?.textContent).toBe('x = 1');
  });

  it.each([
    ['invalid JSON', '{nope'],
    ['a non-object', '[1, 2]'],
    ['an object without cells', '{"metadata": {}}'],
  ])('rejects %s', async (_label, source) => {
    const { context } = renderContext(source, 'bad.ipynb', 'ipynb');
    await expect(notebookRenderer.render(context)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });
});

describe('email', () => {
  it('decodes encoded-word headers', () => {
    expect(decodeHeaderValue('=?UTF-8?B?SsO8cmdlbg==?= <j@example.com>')).toBe(
      'Jürgen <j@example.com>',
    );
    expect(decodeHeaderValue('=?iso-8859-1?Q?Andr=E9_Du_Pont?=')).toBe('André Du Pont');
    expect(decodeHeaderValue('=?UTF-8?Q?one?=   =?UTF-8?Q?_two?=')).toBe('one two');
    expect(decodeHeaderValue('plain')).toBe('plain');
  });

  it('parses headers, alternative bodies and attachments', () => {
    const message = parseEmail(encode(EML_FIXTURE));
    expect(message.headers.get('from')).toBe('Jürgen <jurgen@example.com>');
    expect(message.headers.get('to')).toBe('André Du Pont <andre@example.com>, second@example.com');
    expect(message.headers.get('subject')).toBe('Quarterly report ✓');
    expect(message.text).toBe('Plain café body with a soft line break.');
    expect(message.html).toBe('<p>HTML <b>body</b><script>alert(1)</script></p>');
    expect(message.attachments.map((attachment) => [attachment.name, attachment.mimeType])).toEqual(
      [
        ['data.csv', 'text/csv'],
        ['résumé.txt', 'application/octet-stream'],
        ['attachment-3.png', 'image/png'],
        ['attachment-4.eml', 'message/rfc822'],
        ['long-name.txt', 'text/plain'],
      ],
    );
    expect(new TextDecoder().decode(message.attachments[0]!.data)).toBe('a,b\n1,2');
    expect(Array.from(message.attachments[2]!.data)).toEqual(Array.from(PNG_1X1));
    expect(new TextDecoder().decode(message.attachments[3]!.data)).toBe(
      'Subject: nested\r\n\r\ninner message',
    );
  });

  it('parses a plain single-part message and raw 8-bit headers', () => {
    const message = parseEmail(encode('Subject: Größe\nX-Odd line without colon\n\nHello\nWorld'));
    expect(message.headers.get('subject')).toBe('Größe');
    expect(message.text).toBe('Hello\nWorld');
    expect(message.html).toBeNull();
    expect(message.attachments).toEqual([]);
    expect(parseEmail(encode('Subject: headers only')).text).toBe('');
  });

  it('treats multipart messages without a boundary as plain content', () => {
    const message = parseEmail(encode('Content-Type: multipart/mixed\n\nbody'));
    expect(message.attachments).toHaveLength(1);
  });

  it('shows headers, a sandboxed HTML body and attachments', async () => {
    const { context, container, children } = renderContext(EML_FIXTURE, 'mail.eml', 'eml');
    const instance = instanceOf(await emailRenderer.render(context));
    const head = container.querySelector('.av-mail-head')!;
    expect(Array.from(head.querySelectorAll('dt'), (term) => term.textContent)).toEqual([
      'Subject',
      'From',
      'To',
      'Cc',
      'Date',
    ]);
    const frame = container.querySelector('iframe')!;
    expect(frame.getAttribute('sandbox')).toBe('');
    expect(frame.srcdoc).toContain('<p>HTML <b>body</b>');
    expect(container.querySelector('script')).toBeNull();

    const attachments = container.querySelectorAll<HTMLButtonElement>(
      '.av-mail-attachments button',
    );
    expect(attachments).toHaveLength(5);
    expect(attachments[0]!.textContent).toBe('data.csv (7 B)');
    attachments[0]!.click();
    expect(children[0]).toMatchObject({ name: 'data.csv' });
    expect(children[0]!.source.type).toBe('text/csv');
    expect(instance.printable!()).toBe(container.querySelector('.av-mail'));
  });

  it('shows the text body when there is no HTML', async () => {
    const { context, container } = renderContext('Subject: hi\n\nJust text', 'mail.eml', 'eml');
    await emailRenderer.render(context);
    expect(container.querySelector('.av-mail-text')?.textContent).toBe('Just text');
    expect(container.querySelector('iframe')).toBeNull();
    expect(container.querySelector('.av-mail-attachments')).toBeNull();
  });
});

describe('printNode', () => {
  function fakeFrameWindow() {
    const frameDocument = document.implementation.createHTMLDocument('');
    const listeners: Record<string, () => void> = {};
    const view = {
      focus: vi.fn(),
      print: vi.fn(),
      addEventListener: (type: string, listener: () => void) => void (listeners[type] = listener),
    };
    vi.spyOn(HTMLIFrameElement.prototype, 'contentDocument', 'get').mockReturnValue(frameDocument);
    vi.spyOn(HTMLIFrameElement.prototype, 'contentWindow', 'get').mockReturnValue(view as never);
    return { frameDocument, view, listeners };
  }

  it('prints a copy of the node from a hidden frame and cleans up', () => {
    const { frameDocument, view, listeners } = fakeFrameWindow();
    const node = document.createElement('div');
    node.textContent = 'printable';
    printNode(node, 'My file');
    expect(document.querySelector('iframe[aria-hidden]')).not.toBeNull();
    expect(frameDocument.title).toBe('My file');
    expect(frameDocument.querySelector('.av-root')?.textContent).toBe('printable');
    expect(frameDocument.querySelector('style')?.textContent).toContain('.av-root');
    expect(view.print).toHaveBeenCalledTimes(1);
    listeners['afterprint']!();
    listeners['afterprint']!();
    expect(document.querySelector('iframe[aria-hidden]')).toBeNull();
  });

  it('waits for images and replaces canvases with snapshots', () => {
    vi.useFakeTimers();
    try {
      const { frameDocument, view } = fakeFrameWindow();
      vi.spyOn(HTMLImageElement.prototype, 'complete', 'get').mockReturnValue(false);
      vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockReturnValue(
        'data:image/png;base64,AAAA',
      );
      const node = document.createElement('div');
      node.append(document.createElement('img'), document.createElement('canvas'));
      printNode(node, 'x');
      expect(frameDocument.querySelector('canvas')).toBeNull();
      expect(frameDocument.querySelectorAll('img')).toHaveLength(2);
      expect(view.print).not.toHaveBeenCalled();
      const [first, second] = Array.from(frameDocument.querySelectorAll('img'));
      first!.dispatchEvent(new Event('load'));
      expect(view.print).not.toHaveBeenCalled();
      second!.dispatchEvent(new Event('error'));
      expect(view.print).toHaveBeenCalledTimes(1);
      vi.advanceTimersByTime(61_000);
      expect(view.print).toHaveBeenCalledTimes(1);
      expect(document.querySelector('iframe[aria-hidden]')).toBeNull();
    } finally {
      vi.useRealTimers();
    }
  });

  it('prints after a timeout if an image never settles, and keeps tainted canvases', () => {
    vi.useFakeTimers();
    try {
      const { frameDocument, view } = fakeFrameWindow();
      vi.spyOn(HTMLImageElement.prototype, 'complete', 'get').mockReturnValue(false);
      vi.spyOn(HTMLCanvasElement.prototype, 'toDataURL').mockImplementation(() => {
        throw new DOMException('tainted', 'SecurityError');
      });
      const node = document.createElement('div');
      node.append(document.createElement('img'), document.createElement('canvas'));
      printNode(node, 'x');
      expect(frameDocument.querySelector('canvas')).not.toBeNull();
      vi.advanceTimersByTime(1500);
      expect(view.print).toHaveBeenCalledTimes(1);
      frameDocument.querySelector('img')!.dispatchEvent(new Event('error'));
      expect(view.print).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('falls back to printing the page when the frame is unavailable', () => {
    vi.spyOn(HTMLIFrameElement.prototype, 'contentDocument', 'get').mockReturnValue(null);
    const print = vi.fn();
    vi.stubGlobal('print', print);
    printNode(document.createTextNode('x'), 'x');
    expect(print).toHaveBeenCalledTimes(1);
    expect(document.querySelector('iframe')).toBeNull();
  });
});
