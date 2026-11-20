import { afterEach, describe, expect, it, vi } from 'vitest';
import { AnyViewError, createViewer, isAnyViewError, VERSION } from '../../src';
import type {
  RenderContext,
  Renderer,
  RendererDefinition,
  RendererInstance,
  ViewerOptions,
} from '../../src';
import { mockFetch, mount, textFile, waitFor } from '../helpers/dom';

/** A custom renderer whose behaviour each test controls. */
function fakeRenderer(
  build: (context: RenderContext) => RendererInstance | void | Promise<RendererInstance | void>,
  definition: Partial<RendererDefinition> = {},
): RendererDefinition {
  const renderer: Renderer = { render: build };
  return { id: 'fake', extensions: ['fake'], load: () => renderer, ...definition };
}

const fakeFile = (): File => textFile('content', 'thing.fake');

function codeOf(run: () => unknown): string {
  try {
    run();
  } catch (error) {
    return isAnyViewError(error) ? error.code : `not an AnyViewError: ${String(error)}`;
  }
  return 'no error';
}

afterEach(() => {
  vi.unstubAllGlobals();
  vi.restoreAllMocks();
  document.body.textContent = '';
});

describe('createViewer: setup', () => {
  it('exposes the package version', () => {
    expect(VERSION).toMatch(/^\d+\.\d+\.\d+/);
  });

  it('accepts an element or a selector', () => {
    const container = mount();
    container.id = 'target';
    const viewer = createViewer('#target');
    expect(container.querySelector('.av-root')).toBe(viewer.element);
    expect(viewer.state).toBe('idle');
    expect(viewer.format).toBeNull();
  });

  it('rejects missing or invalid containers', () => {
    expect(codeOf(() => createViewer('#does-not-exist'))).toBe('INVALID_OPTIONS');
    expect(codeOf(() => createViewer(null as never))).toBe('INVALID_OPTIONS');
    expect(codeOf(() => createViewer(document.createTextNode('x') as never))).toBe(
      'INVALID_OPTIONS',
    );
  });

  it.each<[string, unknown]>([
    ['a non-object', 'nope'],
    ['an unknown theme', { theme: 'neon' }],
    ['a negative size limit', { maxFileSize: -1 }],
    ['a non-numeric size limit', { maxFileSize: '10' }],
    ['renderers that are not an array', { renderers: {} }],
    ['gridlines that are not a boolean', { tableGridlines: 'yes' }],
    ['a renderer without load()', { renderers: [{ id: 'x' }] }],
    ['a renderer without an id', { renderers: [{ load: () => ({}) }] }],
  ])('rejects %s', (_label, options) => {
    expect(codeOf(() => createViewer(mount(), options as ViewerOptions))).toBe('INVALID_OPTIONS');
  });

  it('injects its stylesheet once, with the given nonce', () => {
    document.getElementById('anyview-styles')?.remove();
    createViewer(mount(), { styleNonce: 'abc' });
    createViewer(mount());
    const styles = document.querySelectorAll('#anyview-styles');
    expect(styles).toHaveLength(1);
    expect(styles[0]!.getAttribute('nonce')).toBe('abc');
  });

  it('shows the empty state, theme and class name', () => {
    const container = mount();
    const viewer = createViewer(container, {
      theme: 'dark',
      className: 'custom',
      labels: { empty: 'Nothing here' },
    });
    expect(viewer.element.getAttribute('data-theme')).toBe('dark');
    expect(viewer.element.classList.contains('custom')).toBe(true);
    expect(container.querySelector('.av-status')?.textContent).toBe('Nothing here');
  });

  it('omits the toolbar when asked', () => {
    const container = mount();
    createViewer(container, { toolbar: false });
    expect(container.querySelector('.av-toolbar')).toBeNull();
  });

  it('hides individual toolbar controls', async () => {
    const container = mount();
    const viewer = createViewer(container, { toolbar: { download: false, print: false } });
    await viewer.load(textFile('hello', 'a.txt'));
    const hidden = (label: string): boolean =>
      container.querySelector<HTMLElement>(`.av-toolbar [aria-label="${label}"]`)!.hidden === true;
    expect(hidden('Download')).toBe(true);
    expect(hidden('Print')).toBe(true);
    expect(hidden('Zoom in')).toBe(false);
  });

  it('changes theme at runtime and validates it', () => {
    const viewer = createViewer(mount());
    viewer.setTheme('light');
    expect(viewer.element.getAttribute('data-theme')).toBe('light');
    expect(codeOf(() => viewer.setTheme('pink' as never))).toBe('INVALID_OPTIONS');
  });
});

describe('createViewer: loading', () => {
  it('loads a file and reports it', async () => {
    const container = mount();
    const viewer = createViewer(container);
    const events: string[] = [];
    viewer.on('loadstart', () => events.push('loadstart'));
    const loaded = vi.fn();
    viewer.on('load', (detail) => {
      events.push('load');
      loaded(detail);
    });
    const pending = viewer.load(textFile('const a = 1;', 'app.js'));
    expect(viewer.state).toBe('loading');
    expect(container.querySelector('.av-spinner')).not.toBeNull();
    await pending;

    expect(viewer.state).toBe('ready');
    expect(viewer.format?.id).toBe('text');
    expect(events).toEqual(['loadstart', 'load']);
    expect(loaded).toHaveBeenCalledWith({
      format: expect.objectContaining({ id: 'text' }) as unknown,
      file: { name: 'app.js', size: 12, mimeType: '' },
    });
    expect(container.querySelector('.av-content')?.textContent).toContain('const a = 1;');
    expect(container.querySelector<HTMLElement>('.av-status')!.hidden).toBe(true);
    expect(container.querySelector('.av-file-name')?.textContent).toBe('app.js');
    expect(container.querySelector('.av-file-kind')?.textContent).toBe('Text');
  });

  it('opens the src option after listeners can subscribe', async () => {
    const viewer = createViewer(mount(), { src: textFile('# Hi', 'readme.md') });
    const loaded = vi.fn();
    viewer.on('load', loaded);
    await waitFor(() => viewer.state === 'ready');
    expect(loaded).toHaveBeenCalledTimes(1);
    expect(viewer.format?.id).toBe('markdown');
  });

  it('reports a failing src option through the error event only', async () => {
    const viewer = createViewer(mount(), { src: 42 as never });
    const failed = vi.fn();
    viewer.on('error', failed);
    await waitFor(() => viewer.state === 'error');
    expect(failed.mock.calls[0]![0]).toMatchObject({ code: 'INVALID_SOURCE' });
  });

  it('skips the initial src when load() is called first', async () => {
    const viewer = createViewer(mount(), { src: textFile('one', 'one.txt') });
    await viewer.load(textFile('two', 'two.txt'));
    await new Promise((resolve) => setTimeout(resolve, 20));
    expect(viewer.element.querySelector('.av-file-name')?.textContent).toBe('two.txt');
  });

  it('honours name and type overrides', async () => {
    const viewer = createViewer(mount(), { name: 'data.csv' });
    await viewer.load(new Blob(['a,b\n1,2']));
    expect(viewer.format?.id).toBe('csv');
    await viewer.load(new Blob(['x: 1']), { name: 'settings', type: 'text' });
    expect(viewer.format?.id).toBe('text');
    expect(viewer.element.querySelector('.av-file-name')?.textContent).toBe('settings');
  });

  it('detects the format of unnamed data from its content', async () => {
    const viewer = createViewer(mount());
    await viewer.load(new TextEncoder().encode('{"a":1}'));
    expect(viewer.format?.id).toBe('json');
  });

  it('downloads an extension-less URL to learn its type', async () => {
    mockFetch(() => new Response('a,b\n1,2', { headers: { 'content-type': 'text/csv' } }));
    const viewer = createViewer(mount());
    await viewer.load('https://example.com/export?id=1');
    expect(viewer.format?.id).toBe('csv');
  });

  it('shows the fallback for recognised but unsupported formats', async () => {
    const container = mount();
    const viewer = createViewer(container);
    const failed = vi.fn();
    viewer.on('error', failed);
    await expect(viewer.load(new File(['x'], 'legacy.doc'))).rejects.toMatchObject({
      code: 'UNSUPPORTED_FORMAT',
      details: { format: 'doc', name: 'legacy.doc' },
    });
    expect(viewer.state).toBe('error');
    expect(failed).toHaveBeenCalledTimes(1);
    const status = container.querySelector('.av-status')!;
    expect(status.getAttribute('role')).toBe('alert');
    expect(status.textContent).toContain('Preview is not available');
    expect(status.textContent).toContain('Word 97–2003 document');
    expect(status.querySelector('button')?.textContent).toBe('Download');
  });

  it('describes completely unknown files generically', async () => {
    const viewer = createViewer(mount());
    await expect(viewer.load(new Uint8Array([1, 0, 2, 0, 3]))).rejects.toMatchObject({
      code: 'UNSUPPORTED_FORMAT',
      message: 'AnyView cannot display this type of file.',
    });
  });

  it('reports download failures without offering a download', async () => {
    mockFetch(() => new Response('', { status: 500 }));
    const container = mount();
    const viewer = createViewer(container);
    await expect(viewer.load('https://example.com/a.txt')).rejects.toMatchObject({
      code: 'FETCH_FAILED',
    });
    expect(container.querySelector('.av-status button')).toBeNull();
    expect(container.querySelector('.av-status')?.textContent).toContain('HTTP 500');
  });

  it('wraps unexpected renderer errors', async () => {
    const viewer = createViewer(mount(), {
      renderers: [
        fakeRenderer(() => {
          throw new TypeError('boom');
        }),
      ],
    });
    const error = await viewer.load(fakeFile()).catch((caught: unknown) => caught);
    expect(error).toBeInstanceOf(AnyViewError);
    expect(error).toMatchObject({
      code: 'RENDER_FAILED',
      message: expect.stringContaining('boom') as unknown,
    });
    expect(viewer.element.querySelector('.av-content')?.childNodes).toHaveLength(0);
  });

  it('recovers from an error on the next load', async () => {
    const viewer = createViewer(mount());
    await viewer.load(new File(['x'], 'legacy.doc')).catch(() => undefined);
    await viewer.load(textFile('ok', 'a.txt'));
    expect(viewer.state).toBe('ready');
    expect(viewer.element.querySelector<HTMLElement>('.av-status')!.hidden).toBe(true);
  });

  it('lets a newer load supersede an older one', async () => {
    const started: string[] = [];
    const viewer = createViewer(mount(), {
      renderers: [
        fakeRenderer(async (context) => {
          started.push(context.file.name);
          await new Promise((resolve) =>
            setTimeout(resolve, context.file.name === 'slow.fake' ? 40 : 0),
          );
          context.container.textContent = context.file.name;
        }),
      ],
    });
    const loaded = vi.fn();
    viewer.on('load', loaded);
    const first = viewer.load(textFile('1', 'slow.fake'));
    const second = viewer.load(textFile('2', 'fast.fake'));
    await expect(first).resolves.toBeUndefined();
    await second;
    await new Promise((resolve) => setTimeout(resolve, 60));
    expect(loaded).toHaveBeenCalledTimes(1);
    expect(loaded.mock.calls[0]![0]).toMatchObject({ file: { name: 'fast.fake' } });
    expect(viewer.element.querySelector('.av-file-name')?.textContent).toBe('fast.fake');
  });

  it('destroys an instance that finishes rendering after being superseded', async () => {
    const destroyed = vi.fn();
    let release: () => void = () => undefined;
    const viewer = createViewer(mount(), {
      renderers: [
        fakeRenderer(async (context) => {
          if (context.file.name === 'first.fake') {
            await new Promise<void>((resolve) => (release = resolve));
          }
          return { destroy: destroyed };
        }),
      ],
    });
    const first = viewer.load(textFile('1', 'first.fake'));
    await waitFor(() => release !== undefined && viewer.state === 'loading');
    await new Promise((resolve) => setTimeout(resolve, 10));
    await viewer.load(textFile('2', 'second.fake'));
    release();
    await first;
    expect(destroyed).toHaveBeenCalledTimes(1);
    expect(viewer.state).toBe('ready');
  });

  it('clears back to the empty state', async () => {
    const container = mount();
    const viewer = createViewer(container);
    await viewer.load(textFile('hello', 'a.txt'));
    viewer.clear();
    expect(viewer.state).toBe('idle');
    expect(viewer.format).toBeNull();
    expect(container.querySelector('.av-content')?.childNodes).toHaveLength(0);
    expect(container.querySelector('.av-status')?.textContent).toBe('No file selected');
    expect(container.querySelector('.av-file-name')?.textContent).toBe('');
  });
});

describe('createViewer: lifecycle', () => {
  it('tears everything down on destroy', async () => {
    const container = mount();
    const destroyed = vi.fn();
    const revoke = vi.spyOn(URL, 'revokeObjectURL');
    const viewer = createViewer(container, {
      renderers: [
        fakeRenderer((context) => {
          context.objectUrl(new Blob(['x']));
          return { destroy: destroyed };
        }),
      ],
    });
    const onDestroy = vi.fn();
    viewer.on('destroy', onDestroy);
    await viewer.load(fakeFile());
    viewer.destroy();
    viewer.destroy();

    expect(destroyed).toHaveBeenCalledTimes(1);
    expect(revoke).toHaveBeenCalledTimes(1);
    expect(onDestroy).toHaveBeenCalledTimes(1);
    expect(container.querySelector('.av-root')).toBeNull();
    expect(viewer.state).toBe('destroyed');
    await expect(viewer.load(fakeFile())).rejects.toMatchObject({ code: 'DESTROYED' });
    expect(codeOf(() => viewer.clear())).toBe('DESTROYED');
  });

  it('resolves quietly when destroyed mid-load', async () => {
    const viewer = createViewer(mount(), {
      renderers: [fakeRenderer(() => new Promise((resolve) => setTimeout(resolve, 20)))],
    });
    const failed = vi.fn();
    const loaded = vi.fn();
    viewer.on('error', failed);
    viewer.on('load', loaded);
    const pending = viewer.load(fakeFile());
    viewer.destroy();
    await expect(pending).resolves.toBeUndefined();
    expect(failed).not.toHaveBeenCalled();
    expect(loaded).not.toHaveBeenCalled();
  });

  it('aborts the signal given to renderers when the file is replaced', async () => {
    let signal: AbortSignal | null = null;
    const viewer = createViewer(mount(), {
      renderers: [fakeRenderer((context) => void (signal = context.signal))],
    });
    await viewer.load(fakeFile());
    expect(signal!.aborted).toBe(false);
    viewer.clear();
    expect(signal!.aborted).toBe(true);
  });

  it('keeps working when a renderer destroy() or an event handler throws', async () => {
    const reported = vi.fn();
    vi.stubGlobal('reportError', reported);
    const viewer = createViewer(mount(), {
      renderers: [
        fakeRenderer(() => ({
          destroy: () => {
            throw new Error('destroy failed');
          },
        })),
      ],
    });
    const second = vi.fn();
    viewer.on('load', () => {
      throw new Error('handler failed');
    });
    viewer.on('load', second);
    await viewer.load(fakeFile());
    expect(second).toHaveBeenCalledTimes(1);
    viewer.clear();
    expect(reported).toHaveBeenCalledTimes(2);
    expect(viewer.state).toBe('idle');
  });

  it('falls back to console.error when reportError is unavailable', async () => {
    vi.stubGlobal('reportError', undefined);
    const logged = vi.spyOn(console, 'error').mockImplementation(() => undefined);
    const viewer = createViewer(mount());
    viewer.on('load', () => {
      throw new Error('handler failed');
    });
    await viewer.load(textFile('x', 'a.txt'));
    expect(logged).toHaveBeenCalledTimes(1);
  });

  it('unsubscribes with the returned function or off()', async () => {
    const viewer = createViewer(mount());
    const first = vi.fn();
    const second = vi.fn();
    const unsubscribe = viewer.on('load', first);
    viewer.on('load', second);
    unsubscribe();
    viewer.off('load', second);
    viewer.off('error', second);
    await viewer.load(textFile('x', 'a.txt'));
    expect(first).not.toHaveBeenCalled();
    expect(second).not.toHaveBeenCalled();
  });
});

describe('createViewer: zoom and pages', () => {
  function withCapabilities(pageCount = 5) {
    let zoom = 1;
    let page = 1;
    let context: RenderContext | null = null;
    const viewer = createViewer(mount(), {
      renderers: [
        fakeRenderer((received) => {
          context = received;
          return {
            zoom: { get: () => zoom, set: (value) => void (zoom = value) },
            pages: {
              count: () => pageCount,
              current: () => page,
              goTo: (value) => void (page = value),
            },
          };
        }),
      ],
    });
    return {
      viewer,
      context: () => context!,
      setPage: (value: number) => void (page = value),
      setZoom: (value: number) => void (zoom = value),
    };
  }

  it('reports null for files without zoom or pages', async () => {
    const viewer = createViewer(mount(), { renderers: [fakeRenderer(() => undefined)] });
    expect(viewer.getZoom()).toBeNull();
    await viewer.load(fakeFile());
    expect(viewer.getZoom()).toBeNull();
    expect(viewer.getPage()).toBeNull();
    viewer.setZoom(2);
    viewer.zoomIn();
    viewer.goToPage(3);
    const toolbar = viewer.element.querySelector('.av-toolbar')!;
    expect(
      toolbar.querySelector<HTMLElement>('[aria-label="Zoom in"]')!.parentElement!.hidden,
    ).toBe(true);
    expect(toolbar.querySelector<HTMLElement>('[aria-label="Next"]')!.parentElement!.hidden).toBe(
      true,
    );
  });

  it('steps, clamps and resets zoom', async () => {
    const { viewer } = withCapabilities();
    const changes: number[] = [];
    viewer.on('zoomchange', ({ zoom }) => changes.push(zoom));
    await viewer.load(fakeFile());
    viewer.zoomIn();
    viewer.zoomIn();
    viewer.zoomOut();
    viewer.setZoom(999);
    viewer.setZoom(0);
    viewer.setZoom(Number.NaN);
    viewer.setZoom(0.1);
    expect(changes).toEqual([1.1, 1.25, 1.1, 10, 0.1]);
    viewer.zoomOut();
    expect(viewer.getZoom()).toBe(0.1);

    const toolbar = viewer.element.querySelector('.av-toolbar')!;
    toolbar.querySelector<HTMLButtonElement>('[aria-label="Reset zoom"]')!.click();
    expect(viewer.getZoom()).toBe(1);
    expect(toolbar.querySelector('[aria-label="Reset zoom"]')?.textContent).toBe('100%');
    toolbar.querySelector<HTMLButtonElement>('[aria-label="Zoom in"]')!.click();
    toolbar.querySelector<HTMLButtonElement>('[aria-label="Zoom out"]')!.click();
    expect(viewer.getZoom()).toBe(1);
  });

  it('navigates pages and clamps to the valid range', async () => {
    const { viewer } = withCapabilities(5);
    const changes: number[] = [];
    viewer.on('pagechange', ({ page }) => changes.push(page));
    await viewer.load(fakeFile());
    expect(viewer.getPage()).toEqual({ page: 1, pageCount: 5 });
    viewer.goToPage(3);
    viewer.goToPage(99);
    viewer.goToPage(-4);
    viewer.goToPage(1);
    viewer.goToPage(Number.NaN);
    viewer.goToPage(2.4);
    expect(changes).toEqual([3, 5, 1, 2]);

    const toolbar = viewer.element.querySelector('.av-toolbar')!;
    const previous = toolbar.querySelector<HTMLButtonElement>('[aria-label="Previous"]')!;
    const next = toolbar.querySelector<HTMLButtonElement>('[aria-label="Next"]')!;
    next.click();
    expect(viewer.getPage()?.page).toBe(3);
    previous.click();
    previous.click();
    expect(viewer.getPage()?.page).toBe(1);
    expect(previous.disabled).toBe(true);
    expect(toolbar.querySelector('.av-toolbar-info')?.textContent).toBe('1 / 5');
  });

  it('hides page controls for single-page files', async () => {
    const { viewer } = withCapabilities(1);
    await viewer.load(fakeFile());
    const next = viewer.element.querySelector<HTMLElement>('.av-toolbar [aria-label="Next"]')!;
    expect(next.parentElement!.hidden).toBe(true);
  });

  it('emits events when a renderer reports its own changes', async () => {
    const { viewer, context, setPage, setZoom } = withCapabilities();
    const pages = vi.fn();
    const zooms = vi.fn();
    viewer.on('pagechange', pages);
    viewer.on('zoomchange', zooms);
    await viewer.load(fakeFile());
    setPage(4);
    context().notify('page');
    setZoom(2);
    context().notify('zoom');
    expect(pages).toHaveBeenCalledWith({ page: 4, pageCount: 5 });
    expect(zooms).toHaveBeenCalledWith({ zoom: 2 });

    const stale = context();
    viewer.clear();
    stale.notify('page');
    expect(pages).toHaveBeenCalledTimes(1);
  });
});

describe('createViewer: actions', () => {
  it('downloads local files through a temporary object URL', async () => {
    vi.useFakeTimers();
    try {
      const click = vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
        this: HTMLAnchorElement,
      ) {
        expect(this.download).toBe('a.txt');
        expect(this.href).toMatch(/^blob:/);
      });
      const revoke = vi.spyOn(URL, 'revokeObjectURL');
      const viewer = createViewer(mount());
      await viewer.download();
      expect(click).not.toHaveBeenCalled();
      await viewer.load(textFile('hello', 'a.txt'));
      await viewer.download();
      expect(click).toHaveBeenCalledTimes(1);
      expect(document.querySelector('a[download]')).toBeNull();
      vi.advanceTimersByTime(60_000);
      expect(revoke).toHaveBeenCalledTimes(1);
    } finally {
      vi.useRealTimers();
    }
  });

  it('downloads streamed remote files straight from their URL', async () => {
    const fetchMock = mockFetch(() => new Response('x'));
    let href = '';
    let target = '';
    vi.spyOn(HTMLAnchorElement.prototype, 'click').mockImplementation(function (
      this: HTMLAnchorElement,
    ) {
      href = this.href;
      target = this.target;
    });
    const viewer = createViewer(mount(), { renderers: [fakeRenderer(() => undefined)] });
    await viewer.load('https://example.com/movie.fake');
    viewer.element.querySelector<HTMLButtonElement>('.av-toolbar [aria-label="Download"]')!.click();
    await waitFor(() => href);
    expect(href).toBe('https://example.com/movie.fake');
    // A cross-origin "download" may navigate instead, so it must not replace the app.
    expect(target).toBe('_blank');
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('downloads from the error card', async () => {
    const click = vi
      .spyOn(HTMLAnchorElement.prototype, 'click')
      .mockImplementation(() => undefined);
    const viewer = createViewer(mount());
    await viewer.load(new File(['x'], 'legacy.doc')).catch(() => undefined);
    viewer.element.querySelector<HTMLButtonElement>('.av-status button')!.click();
    await waitFor(() => click.mock.calls.length > 0);
  });

  it('prints through the renderer when it knows how', async () => {
    const print = vi.fn();
    const viewer = createViewer(mount(), { renderers: [fakeRenderer(() => ({ print }))] });
    viewer.print();
    await viewer.load(fakeFile());
    viewer.element.querySelector<HTMLButtonElement>('.av-toolbar [aria-label="Print"]')!.click();
    expect(print).toHaveBeenCalledTimes(1);
  });

  it('toggles full screen on the root element', async () => {
    const viewer = createViewer(mount());
    const request = vi.fn(() => Promise.resolve());
    const exit = vi.fn(() => Promise.resolve());
    viewer.element.requestFullscreen = request;
    document.exitFullscreen = exit;
    await viewer.toggleFullscreen();
    expect(request).toHaveBeenCalledTimes(1);
    Object.defineProperty(document, 'fullscreenElement', {
      configurable: true,
      get: () => viewer.element,
    });
    await viewer.toggleFullscreen();
    expect(exit).toHaveBeenCalledTimes(1);
    Object.defineProperty(document, 'fullscreenElement', { configurable: true, get: () => null });
  });

  it('opens contained files in a nested viewer with a way back', async () => {
    let context: RenderContext | null = null;
    const viewer = createViewer(mount(), {
      renderers: [fakeRenderer((received) => void (context = received))],
    });
    await viewer.load(fakeFile());
    context!.openChild(new Blob(['# Inner']), 'inner.md');
    const child = await waitFor(() => viewer.element.querySelector('.av-child'));
    await waitFor(() => child.querySelector('h1'));
    expect(child.querySelector('h1')?.textContent).toBe('Inner');

    // Opening another child replaces the first one.
    context!.openChild(new Blob(['plain']), 'other.txt');
    expect(viewer.element.querySelectorAll('.av-child')).toHaveLength(1);

    viewer.element.querySelector<HTMLButtonElement>('.av-child-bar button')!.click();
    expect(viewer.element.querySelector('.av-child')).toBeNull();

    context!.openChild(new Blob(['again']), 'again.txt');
    viewer.clear();
    expect(viewer.element.querySelector('.av-child')).toBeNull();
  });

  it('gives renderers a file URL: the original when usable, otherwise an object URL', async () => {
    const urls: string[] = [];
    const definition = fakeRenderer(async (context) => void urls.push(await context.fileUrl()));
    mockFetch(() => new Response('x'));
    const viewer = createViewer(mount(), { renderers: [definition] });
    await viewer.load('https://example.com/a.fake');
    await viewer.load(fakeFile());
    expect(urls[0]).toBe('https://example.com/a.fake');
    expect(urls[1]).toMatch(/^blob:/);
  });
});
