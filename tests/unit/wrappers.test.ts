import { createElement, createRef, StrictMode } from 'react';
import { act } from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { afterEach, describe, expect, it, vi } from 'vitest';
import { createApp, h, nextTick, ref, type App } from 'vue';
import { defineAnyViewElement, type AnyViewElement } from '../../src/element';
import { parseEmail } from '../../src/formats/eml';
import { AnyView as ReactAnyView, type AnyViewHandle } from '../../src/react';
import { AnyView as VueAnyView } from '../../src/vue';
import type { Viewer } from '../../src';
import { mount, textFile, waitFor } from '../helpers/dom';

(globalThis as { IS_REACT_ACT_ENVIRONMENT?: boolean }).IS_REACT_ACT_ENVIRONMENT = true;

afterEach(() => {
  vi.unstubAllGlobals();
  document.body.textContent = '';
});

const fileName = (root: ParentNode): string =>
  root.querySelector('.av-file-name')?.textContent ?? '';

describe('<any-view> custom element', () => {
  const create = (attributes: Record<string, string> = {}): AnyViewElement => {
    const element = document.createElement('any-view');
    for (const [name, value] of Object.entries(attributes)) element.setAttribute(name, value);
    return element;
  };

  it('registers itself once', () => {
    expect(customElements.get('any-view')).toBeDefined();
    expect(() => defineAnyViewElement()).not.toThrow();
    defineAnyViewElement('any-view-alias');
    expect(customElements.get('any-view-alias')).toBeDefined();
  });

  it('does nothing where custom elements do not exist', () => {
    vi.stubGlobal('customElements', undefined);
    expect(() => defineAnyViewElement('never-defined')).not.toThrow();
  });

  it('creates a viewer when connected and destroys it when removed', () => {
    const element = create({ theme: 'dark' });
    expect(element.viewer).toBeNull();
    document.body.appendChild(element);
    expect(element.style.display).toBe('block');
    expect(element.viewer?.state).toBe('idle');
    expect(element.querySelector('.av-root')?.getAttribute('data-theme')).toBe('dark');
    const viewer = element.viewer!;
    element.remove();
    expect(viewer.state).toBe('destroyed');
    expect(element.viewer).toBeNull();
  });

  it('loads the source property and forwards events', async () => {
    const element = create();
    document.body.appendChild(element);
    const loaded = vi.fn();
    const started = vi.fn();
    element.addEventListener('av-load', loaded);
    element.addEventListener('av-loadstart', started);
    element.source = textFile('# Hello', 'readme.md');
    await waitFor(() => loaded.mock.calls.length > 0);
    expect(started).toHaveBeenCalledTimes(1);
    const event = loaded.mock.calls[0]![0] as CustomEvent<{ format: { id: string } }>;
    expect(event.detail.format.id).toBe('markdown');
    expect(element.querySelector('h1')?.textContent).toBe('Hello');
    expect(element.source).toBeInstanceOf(File);
  });

  it('reports failures as av-error events', async () => {
    const element = create();
    document.body.appendChild(element);
    const failed = vi.fn();
    element.addEventListener('av-error', failed);
    element.source = new File(['x'], 'legacy.doc');
    await waitFor(() => failed.mock.calls.length > 0);
    expect((failed.mock.calls[0]![0] as CustomEvent<{ code: string }>).detail.code).toBe(
      'UNSUPPORTED_FORMAT',
    );
  });

  it('reacts to src, name and type attributes', async () => {
    vi.stubGlobal(
      'fetch',
      vi.fn(() => Promise.resolve(new Response('a,b\n1,2'))),
    );
    const element = create({ src: 'https://example.com/export', type: 'csv', name: 'export.csv' });
    document.body.appendChild(element);
    await waitFor(() => element.viewer?.state === 'ready');
    expect(element.viewer?.format?.id).toBe('csv');
    expect(fileName(element)).toBe('export.csv');
    expect(element.source).toBe('https://example.com/export');

    element.setAttribute('type', 'text');
    await waitFor(() => element.viewer?.format?.id === 'text');

    // Setting the attribute replaces a source given as a property.
    element.source = textFile('x', 'prop.txt');
    await waitFor(() => fileName(element) === 'export.csv' && element.viewer?.state === 'ready');
    element.setAttribute('src', 'https://example.com/other');
    expect(element.source).toBe('https://example.com/other');

    element.removeAttribute('src');
    await waitFor(() => element.viewer?.state === 'idle');
  });

  it('batches several changes into a single load', async () => {
    const element = create();
    document.body.appendChild(element);
    const started = vi.fn();
    element.addEventListener('av-loadstart', started);
    element.setAttribute('name', 'renamed.txt');
    element.source = textFile('one', 'one.txt');
    element.source = textFile('two', 'two.txt');
    await waitFor(() => element.viewer?.state === 'ready');
    expect(started).toHaveBeenCalledTimes(1);
    expect(fileName(element)).toBe('renamed.txt');
  });

  it('switches theme and toolbar through attributes', () => {
    const element = create({ toolbar: 'false' });
    document.body.appendChild(element);
    expect(element.querySelector('.av-toolbar')).toBeNull();
    element.setAttribute('toolbar', 'true');
    expect(element.querySelector('.av-toolbar')).not.toBeNull();
    element.setAttribute('theme', 'light');
    expect(element.querySelector('.av-root')?.getAttribute('data-theme')).toBe('light');
    element.setAttribute('theme', 'nonsense');
    expect(element.querySelector('.av-root')?.getAttribute('data-theme')).toBe('auto');
    element.setAttribute('theme', 'nonsense');
  });

  it('applies the options property, including defaults for src, name and toolbar', async () => {
    const element = create();
    expect(element.options).toEqual({});
    element.options = { toolbar: false, labels: { empty: 'Drop a file' } };
    document.body.appendChild(element);
    expect(element.querySelector('.av-toolbar')).toBeNull();
    expect(element.textContent).toBe('Drop a file');

    element.options = {
      src: textFile('body', 'from-options.txt'),
      name: 'renamed.txt',
      type: 'text',
    };
    await waitFor(() => fileName(element) === 'renamed.txt');
    element.options = null as never;
    expect(element.options).toEqual({});
  });
});

describe('React <AnyView>', () => {
  let root: Root | null = null;

  afterEach(() => {
    if (root) act(() => root!.unmount());
    root = null;
  });

  async function render(element: ReturnType<typeof createElement>): Promise<HTMLElement> {
    const container = mount();
    root = createRoot(container);
    await act(async () => {
      root!.render(element);
      await Promise.resolve();
    });
    return container;
  }

  it('mounts a viewer, loads src and calls back', async () => {
    const onLoad = vi.fn();
    const onLoadStart = vi.fn();
    const handle = createRef<AnyViewHandle>();
    const container = await render(
      createElement(ReactAnyView, {
        ref: handle,
        src: textFile('hello', 'a.txt'),
        className: 'wrapper',
        id: 'preview',
        style: { height: 320 },
        onLoad,
        onLoadStart,
      }),
    );
    await waitFor(() => onLoad.mock.calls.length > 0);
    const wrapper = container.firstElementChild as HTMLElement;
    expect(wrapper.id).toBe('preview');
    expect(wrapper.className).toBe('wrapper');
    expect(wrapper.style.height).toBe('320px');
    expect(fileName(container)).toBe('a.txt');
    expect(onLoadStart).toHaveBeenCalledTimes(1);
    expect(handle.current?.viewer?.state).toBe('ready');
  });

  it('reacts to src, name, type and theme changes without recreating the viewer', async () => {
    const container = await render(
      createElement(ReactAnyView, { src: textFile('one', 'one.txt'), theme: 'dark' }),
    );
    await waitFor(() => fileName(container) === 'one.txt');
    const viewerRoot = container.querySelector('.av-root');
    expect(viewerRoot?.getAttribute('data-theme')).toBe('dark');

    await act(async () => {
      root!.render(
        createElement(ReactAnyView, {
          src: new Blob(['a,b']),
          name: 'two.csv',
          type: 'csv',
          theme: 'light',
        }),
      );
      await Promise.resolve();
    });
    await waitFor(() => fileName(container) === 'two.csv');
    expect(container.querySelector('.av-root')).toBe(viewerRoot);
    expect(viewerRoot?.getAttribute('data-theme')).toBe('light');

    await act(async () => {
      root!.render(createElement(ReactAnyView, { src: null }));
      await Promise.resolve();
    });
    expect(container.querySelector('.av-status')?.textContent).toBe('No file selected');
    expect(viewerRoot?.getAttribute('data-theme')).toBe('auto');
  });

  it('reports errors, page and zoom changes through props', async () => {
    const onError = vi.fn();
    const onZoomChange = vi.fn();
    const onPageChange = vi.fn();
    const handle = createRef<AnyViewHandle>();
    await render(
      createElement(ReactAnyView, {
        ref: handle,
        src: new File(['x'], 'legacy.doc'),
        onError,
        onZoomChange,
        onPageChange,
      }),
    );
    await waitFor(() => onError.mock.calls.length > 0);
    expect(onError.mock.calls[0]![0]).toMatchObject({ code: 'UNSUPPORTED_FORMAT' });

    await act(async () => {
      await handle.current!.viewer!.load(textFile('x', 'a.txt'));
    });
    act(() => handle.current!.viewer!.zoomIn());
    expect(onZoomChange).toHaveBeenCalledWith({ zoom: 1.1 });
    expect(onPageChange).not.toHaveBeenCalled();
  });

  it('passes construction options through', async () => {
    const renderer = { render: vi.fn() };
    const container = await render(
      createElement(ReactAnyView, {
        src: textFile('x', 'thing.fake'),
        toolbar: false,
        maxFileSize: 1000,
        fetchOptions: { credentials: 'include' },
        labels: { loading: 'Please wait' },
        styleNonce: 'n',
        pdfjs: { getDocument: () => ({ promise: Promise.reject(new Error('unused')) }) },
        renderers: [{ id: 'fake', extensions: ['fake'], load: () => renderer }],
      }),
    );
    await waitFor(() => renderer.render.mock.calls.length > 0);
    expect(container.querySelector('.av-toolbar')).toBeNull();
  });

  it('survives StrictMode double mounting and cleans up on unmount', async () => {
    const handle = createRef<AnyViewHandle>();
    const container = await render(
      createElement(
        StrictMode,
        null,
        createElement(ReactAnyView, { ref: handle, src: textFile('x', 'a.txt') }),
      ),
    );
    await waitFor(() => fileName(container) === 'a.txt');
    expect(container.querySelectorAll('.av-root')).toHaveLength(1);
    const viewer = handle.current!.viewer!;
    act(() => root!.unmount());
    root = null;
    expect(viewer.state).toBe('destroyed');
    expect(container.querySelector('.av-root')).toBeNull();
  });
});

describe('Vue <AnyView>', () => {
  let app: App | null = null;

  afterEach(() => {
    app?.unmount();
    app = null;
  });

  it('mounts a viewer, loads src and emits events', async () => {
    const onLoad = vi.fn();
    const onLoadstart = vi.fn();
    const container = mount();
    app = createApp({
      render: () => h(VueAnyView, { src: textFile('hello', 'a.txt'), onLoad, onLoadstart }),
    });
    app.mount(container);
    await waitFor(() => onLoad.mock.calls.length > 0);
    expect(fileName(container)).toBe('a.txt');
    expect(onLoadstart).toHaveBeenCalledTimes(1);
    expect((onLoad.mock.calls[0]![0] as { format: { id: string } }).format.id).toBe('text');
  });

  it('reacts to prop changes and exposes the viewer', async () => {
    const source = ref<File | Blob | null>(textFile('one', 'one.txt'));
    const name = ref<string | undefined>(undefined);
    const type = ref<string | undefined>(undefined);
    const theme = ref<'dark' | 'light'>('dark');
    const component = ref<{ viewer: Viewer | null } | null>(null);
    const onZoomchange = vi.fn();
    const onPagechange = vi.fn();
    const container = mount();
    app = createApp({
      render: () =>
        h(VueAnyView, {
          ref: component,
          src: source.value,
          name: name.value,
          type: type.value,
          theme: theme.value,
          onZoomchange,
          onPagechange,
        }),
    });
    app.mount(container);
    await waitFor(() => fileName(container) === 'one.txt');
    const viewerRoot = container.querySelector('.av-root')!;
    expect(viewerRoot.getAttribute('data-theme')).toBe('dark');

    source.value = new Blob(['a,b']);
    name.value = 'two.csv';
    type.value = 'csv';
    theme.value = 'light';
    await nextTick();
    await waitFor(() => fileName(container) === 'two.csv');
    expect(viewerRoot.getAttribute('data-theme')).toBe('light');
    expect(component.value?.viewer?.format?.id).toBe('csv');

    component.value!.viewer!.zoomIn();
    expect(onZoomchange).toHaveBeenCalledWith({ zoom: 1.1 });
    expect(onPagechange).not.toHaveBeenCalled();

    source.value = null;
    await nextTick();
    expect(container.querySelector('.av-status')?.textContent).toBe('No file selected');
  });

  it('emits errors and applies toolbar and options props', async () => {
    const onError = vi.fn();
    const container = mount();
    app = createApp({
      render: () =>
        h(VueAnyView, {
          src: new File(['x'], 'legacy.doc'),
          toolbar: false,
          options: { labels: { unsupportedTitle: 'Cannot preview' } },
          onError,
        }),
    });
    app.mount(container);
    await waitFor(() => onError.mock.calls.length > 0);
    expect(onError.mock.calls[0]![0]).toMatchObject({ code: 'UNSUPPORTED_FORMAT' });
    expect(container.querySelector('.av-toolbar')).toBeNull();
    expect(container.textContent).toContain('Cannot preview');
  });

  it('destroys the viewer when unmounted', async () => {
    const component = ref<{ viewer: Viewer | null } | null>(null);
    const container = mount();
    app = createApp({ render: () => h(VueAnyView, { ref: component }) });
    app.mount(container);
    await nextTick();
    const viewer = component.value!.viewer!;
    expect(viewer.state).toBe('idle');
    app.unmount();
    app = null;
    expect(viewer.state).toBe('destroyed');
  });
});

describe('regression: binary safety', () => {
  it('keeps bytes 0x80–0x9F intact in 8-bit email attachments', () => {
    // Browsers decode "latin1" as Windows-1252, which would remap these bytes.
    const payload = Uint8Array.from({ length: 256 }, (_unused, index) => index).subarray(
      0x7e,
      0xa2,
    );
    const head = new TextEncoder().encode(
      'Content-Type: application/octet-stream\r\nContent-Disposition: attachment; filename="raw.bin"\r\n\r\n',
    );
    const message = new Uint8Array(head.length + payload.length);
    message.set(head);
    message.set(payload, head.length);
    expect(Array.from(parseEmail(message).attachments[0]!.data)).toEqual(Array.from(payload));
  });
});
