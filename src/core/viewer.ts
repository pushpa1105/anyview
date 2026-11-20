import { clear, h, icon } from '../ui/dom';
import { printNode } from '../ui/print';
import { ensureStyles } from '../ui/styles';
import { createToolbar, type Toolbar } from '../ui/toolbar';
import { extensionOf } from '../util/url';
import { AnyViewError, toAnyViewError } from './errors';
import { detectFormat, getFormatByExtension, languageFor, resolveFormatHint } from './formats';
import { DEFAULT_LABELS, type Labels } from './labels';
import { findRenderer, matchCustomRenderer } from './registry';
import { createFile } from './source';
import type {
  FormatInfo,
  LoadedFile,
  LoadOptions,
  RenderContext,
  RendererInstance,
  ResolvedViewerOptions,
  ToolbarOptions,
  Viewer,
  ViewerEventHandler,
  ViewerEventMap,
  ViewerEventName,
  ViewerOptions,
  ViewerSource,
  ViewerState,
  ViewerTheme,
} from './types';

const THEMES: readonly ViewerTheme[] = ['light', 'dark', 'auto'];
const DEFAULT_MAX_FILE_SIZE = 200 * 1024 * 1024;
const ZOOM_STEPS = [
  0.1, 0.25, 0.33, 0.5, 0.67, 0.75, 0.9, 1, 1.1, 1.25, 1.5, 1.75, 2, 2.5, 3, 4, 5, 6, 8, 10,
];
const MIN_ZOOM = ZOOM_STEPS[0]!;
const MAX_ZOOM = ZOOM_STEPS[ZOOM_STEPS.length - 1]!;
const ALL_TOOLBAR_ITEMS: Required<ToolbarOptions> = {
  fileName: true,
  zoom: true,
  pagination: true,
  download: true,
  print: true,
  fullscreen: true,
};

function invalidOption(message: string): AnyViewError {
  return new AnyViewError('INVALID_OPTIONS', message);
}

function resolveOptions(options: ViewerOptions): ResolvedViewerOptions {
  if (typeof options !== 'object' || options === null) {
    throw invalidOption('Viewer options must be an object.');
  }
  const theme = options.theme ?? 'auto';
  if (!THEMES.includes(theme)) {
    throw invalidOption(`Invalid theme "${String(theme)}". Use "light", "dark" or "auto".`);
  }
  const maxFileSize = options.maxFileSize ?? DEFAULT_MAX_FILE_SIZE;
  if (typeof maxFileSize !== 'number' || !(maxFileSize > 0)) {
    throw invalidOption('maxFileSize must be a positive number of bytes.');
  }
  const renderers = options.renderers ?? [];
  if (!Array.isArray(renderers)) throw invalidOption('renderers must be an array.');
  for (const definition of renderers as readonly unknown[]) {
    const candidate = definition as { id?: unknown; load?: unknown } | null;
    if (!candidate || typeof candidate.id !== 'string' || typeof candidate.load !== 'function') {
      throw invalidOption('Each custom renderer needs a string "id" and a "load" function.');
    }
  }
  const toolbar =
    options.toolbar === false
      ? null
      : { ...ALL_TOOLBAR_ITEMS, ...(typeof options.toolbar === 'object' ? options.toolbar : {}) };
  return {
    name: options.name,
    type: options.type,
    toolbar,
    theme,
    maxFileSize,
    fetchOptions: options.fetchOptions,
    renderers,
    pdfjs: options.pdfjs,
    className: options.className,
    styleNonce: options.styleNonce,
  };
}

function resolveContainer(target: HTMLElement | string): HTMLElement {
  if (typeof document === 'undefined') {
    throw invalidOption('createViewer() needs a browser environment (document is not defined).');
  }
  if (typeof target === 'string') {
    const found = document.querySelector<HTMLElement>(target);
    if (!found) throw invalidOption(`No element matches the selector "${target}".`);
    return found;
  }
  if (!target || typeof target !== 'object' || (target as Node).nodeType !== 1) {
    throw invalidOption('The container must be an HTMLElement or a CSS selector.');
  }
  return target;
}

function reportHandlerError(error: unknown): void {
  // A faulty event handler must not break the viewer; surface it like an uncaught error.
  if (typeof reportError === 'function') reportError(error);
  else console.error(error);
}

/**
 * Creates a viewer inside `container`.
 *
 * @example
 * const viewer = createViewer('#preview', { src: '/files/report.xlsx' });
 * viewer.on('error', (error) => console.warn(error.code, error.message));
 */
export function createViewer(container: HTMLElement | string, options: ViewerOptions = {}): Viewer {
  const host = resolveContainer(container);
  const resolved = resolveOptions(options);
  const labels: Readonly<Labels> = { ...DEFAULT_LABELS, ...options.labels };
  ensureStyles(host.ownerDocument, resolved.styleNonce);

  const content = h('div', { class: 'av-content' });
  const status = h('div', { class: 'av-status', attrs: { role: 'status' } });
  const body = h('div', { class: 'av-body' }, content, status);
  const root = h('div', {
    class: resolved.className ? `av-root ${resolved.className}` : 'av-root',
    attrs: { 'data-theme': resolved.theme },
  });

  const listeners = new Map<ViewerEventName, Set<(detail: never) => void>>();
  const objectUrls: string[] = [];
  let state: ViewerState = 'idle';
  let format: FormatInfo | null = null;
  let file: LoadedFile | null = null;
  let instance: RendererInstance | null = null;
  let controller: AbortController | null = null;
  let child: { viewer: Viewer; element: HTMLElement } | null = null;
  let baseZoom = 1;

  const emit = <K extends ViewerEventName>(event: K, detail: ViewerEventMap[K]): void => {
    const handlers = listeners.get(event);
    if (!handlers) return;
    for (const handler of Array.from(handlers)) {
      try {
        (handler as ViewerEventHandler<K>)(detail);
      } catch (error) {
        reportHandlerError(error);
      }
    }
  };

  const getPage = (): { page: number; pageCount: number } | null =>
    instance?.pages ? { page: instance.pages.current(), pageCount: instance.pages.count() } : null;
  const getZoom = (): number | null => instance?.zoom?.get() ?? null;

  let toolbar: Toolbar | null = null;
  const refreshToolbar = (): void => {
    toolbar?.update({
      fileName: file?.name ?? '',
      formatLabel: format && file ? format.label : '',
      zoom: getZoom(),
      page: getPage(),
      hasFile: file !== null && state !== 'loading',
      canPrint: state === 'ready' && instance !== null,
    });
  };

  const setZoom = (value: number): void => {
    const zoom = instance?.zoom;
    if (!zoom || !Number.isFinite(value)) return;
    const clamped = Math.min(MAX_ZOOM, Math.max(MIN_ZOOM, value));
    if (clamped === zoom.get()) return;
    zoom.set(clamped);
    refreshToolbar();
    emit('zoomchange', { zoom: zoom.get() });
  };
  const stepZoom = (direction: 1 | -1): void => {
    const current = getZoom();
    if (current === null) return;
    const next =
      direction === 1
        ? ZOOM_STEPS.find((step) => step > current + 1e-6)
        : [...ZOOM_STEPS].reverse().find((step) => step < current - 1e-6);
    if (next !== undefined) setZoom(next);
  };
  const goToPage = (page: number): void => {
    const pages = instance?.pages;
    if (!pages || !Number.isFinite(page)) return;
    const target = Math.min(pages.count(), Math.max(1, Math.round(page)));
    if (target === pages.current()) return;
    pages.goTo(target);
    refreshToolbar();
    emit('pagechange', { page: pages.current(), pageCount: pages.count() });
  };

  const download = async (): Promise<void> => {
    const current = file;
    if (!current) return;
    let href: string;
    let temporary = false;
    if (current.size === null && current.direct && current.url) {
      // Not downloaded yet (streamed media): let the browser fetch it itself.
      href = current.url;
    } else {
      href = URL.createObjectURL(await current.blob());
      temporary = true;
    }
    const link = h('a', {
      attrs: {
        href,
        download: current.name || 'download',
        rel: 'noopener',
        // Browsers ignore "download" for other origins and navigate instead: keep the app open.
        target: temporary ? undefined : '_blank',
      },
      style: { display: 'none' },
    });
    document.body.appendChild(link);
    link.click();
    link.remove();
    if (temporary) setTimeout(() => URL.revokeObjectURL(href), 60_000);
  };

  const print = (): void => {
    if (state !== 'ready' || !instance) return;
    if (instance.print) instance.print();
    else printNode(instance.printable?.() ?? content, file?.name ?? '');
  };

  const toggleFullscreen = async (): Promise<void> => {
    const doc = root.ownerDocument;
    if (doc.fullscreenElement === root) await doc.exitFullscreen();
    else if (typeof root.requestFullscreen === 'function') await root.requestFullscreen();
  };

  if (resolved.toolbar) {
    toolbar = createToolbar(labels, resolved.toolbar, {
      zoomIn: () => stepZoom(1),
      zoomOut: () => stepZoom(-1),
      resetZoom: () => setZoom(baseZoom),
      previousPage: () => goToPage((getPage()?.page ?? 1) - 1),
      nextPage: () => goToPage((getPage()?.page ?? 1) + 1),
      download: () => void download().catch(reportHandlerError),
      print,
      fullscreen: () => void toggleFullscreen().catch(() => undefined),
    });
    root.appendChild(toolbar.element);
  }
  root.appendChild(body);
  host.appendChild(root);

  const showStatus = (...children: (Node | null)[]): void => {
    clear(status);
    for (const node of children) if (node) status.appendChild(node);
    status.hidden = false;
    // Covered by the status layer but still laid out, so renderers can measure their container.
    content.style.visibility = 'hidden';
    content.setAttribute('aria-hidden', 'true');
  };
  const showEmpty = (): void => showStatus(h('p', { class: 'av-status-text', text: labels.empty }));

  const showError = (error: AnyViewError): void => {
    const unsupported = error.code === 'UNSUPPORTED_FORMAT';
    const canDownload =
      file !== null && error.code !== 'FETCH_FAILED' && error.code !== 'INVALID_SOURCE';
    const iconHolder = h('div', { class: 'av-status-icon' });
    iconHolder.appendChild(icon(unsupported ? 'file' : 'warning', 40));
    showStatus(
      iconHolder,
      h('p', {
        class: 'av-status-title',
        text: unsupported ? labels.unsupportedTitle : labels.errorTitle,
      }),
      h('p', { class: 'av-status-text', text: error.message }),
      unsupported && canDownload
        ? h('p', { class: 'av-status-text', text: labels.unsupportedHint })
        : null,
      canDownload
        ? h('button', {
            class: 'av-btn av-btn-text',
            text: labels.download,
            attrs: { type: 'button' },
            on: { click: () => void download().catch(reportHandlerError) },
          })
        : null,
    );
    status.setAttribute('role', 'alert');
  };

  const closeChild = (): void => {
    if (!child) return;
    child.viewer.destroy();
    child.element.remove();
    child = null;
  };

  /** Releases everything that belongs to the currently open file. */
  const teardown = (): void => {
    controller?.abort();
    controller = null;
    closeChild();
    const previous = instance;
    instance = null;
    try {
      previous?.destroy?.();
    } catch (error) {
      reportHandlerError(error);
    }
    clear(content);
    for (const url of objectUrls.splice(0)) URL.revokeObjectURL(url);
    file = null;
    format = null;
  };

  const openChild = (source: Blob, name: string): void => {
    closeChild();
    const childBody = h('div', { class: 'av-child-body' });
    const backButton = h('button', {
      class: 'av-btn',
      attrs: { type: 'button', 'aria-label': labels.back, title: labels.back },
      on: { click: closeChild },
    });
    backButton.append(icon('back'), labels.back);
    const element = h(
      'div',
      { class: 'av-child' },
      h('div', { class: 'av-child-bar' }, backButton),
      childBody,
    );
    body.appendChild(element);
    const viewer = createViewer(childBody, {
      src: source,
      name,
      theme: resolved.theme,
      toolbar: options.toolbar ?? true,
      maxFileSize: resolved.maxFileSize,
      renderers: resolved.renderers,
      pdfjs: resolved.pdfjs,
      labels,
    });
    child = { viewer, element };
  };

  const assertAlive = (): void => {
    if (state === 'destroyed') {
      throw new AnyViewError('DESTROYED', 'This viewer has been destroyed and cannot be used.');
    }
  };

  const load = async (source: ViewerSource, loadOptions: LoadOptions = {}): Promise<void> => {
    assertAlive();
    teardown();
    const current = new AbortController();
    controller = current;
    const { signal } = current;
    const superseded = (): boolean => signal.aborted || controller !== current;

    state = 'loading';
    status.setAttribute('role', 'status');
    showStatus(
      h('div', { class: 'av-spinner' }),
      h('p', { class: 'av-status-text', text: labels.loading }),
    );
    refreshToolbar();
    emit('loadstart', { source });

    try {
      const opened = createFile(source, {
        name: loadOptions.name ?? resolved.name,
        fetchOptions: resolved.fetchOptions,
        maxFileSize: resolved.maxFileSize,
        signal,
      });
      file = opened;
      const hint = loadOptions.type ?? resolved.type;

      // When neither a hint nor the extension identifies a remote file, download
      // it first: the response headers may carry its real name and MIME type.
      const identified =
        (hint !== undefined && resolveFormatHint(hint) !== null) ||
        getFormatByExtension(extensionOf(opened.name)) !== null ||
        matchCustomRenderer(resolved.renderers, opened.name, opened.mimeType) !== null;
      if (!identified && opened.size === null) await opened.blob();
      if (superseded()) return;

      const custom = hint
        ? null
        : matchCustomRenderer(resolved.renderers, opened.name, opened.mimeType);
      const detected =
        custom?.format ??
        (await detectFormat({
          hint,
          name: opened.name,
          mimeType: opened.mimeType,
          bytes: () => opened.bytes(),
        }));
      if (superseded()) return;
      format = detected;

      const definition = custom?.definition ?? findRenderer(detected.id, resolved.renderers);
      if (!definition) {
        const what = detected.id === 'unknown' ? 'this type of file' : `${detected.label} files`;
        throw new AnyViewError('UNSUPPORTED_FORMAT', `AnyView cannot display ${what}.`, {
          details: { format: detected.id, name: opened.name },
        });
      }

      const renderer = await definition.load();
      if (superseded()) return;

      const context: RenderContext = {
        file: opened,
        format: detected,
        language: languageFor(opened.name, detected.id, hint),
        container: content,
        signal,
        options: resolved,
        labels,
        objectUrl(blob: Blob): string {
          const url = URL.createObjectURL(blob);
          objectUrls.push(url);
          return url;
        },
        async fileUrl(): Promise<string> {
          if (opened.direct && opened.url) return opened.url;
          return context.objectUrl(await opened.blob());
        },
        notify(change): void {
          if (superseded()) return;
          refreshToolbar();
          if (change === 'zoom') {
            const zoom = getZoom();
            if (zoom !== null) emit('zoomchange', { zoom });
          } else {
            const page = getPage();
            if (page) emit('pagechange', page);
          }
        },
        openChild,
      };

      const rendered = (await renderer.render(context)) ?? {};
      if (superseded()) {
        rendered.destroy?.();
        return;
      }
      instance = rendered;
      baseZoom = rendered.zoom?.get() ?? 1;
      state = 'ready';
      status.hidden = true;
      content.style.visibility = '';
      content.removeAttribute('aria-hidden');
      refreshToolbar();
      emit('load', {
        format: detected,
        file: { name: opened.name, size: opened.size, mimeType: opened.mimeType },
      });
    } catch (cause) {
      if (superseded()) return;
      const error = toAnyViewError(cause, 'RENDER_FAILED', 'The file could not be displayed');
      clear(content);
      state = 'error';
      showError(error);
      refreshToolbar();
      emit('error', error);
      throw error;
    }
  };

  const viewer: Viewer = {
    element: root,
    get state() {
      return state;
    },
    get format() {
      return format;
    },
    load,
    clear(): void {
      assertAlive();
      teardown();
      state = 'idle';
      status.setAttribute('role', 'status');
      showEmpty();
      refreshToolbar();
    },
    on(event, handler) {
      let handlers = listeners.get(event);
      if (!handlers) listeners.set(event, (handlers = new Set()));
      handlers.add(handler);
      return () => viewer.off(event, handler);
    },
    off(event, handler) {
      listeners.get(event)?.delete(handler);
    },
    getZoom,
    setZoom,
    zoomIn: () => stepZoom(1),
    zoomOut: () => stepZoom(-1),
    getPage,
    goToPage,
    download,
    print,
    toggleFullscreen,
    setTheme(theme: ViewerTheme): void {
      if (!THEMES.includes(theme)) {
        throw invalidOption(`Invalid theme "${String(theme)}". Use "light", "dark" or "auto".`);
      }
      resolved.theme = theme;
      root.setAttribute('data-theme', theme);
    },
    destroy(): void {
      if (state === 'destroyed') return;
      teardown();
      state = 'destroyed';
      root.remove();
      emit('destroy', undefined);
      listeners.clear();
    },
  };

  showEmpty();
  refreshToolbar();
  if (options.src !== undefined && options.src !== null) {
    // Deferred by a microtask so callers can subscribe to events first. Failures of
    // the initial load are delivered through the "error" event and the UI.
    const initial = options.src;
    void Promise.resolve().then(() => {
      if (state === 'idle' && controller === null) load(initial).catch(() => undefined);
    });
  }
  return viewer;
}
