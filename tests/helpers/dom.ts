import { vi } from 'vitest';
import { createViewer } from '../../src';
import type {
  LoadedFile,
  RenderContext,
  RendererInstance,
  Viewer,
  ViewerOptions,
  ViewerSource,
} from '../../src';
import { getFormatById } from '../../src';
import { DEFAULT_LABELS } from '../../src';
import { decodeText } from '../../src/util/text';

/** Creates an attached container element. */
export function mount(): HTMLElement {
  const container = document.createElement('div');
  document.body.appendChild(container);
  return container;
}

/** Waits until `predicate` returns a truthy value, polling the task queue. */
export async function waitFor<T>(predicate: () => T, timeout = 2000): Promise<NonNullable<T>> {
  const started = Date.now();
  for (;;) {
    const value = predicate();
    if (value) return value;
    if (Date.now() - started > timeout) throw new Error('waitFor: timed out');
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
}

/** Creates a viewer, loads a source and resolves once it is ready. */
export async function open(
  source: ViewerSource,
  options: ViewerOptions & { name?: string } = {},
): Promise<{ viewer: Viewer; container: HTMLElement }> {
  const container = mount();
  const viewer = createViewer(container, options);
  await viewer.load(source);
  return { viewer, container };
}

export function textFile(content: string, name: string, type = ''): File {
  return new File([content], name, { type });
}

export function binaryFile(content: Uint8Array, name: string, type = ''): File {
  return new File([content as Uint8Array<ArrayBuffer>], name, { type });
}

export interface TestContext {
  context: RenderContext;
  container: HTMLElement;
  objectUrls: Blob[];
  children: { source: Blob; name: string }[];
  notifications: string[];
  abort(): void;
}

/** Builds a `RenderContext` for testing a renderer in isolation. */
export function renderContext(
  content: Uint8Array | string,
  name: string,
  formatId: string,
  overrides: Partial<RenderContext> = {},
): TestContext {
  const bytes = typeof content === 'string' ? new TextEncoder().encode(content) : content;
  const container = mount();
  const controller = new AbortController();
  const objectUrls: Blob[] = [];
  const children: { source: Blob; name: string }[] = [];
  const notifications: string[] = [];
  const blob = new Blob([bytes as Uint8Array<ArrayBuffer>]);
  const file: LoadedFile = {
    name,
    mimeType: '',
    size: bytes.length,
    url: null,
    direct: false,
    blob: () => Promise.resolve(blob),
    bytes: () => Promise.resolve(bytes),
    text: () => Promise.resolve(decodeText(bytes)),
  };
  const context: RenderContext = {
    file,
    format: getFormatById(formatId)!,
    language: null,
    container,
    signal: controller.signal,
    options: {
      name: undefined,
      type: undefined,
      toolbar: null,
      theme: 'auto',
      maxFileSize: Infinity,
      fetchOptions: undefined,
      renderers: [],
      pdfjs: undefined,
      className: undefined,
      styleNonce: undefined,
    },
    labels: DEFAULT_LABELS,
    objectUrl(value: Blob): string {
      objectUrls.push(value);
      return `blob:test/${objectUrls.length}`;
    },
    fileUrl: () => Promise.resolve('blob:test/file'),
    notify: (change) => void notifications.push(change),
    openChild: (source, childName) => void children.push({ source, name: childName }),
    ...overrides,
  };
  return {
    context,
    container,
    objectUrls,
    children,
    notifications,
    abort: () => controller.abort(),
  };
}

/** Narrows a renderer result to an instance. */
export function instanceOf(result: RendererInstance | void): RendererInstance {
  return result ?? {};
}

/** Replaces `fetch` for one test. */
export function mockFetch(
  handler: (url: string, init?: RequestInit) => Response | Promise<Response>,
): ReturnType<typeof vi.fn> {
  const mock = vi.fn((input: RequestInfo | URL, init?: RequestInit) =>
    Promise.resolve(
      handler(
        typeof input === 'string' ? input : input instanceof URL ? input.href : input.url,
        init,
      ),
    ),
  );
  vi.stubGlobal('fetch', mock);
  return mock;
}
