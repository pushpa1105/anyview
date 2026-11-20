import type { AnyViewError } from './errors';
import type { Labels } from './labels';

/** Anything AnyView can open: a URL, a `File`/`Blob`, or raw bytes. */
export type ViewerSource = string | URL | Blob | ArrayBuffer | ArrayBufferView;

export type FormatCategory =
  | 'image'
  | 'video'
  | 'audio'
  | 'document'
  | 'spreadsheet'
  | 'presentation'
  | 'text'
  | 'archive'
  | 'font'
  | 'other';

/** A file format AnyView knows about. */
export interface FormatInfo {
  /** Stable identifier, e.g. `"pdf"`, `"docx"`, `"image"`. */
  readonly id: string;
  /** Human readable name, e.g. `"Word document"`. */
  readonly label: string;
  readonly category: FormatCategory;
  /** Lower-case extensions without the dot. */
  readonly extensions: readonly string[];
  readonly mimeTypes: readonly string[];
}

/** A file being viewed. Content is fetched lazily and cached. */
export interface LoadedFile {
  readonly name: string;
  /** MIME type if known, otherwise an empty string. */
  readonly mimeType: string;
  /** Size in bytes, or `null` while a remote file has not been downloaded. */
  readonly size: number | null;
  /** The original URL when the source was a URL, otherwise `null`. */
  readonly url: string | null;
  /** True when `url` can be handed straight to the browser (no custom fetch options). */
  readonly direct: boolean;
  blob(): Promise<Blob>;
  bytes(): Promise<Uint8Array>;
  /** Decoded text (UTF-8/UTF-16 by BOM, Windows-1252 as a last resort). */
  text(): Promise<string>;
}

export interface ZoomCapability {
  /** Current zoom, where `1` means 100%. */
  get(): number;
  set(value: number): void;
}

export interface PageCapability {
  count(): number;
  /** Current page, starting at 1. */
  current(): number;
  goTo(page: number): void;
}

/** What a renderer hands back so the viewer can drive it. All members are optional. */
export interface RendererInstance {
  /** Release everything the renderer created (listeners, observers, workers). */
  destroy?(): void;
  zoom?: ZoomCapability;
  pages?: PageCapability;
  /** Custom print implementation. When absent, `printable()` or the content is printed. */
  print?(): void;
  /** Node to print instead of the live content (e.g. a full table for a virtualized grid). */
  printable?(): Node | null;
}

export interface RenderContext {
  readonly file: LoadedFile;
  readonly format: FormatInfo;
  /** Syntax-highlighting language hint for text formats, e.g. `"typescript"`. */
  readonly language: string | null;
  /** Empty element the renderer fills. */
  readonly container: HTMLElement;
  /** Aborted when the viewer loads something else or is destroyed. */
  readonly signal: AbortSignal;
  readonly options: Readonly<ResolvedViewerOptions>;
  readonly labels: Readonly<Labels>;
  /** Creates an object URL that is revoked automatically when this file is closed. */
  objectUrl(blob: Blob): string;
  /** A URL for the whole file: the original URL when usable, otherwise an object URL. */
  fileUrl(): Promise<string>;
  /** Tell the viewer that zoom or page state changed so the toolbar and events update. */
  notify(change: 'zoom' | 'page'): void;
  /** Open a file contained in this one (archive entry, attachment) in a nested preview. */
  openChild(source: Blob, name: string): void;
}

export interface Renderer {
  render(context: RenderContext): RendererInstance | void | Promise<RendererInstance | void>;
}

/** Registers a renderer. `load` runs only when a matching file is opened. */
export interface RendererDefinition {
  /** Unique id. A custom definition with a built-in id replaces the built-in. */
  readonly id: string;
  /** Format ids this renderer handles (built-in ids such as `"pdf"` or your own). */
  readonly formats?: readonly string[];
  /** Extra extensions (no dot) that should be routed to this renderer. */
  readonly extensions?: readonly string[];
  /** Extra MIME types that should be routed to this renderer. */
  readonly mimeTypes?: readonly string[];
  load(): Renderer | Promise<Renderer>;
}

export type ViewerTheme = 'light' | 'dark' | 'auto';

export interface ToolbarOptions {
  /** Show file name and format. Default `true`. */
  fileName?: boolean;
  zoom?: boolean;
  pagination?: boolean;
  download?: boolean;
  print?: boolean;
  fullscreen?: boolean;
}

/**
 * Minimal shape of the `pdfjs-dist` module used by the optional PDF.js adapter.
 * Pass the module (or a function returning it) as the `pdfjs` option.
 */
export interface PdfJsLike {
  getDocument(source: { data: Uint8Array }): PdfJsLoadingTask;
}

export interface PdfJsLoadingTask {
  promise: Promise<PdfJsDocument>;
  destroy?(): Promise<void> | void;
}

export interface PdfJsDocument {
  numPages: number;
  getPage(pageNumber: number): Promise<PdfJsPage>;
  /** Present in PDF.js 2–4; later versions release the document through the loading task. */
  destroy?(): Promise<void> | void;
}

export interface PdfJsViewport {
  width: number;
  height: number;
}

export interface PdfJsPage {
  getViewport(params: { scale: number }): PdfJsViewport;
  render(params: {
    canvasContext: CanvasRenderingContext2D;
    viewport: PdfJsViewport;
    canvas?: HTMLCanvasElement;
  }): { promise: Promise<unknown>; cancel?(): void };
}

export interface ViewerOptions {
  /** File to open immediately. Errors are reported through the `error` event. */
  src?: ViewerSource | null;
  /** File name override; also used for format detection and downloads. */
  name?: string;
  /** Force a format: a format id (`"docx"`), an extension (`".docx"`) or a MIME type. */
  type?: string;
  /** `false` hides the toolbar; an object toggles individual controls. Default `true`. */
  toolbar?: boolean | ToolbarOptions;
  /** Default `"auto"` (follows the OS setting). */
  theme?: ViewerTheme;
  /** Reject files larger than this many bytes. Default 200 MB. */
  maxFileSize?: number;
  /** Passed to `fetch` for URL sources, e.g. `{ headers: { Authorization: … } }`. */
  fetchOptions?: RequestInit;
  /** Custom renderers; they take priority over built-in ones. */
  renderers?: readonly RendererDefinition[];
  /**
   * Optional PDF.js module. Without it PDFs use the browser's built-in viewer;
   * with it they are drawn to canvases for identical results on every device.
   */
  pdfjs?: PdfJsLike | (() => PdfJsLike | Promise<PdfJsLike>);
  /** Override any UI text (for translation). */
  labels?: Partial<Labels>;
  /**
   * Draw dashed gridlines on the edges of Word table cells that have no border, like Word does
   * on screen with "View Gridlines" switched on. They are not printed. Default `false`.
   */
  tableGridlines?: boolean;
  /** Extra class name(s) for the root element. */
  className?: string;
  /** CSP nonce for the injected `<style>` element. */
  styleNonce?: string;
}

export interface ResolvedViewerOptions {
  name: string | undefined;
  type: string | undefined;
  toolbar: Required<ToolbarOptions> | null;
  theme: ViewerTheme;
  maxFileSize: number;
  fetchOptions: RequestInit | undefined;
  renderers: readonly RendererDefinition[];
  pdfjs: ViewerOptions['pdfjs'];
  tableGridlines: boolean;
  className: string | undefined;
  styleNonce: string | undefined;
}

export type ViewerState = 'idle' | 'loading' | 'ready' | 'error' | 'destroyed';

export interface LoadEventDetail {
  format: FormatInfo;
  file: { name: string; size: number | null; mimeType: string };
}

export interface ViewerEventMap {
  loadstart: { source: ViewerSource };
  load: LoadEventDetail;
  error: AnyViewError;
  pagechange: { page: number; pageCount: number };
  zoomchange: { zoom: number };
  destroy: undefined;
}

export type ViewerEventName = keyof ViewerEventMap;
export type ViewerEventHandler<K extends ViewerEventName> = (detail: ViewerEventMap[K]) => void;

/** Per-load overrides for {@link Viewer.load}. */
export interface LoadOptions {
  name?: string;
  type?: string;
}

export interface Viewer {
  /** The root element created inside the container. */
  readonly element: HTMLElement;
  readonly state: ViewerState;
  /** Format of the current file, or `null` before the first successful detection. */
  readonly format: FormatInfo | null;
  /**
   * Opens a file, replacing the current one. Rejects with an `AnyViewError` on
   * failure; resolves quietly if superseded by a newer `load()` or `destroy()`.
   */
  load(source: ViewerSource, options?: LoadOptions): Promise<void>;
  /** Closes the current file and returns to the empty state. */
  clear(): void;
  /** Subscribes to an event and returns an unsubscribe function. */
  on<K extends ViewerEventName>(event: K, handler: ViewerEventHandler<K>): () => void;
  off<K extends ViewerEventName>(event: K, handler: ViewerEventHandler<K>): void;
  /** Current zoom (1 = 100%), or `null` if the current file cannot be zoomed. */
  getZoom(): number | null;
  setZoom(zoom: number): void;
  zoomIn(): void;
  zoomOut(): void;
  /** Current page and page count, or `null` if the current file has no pages. */
  getPage(): { page: number; pageCount: number } | null;
  goToPage(page: number): void;
  /** Saves the current file through the browser's download flow. */
  download(): Promise<void>;
  print(): void;
  toggleFullscreen(): Promise<void>;
  setTheme(theme: ViewerTheme): void;
  /** Removes all DOM, listeners and object URLs. The viewer cannot be reused. */
  destroy(): void;
}
