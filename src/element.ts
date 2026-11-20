import { createViewer } from './core/viewer';
import type {
  ToolbarOptions,
  Viewer,
  ViewerOptions,
  ViewerSource,
  ViewerTheme,
} from './core/types';

/** The `<any-view>` custom element's public surface. */
export interface AnyViewElement extends HTMLElement {
  /** File to show: a URL, `File`, `Blob` or bytes. Mirrors the `src` attribute for strings. */
  source: ViewerSource | null;
  /** Options that cannot be expressed as attributes (renderers, pdfjs, labels, fetchOptions…). */
  options: ViewerOptions;
  /** The underlying viewer, available while the element is connected. */
  readonly viewer: Viewer | null;
}

const OBSERVED = ['src', 'name', 'type', 'theme', 'toolbar'] as const;

/**
 * Registers the `<any-view>` custom element.
 *
 * @example
 * <any-view src="/files/report.pdf" style="height: 600px"></any-view>
 *
 * Events: `av-loadstart`, `av-load`, `av-error`, `av-pagechange`, `av-zoomchange`
 * (the payload is in `event.detail`).
 */
export function defineAnyViewElement(tagName = 'any-view'): void {
  if (typeof customElements === 'undefined' || customElements.get(tagName)) return;

  // Declared inside the function so that importing this module during SSR,
  // where HTMLElement does not exist, is harmless.
  class AnyView extends HTMLElement implements AnyViewElement {
    static readonly observedAttributes = OBSERVED;
    #viewer: Viewer | null = null;
    #source: ViewerSource | null = null;
    #options: ViewerOptions = {};
    #scheduled = false;

    get viewer(): Viewer | null {
      return this.#viewer;
    }

    get source(): ViewerSource | null {
      return this.#source ?? this.getAttribute('src');
    }

    set source(value: ViewerSource | null) {
      this.#source = value;
      this.#reload();
    }

    get options(): ViewerOptions {
      return this.#options;
    }

    set options(value: ViewerOptions) {
      this.#options = value ?? {};
      this.#rebuild();
    }

    connectedCallback(): void {
      if (!this.style.display) this.style.display = 'block';
      this.#rebuild();
    }

    disconnectedCallback(): void {
      this.#viewer?.destroy();
      this.#viewer = null;
    }

    attributeChangedCallback(name: string, previous: string | null, next: string | null): void {
      if (previous === next || !this.#viewer) return;
      if (name === 'theme') {
        this.#viewer.setTheme(this.#theme());
      } else if (name === 'toolbar') {
        this.#rebuild();
      } else {
        if (name === 'src') this.#source = null;
        this.#reload();
      }
    }

    #theme(): ViewerTheme {
      const value = this.getAttribute('theme');
      return value === 'light' || value === 'dark' ? value : 'auto';
    }

    #toolbar(): boolean | ToolbarOptions | undefined {
      const value = this.getAttribute('toolbar');
      if (value === null) return this.#options.toolbar;
      return value !== 'false' && value !== 'off' && value !== 'none';
    }

    /** Recreates the viewer; needed when construction-time options change. */
    #rebuild(): void {
      if (!this.isConnected) return;
      this.#viewer?.destroy();
      const viewer = createViewer(this, {
        ...this.#options,
        src: null,
        theme: this.#theme(),
        toolbar: this.#toolbar() ?? true,
      });
      this.#viewer = viewer;
      const forward = (type: string) => (detail: unknown) =>
        this.dispatchEvent(new CustomEvent(`av-${type}`, { detail }));
      viewer.on('loadstart', forward('loadstart'));
      viewer.on('load', forward('load'));
      viewer.on('error', forward('error'));
      viewer.on('pagechange', forward('pagechange'));
      viewer.on('zoomchange', forward('zoomchange'));
      this.#reload();
    }

    /** Batches attribute/property changes made in the same tick into one load. */
    #reload(): void {
      if (!this.#viewer || this.#scheduled) return;
      this.#scheduled = true;
      queueMicrotask(() => {
        this.#scheduled = false;
        const viewer = this.#viewer;
        if (!viewer) return;
        const source = this.source ?? this.#options.src ?? null;
        if (source === null || source === '') {
          viewer.clear();
          return;
        }
        const name = this.getAttribute('name') ?? this.#options.name;
        const type = this.getAttribute('type') ?? this.#options.type;
        // Errors are reported through the av-error event and the viewer's own UI.
        viewer
          .load(source, { ...(name ? { name } : {}), ...(type ? { type } : {}) })
          .catch(() => undefined);
      });
    }
  }

  customElements.define(tagName, AnyView);
}

defineAnyViewElement();

declare global {
  interface HTMLElementTagNameMap {
    'any-view': AnyViewElement;
  }
}
