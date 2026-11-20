import { AnyViewError, isAnyViewError } from '../core/errors';
import type {
  PdfJsDocument,
  PdfJsLike,
  RenderContext,
  Renderer,
  RendererInstance,
} from '../core/types';
import { clear, h, rafThrottle } from '../ui/dom';

/** CSS pixels per PDF point. */
const POINT_TO_PIXEL = 96 / 72;

function isCrossOrigin(url: string): boolean {
  try {
    return new URL(url).origin !== window.location.origin;
  } catch {
    return false;
  }
}

function hasPdfHeader(bytes: Uint8Array): boolean {
  // The spec allows junk before the header, readers accept it within the first 1024 bytes.
  const limit = Math.min(bytes.length - 4, 1024);
  for (let i = 0; i <= limit; i++) {
    if (
      bytes[i] === 0x25 &&
      bytes[i + 1] === 0x50 &&
      bytes[i + 2] === 0x44 &&
      bytes[i + 3] === 0x46
    ) {
      return true;
    }
  }
  return false;
}

/** Shows the PDF with the browser's built-in viewer. */
async function renderNative(context: RenderContext): Promise<RendererInstance> {
  const { file, container } = context;
  let url: string;
  let sameOrigin = true;
  try {
    const bytes = await file.bytes();
    if (!hasPdfHeader(bytes)) {
      throw new AnyViewError('PARSE_FAILED', 'This file is not a valid PDF document.');
    }
    // Typing the blob guarantees the browser treats the content as PDF and nothing else.
    url = context.objectUrl(new Blob([await file.blob()], { type: 'application/pdf' }));
  } catch (error) {
    // A cross-origin PDF without CORS headers cannot be fetched, but it can still be framed.
    const blocked = isAnyViewError(error) && error.code === 'FETCH_FAILED';
    if (!blocked || !file.direct || !file.url || !isCrossOrigin(file.url)) throw error;
    url = file.url;
    sameOrigin = false;
  }
  const frame = h('iframe', {
    class: 'av-frame',
    attrs: { src: url, title: file.name || 'PDF document' },
  });
  container.appendChild(frame);
  return {
    print(): void {
      try {
        if (!sameOrigin || !frame.contentWindow) throw new Error('not printable in place');
        frame.contentWindow.focus();
        frame.contentWindow.print();
      } catch {
        window.open(url, '_blank', 'noopener');
      }
    },
  };
}

interface PageSlot {
  holder: HTMLElement;
  /** Zoom the canvas was drawn at, or 0 when it needs drawing. */
  drawnAt: number;
  cancel: (() => void) | null;
}

/** Draws the PDF to canvases through a PDF.js module supplied by the application. */
async function renderWithPdfJs(
  context: RenderContext,
  pdfjs: PdfJsLike,
): Promise<RendererInstance> {
  const { file, container, signal } = context;
  if (typeof pdfjs?.getDocument !== 'function') {
    throw new AnyViewError('INVALID_OPTIONS', 'The "pdfjs" option must be the pdfjs-dist module.');
  }
  const bytes = await file.bytes();
  let pdf: PdfJsDocument;
  let release: () => void = () => undefined;
  try {
    // PDF.js takes ownership of the buffer it is given, so hand it a copy.
    const task = pdfjs.getDocument({ data: bytes.slice() });
    pdf = await task.promise;
    const opened = pdf;
    // PDF.js 5+ releases a document through its loading task, older versions through the document.
    release = () => {
      const pending = task.destroy ? task.destroy() : opened.destroy?.();
      if (pending) pending.catch(() => undefined);
    };
  } catch (cause) {
    throw new AnyViewError('PARSE_FAILED', 'This PDF could not be opened.', { cause });
  }
  if (signal.aborted) {
    release();
    throw new AnyViewError('ABORTED', 'The operation was aborted.');
  }

  const scroller = h('div', { class: 'av-pages' });
  container.appendChild(scroller);
  let zoom = 1;
  let destroyed = false;
  let currentPage = 1;

  const first = await pdf.getPage(1);
  const base = first.getViewport({ scale: POINT_TO_PIXEL });
  const slots: PageSlot[] = [];
  for (let index = 0; index < pdf.numPages; index++) {
    const holder = h('div', {
      class: 'av-page',
      attrs: { 'data-page': index + 1 },
      style: { width: `${base.width}px`, height: `${base.height}px` },
    });
    scroller.appendChild(holder);
    slots.push({ holder, drawnAt: 0, cancel: null });
  }

  const draw = async (index: number): Promise<void> => {
    const slot = slots[index];
    if (!slot || destroyed || slot.drawnAt === zoom) return;
    const target = zoom;
    slot.drawnAt = target;
    slot.cancel?.();
    try {
      const page = await pdf.getPage(index + 1);
      if (destroyed || zoom !== target) return;
      const ratio = Math.min(window.devicePixelRatio || 1, 2);
      const viewport = page.getViewport({ scale: POINT_TO_PIXEL * target * ratio });
      const canvas = h('canvas', {
        style: {
          width: `${viewport.width / ratio}px`,
          height: `${viewport.height / ratio}px`,
          display: 'block',
        },
      });
      canvas.width = Math.ceil(viewport.width);
      canvas.height = Math.ceil(viewport.height);
      const canvasContext = canvas.getContext('2d');
      if (!canvasContext) return;
      const task = page.render({ canvasContext, viewport, canvas });
      slot.cancel = () => task.cancel?.();
      await task.promise;
      slot.cancel = null;
      if (destroyed || zoom !== target) return;
      slot.holder.style.width = `${viewport.width / ratio}px`;
      slot.holder.style.height = `${viewport.height / ratio}px`;
      clear(slot.holder);
      slot.holder.appendChild(canvas);
    } catch {
      // Cancelled or failed: allow another attempt next time the page scrolls into view.
      if (slot.drawnAt === target) slot.drawnAt = 0;
    }
  };

  // Only pages near the viewport are drawn, so a 1 000-page PDF opens instantly.
  let observer: IntersectionObserver | null = null;
  if (typeof IntersectionObserver !== 'undefined') {
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting)
            void draw(slots.findIndex((slot) => slot.holder === entry.target));
        }
      },
      { root: scroller, rootMargin: '600px 0px' },
    );
    for (const slot of slots) observer.observe(slot.holder);
  } else {
    for (let index = 0; index < Math.min(slots.length, 20); index++) void draw(index);
  }

  const trackPage = rafThrottle(() => {
    const middle = scroller.scrollTop + scroller.clientHeight / 2;
    let page = 1;
    for (let index = 0; index < slots.length; index++) {
      if (slots[index]!.holder.offsetTop <= middle) page = index + 1;
      else break;
    }
    if (page !== currentPage) {
      currentPage = page;
      context.notify('page');
    }
  });
  scroller.addEventListener('scroll', trackPage, { passive: true });

  return {
    zoom: {
      get: () => zoom,
      set(value: number): void {
        const ratio = value / zoom;
        zoom = value;
        for (const slot of slots) {
          slot.holder.style.width = `${parseFloat(slot.holder.style.width) * ratio}px`;
          slot.holder.style.height = `${parseFloat(slot.holder.style.height) * ratio}px`;
          // Re-observing makes the observer report visible pages again for redrawing.
          observer?.unobserve(slot.holder);
          observer?.observe(slot.holder);
        }
        if (!observer)
          for (let index = 0; index < Math.min(slots.length, 20); index++) void draw(index);
      },
    },
    pages: {
      count: () => slots.length,
      current: () => currentPage,
      goTo(page: number): void {
        const slot = slots[page - 1];
        if (!slot) return;
        currentPage = page;
        scroller.scrollTop = slot.holder.offsetTop - 16;
      },
    },
    printable: () => scroller,
    destroy(): void {
      destroyed = true;
      trackPage.cancel();
      observer?.disconnect();
      for (const slot of slots) slot.cancel?.();
      release();
    },
  };
}

export const renderer: Renderer = {
  async render(context) {
    const provided = context.options.pdfjs;
    if (!provided) return renderNative(context);
    let pdfjs: PdfJsLike;
    try {
      pdfjs = typeof provided === 'function' ? await provided() : provided;
    } catch (cause) {
      throw new AnyViewError('RENDER_FAILED', 'The PDF.js module could not be loaded.', { cause });
    }
    return renderWithPdfJs(context, pdfjs);
  },
};
