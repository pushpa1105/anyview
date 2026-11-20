import type { PageCapability, RenderContext, RendererInstance } from '../core/types';
import { h } from './dom';

export interface PagedDocument {
  /** Number of pages currently laid out. */
  pageCount(): number;
  /** Element of a page (zero-based). */
  page(index: number): HTMLElement | null;
}

/**
 * Shows a word-processor style document: white "sheets of paper" centred in a
 * scrolling area. Used by DOCX (a stack of pages), ODT and RTF (one sheet).
 */
export function createPaperView(
  context: RenderContext,
  paper: HTMLElement,
  paged?: PagedDocument,
): RendererInstance {
  // A page stack styles its own pages; a single sheet is the paper itself.
  if (!paper.classList.contains('av-doc')) paper.classList.add('av-paper');
  const scroller = h('div', { class: 'av-pages' }, paper);
  context.container.appendChild(scroller);
  let zoom = 1;
  let pages: PageCapability | undefined;

  if (paged) {
    /** The page that crosses the middle of the visible area. */
    const current = (): number => {
      const middle = scroller.getBoundingClientRect().top + scroller.clientHeight / 2;
      const count = paged.pageCount();
      for (let index = 0; index < count; index++) {
        const page = paged.page(index);
        if (page && page.getBoundingClientRect().bottom >= middle) return index + 1;
      }
      return Math.max(1, count);
    };
    let last = 1;
    scroller.addEventListener(
      'scroll',
      () => {
        const now = current();
        if (now !== last) {
          last = now;
          context.notify('page');
        }
      },
      { passive: true },
    );
    pages = {
      count: () => paged.pageCount(),
      current,
      goTo(page: number): void {
        const target = paged.page(Math.max(1, Math.min(page, paged.pageCount())) - 1);
        if (!target) return;
        scroller.scrollTop +=
          target.getBoundingClientRect().top - scroller.getBoundingClientRect().top - 16;
        context.notify('page');
      },
    };
  }

  const instance: RendererInstance = {
    zoom: {
      get: () => zoom,
      set(value: number): void {
        zoom = value;
        // `zoom` scales layout as well as paint, so scrollbars stay correct.
        paper.style.setProperty('zoom', String(value));
      },
    },
    printable: () => paper,
  };
  if (pages) instance.pages = pages;
  return instance;
}
