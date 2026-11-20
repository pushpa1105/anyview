import { toAnyViewError } from '../core/errors';
import type { RenderContext, RendererInstance } from '../core/types';
import { clear, h, observeResize, rafThrottle } from './dom';

export interface SlideDeck {
  /** Slide size in CSS pixels at 100%. */
  width: number;
  height: number;
  count: number;
  /** Builds the slide at `index` (zero-based) as an element of exactly `width` × `height`. */
  renderSlide(index: number): Promise<HTMLElement>;
}

/**
 * Shows slides as a vertical, scrollable list scaled to the available width.
 * Slides are built only when they come near the viewport.
 */
export function createSlideView(context: RenderContext, deck: SlideDeck): RendererInstance {
  const scroller = h('div', { class: 'av-pages' });
  context.container.appendChild(scroller);
  let zoom = 1;
  let currentPage = 1;
  let destroyed = false;
  const frames: HTMLElement[] = [];
  const started = new Set<number>();

  const scale = (): number => {
    const available = (scroller.clientWidth || deck.width + 32) - 32;
    return Math.max(0.05, Math.min(available / deck.width, 2)) * zoom;
  };

  const layout = (): void => {
    const factor = scale();
    for (const frame of frames) {
      frame.style.width = `${deck.width * factor}px`;
      frame.style.height = `${deck.height * factor}px`;
      const slide = frame.firstElementChild as HTMLElement | null;
      if (slide?.classList.contains('av-slide')) slide.style.transform = `scale(${factor})`;
    }
  };

  const build = async (index: number): Promise<void> => {
    const frame = frames[index];
    if (!frame || started.has(index)) return;
    started.add(index);
    try {
      const slide = await deck.renderSlide(index);
      if (destroyed) return;
      slide.classList.add('av-slide');
      slide.style.width = `${deck.width}px`;
      slide.style.height = `${deck.height}px`;
      slide.style.transform = `scale(${scale()})`;
      clear(frame);
      frame.appendChild(slide);
    } catch (error) {
      if (destroyed) return;
      clear(frame);
      frame.appendChild(
        h('p', {
          class: 'av-status-text',
          style: { padding: '24px' },
          text: toAnyViewError(error, 'RENDER_FAILED', `Slide ${index + 1} could not be displayed`)
            .message,
        }),
      );
    }
  };

  for (let index = 0; index < deck.count; index++) {
    const frame = h('div', { class: 'av-page', attrs: { 'data-page': index + 1 } });
    frames.push(frame);
    scroller.appendChild(frame);
  }
  layout();

  let observer: IntersectionObserver | null = null;
  if (typeof IntersectionObserver !== 'undefined') {
    observer = new IntersectionObserver(
      (entries) => {
        for (const entry of entries) {
          if (entry.isIntersecting) void build(frames.indexOf(entry.target as HTMLElement));
        }
      },
      { root: scroller, rootMargin: '800px 0px' },
    );
    for (const frame of frames) observer.observe(frame);
  } else {
    for (let index = 0; index < deck.count; index++) void build(index);
  }

  const stopObserving = observeResize(scroller, layout);
  const trackPage = rafThrottle(() => {
    const middle = scroller.scrollTop + scroller.clientHeight / 2;
    let page = 1;
    for (let index = 0; index < frames.length; index++) {
      if (frames[index]!.offsetTop <= middle) page = index + 1;
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
        zoom = value;
        layout();
      },
    },
    pages: {
      count: () => deck.count,
      current: () => currentPage,
      goTo(page: number): void {
        const frame = frames[page - 1];
        if (!frame) return;
        currentPage = page;
        scroller.scrollTop = frame.offsetTop - 16;
      },
    },
    printable: () => scroller,
    destroy(): void {
      destroyed = true;
      trackPage.cancel();
      stopObserving();
      observer?.disconnect();
    },
  };
}
