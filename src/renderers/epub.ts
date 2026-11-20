import { toAnyViewError } from '../core/errors';
import type { Renderer } from '../core/types';
import { openEpub } from '../formats/epub';
import { clear, h } from '../ui/dom';
import { createSandboxedFrame } from '../ui/frame';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const { container, file } = context;
    const book = await openEpub(openZip(await file.bytes()));
    const holder = h('div', { style: { flex: '1 1 auto', 'min-height': '0', display: 'flex' } });
    container.appendChild(holder);
    let chapter = 0;
    let zoom = 1;
    let destroyed = false;
    let request = 0;

    const show = async (): Promise<void> => {
      const current = ++request;
      let frame: HTMLElement;
      try {
        const html = await book.chapterHtml(chapter, zoom);
        frame = createSandboxedFrame(html, book.title || file.name || 'EPUB');
      } catch (error) {
        frame = h('p', {
          class: 'av-status-text',
          style: { padding: '24px', margin: '0 auto' },
          text: toAnyViewError(error, 'PARSE_FAILED', 'This chapter could not be displayed')
            .message,
        });
      }
      if (destroyed || current !== request) return;
      clear(holder);
      holder.appendChild(frame);
    };
    await show();

    return {
      pages: {
        count: () => book.chapterCount,
        current: () => chapter + 1,
        goTo(page: number): void {
          chapter = page - 1;
          void show();
        },
      },
      zoom: {
        get: () => zoom,
        set(value: number): void {
          zoom = value;
          void show();
        },
      },
      destroy(): void {
        destroyed = true;
      },
    };
  },
};
