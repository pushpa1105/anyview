import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { renderMarkdown } from '../util/markdown';

/** Markdown longer than this is displayed as text; parsing it would freeze the page. */
const MAX_RENDERED_CHARS = 1_500_000;

export const renderer: Renderer = {
  async render(context) {
    const { file, container } = context;
    const source = await file.text();
    if (source.length > MAX_RENDERED_CHARS) {
      // Very large files are shown as plain text rather than parsed on the main thread.
      return (await import('./text')).renderer.render(context);
    }
    const prose = h('div', { class: 'av-prose' });
    prose.appendChild(renderMarkdown(source, { baseUrl: file.url }));
    container.appendChild(h('div', { class: 'av-scroll' }, prose));
    let zoom = 1;
    return {
      zoom: {
        get: () => zoom,
        set(value: number): void {
          zoom = value;
          prose.style.fontSize = `${14 * zoom}px`;
        },
      },
      printable: () => prose,
    };
  },
};
