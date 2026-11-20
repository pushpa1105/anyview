import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { waitForLoad } from '../ui/media';

export const renderer: Renderer = {
  async render(context) {
    const { file, format, container, signal } = context;
    let url: string;
    if (format.id === 'svg' && !(file.direct && file.url)) {
      // An SVG only renders in <img> when its blob is typed; inside <img> its scripts never run.
      url = context.objectUrl(new Blob([await file.blob()], { type: 'image/svg+xml' }));
    } else {
      url = await context.fileUrl();
    }

    const image = h('img', { class: 'av-fit', attrs: { alt: file.name, draggable: 'false' } });
    const stage = h('div', { class: 'av-image' }, image);
    container.appendChild(stage);
    const loaded = waitForLoad(image, 'load', signal, 'The image could not be decoded.');
    image.src = url;
    await loaded;

    // SVGs without intrinsic dimensions report a natural size of 0.
    const naturalWidth = image.naturalWidth || 800;
    const naturalHeight = image.naturalHeight || 600;
    const availableWidth = stage.clientWidth - 32;
    const availableHeight = stage.clientHeight - 32;
    let zoom =
      availableWidth > 0 && availableHeight > 0
        ? Math.min(1, availableWidth / naturalWidth, availableHeight / naturalHeight)
        : 1;

    return {
      zoom: {
        get: () => zoom,
        set(value: number): void {
          zoom = value;
          image.classList.remove('av-fit');
          image.style.width = `${naturalWidth * zoom}px`;
          image.style.height = image.naturalHeight ? 'auto' : `${naturalHeight * zoom}px`;
        },
      },
      printable: () => image,
    };
  },
};
