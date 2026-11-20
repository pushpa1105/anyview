import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { waitForLoad } from '../ui/media';

export const renderer: Renderer = {
  async render(context) {
    const { file, format, container, signal } = context;
    const isAudio = format.id === 'audio';
    const player = h(isAudio ? 'audio' : 'video', {
      attrs: {
        controls: true,
        preload: 'metadata',
        playsinline: !isAudio,
        'aria-label': file.name,
      },
    });
    container.appendChild(
      h('div', { class: isAudio ? 'av-media av-media-audio' : 'av-media' }, player),
    );
    const ready = waitForLoad(
      player,
      'loadedmetadata',
      signal,
      `This ${isAudio ? 'audio' : 'video'} file uses a format or codec the browser cannot play.`,
    );
    player.src = await context.fileUrl();
    await ready;
    return {
      destroy(): void {
        // Stop playback and release the decoder and network connection.
        player.pause();
        player.removeAttribute('src');
        player.load();
      },
    };
  },
};
