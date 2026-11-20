import { AnyViewError } from '../core/errors';
import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { toBlobPart } from '../util/text';

const SAMPLES = [
  'ABCDEFGHIJKLMNOPQRSTUVWXYZ',
  'abcdefghijklmnopqrstuvwxyz',
  '0123456789 .,:;!?@#$%&*()[]{}',
];
const PANGRAM = 'The quick brown fox jumps over the lazy dog';
const SIZES = [12, 18, 24, 36, 48, 72];

let fontCounter = 0;

export const renderer: Renderer = {
  async render(context) {
    if (typeof FontFace === 'undefined' || !document.fonts) {
      throw new AnyViewError('RENDER_FAILED', 'This browser cannot preview font files.');
    }
    const family = `anyview-font-${++fontCounter}`;
    let face: FontFace;
    try {
      face = new FontFace(family, toBlobPart(await context.file.bytes()));
      await face.load();
    } catch (cause) {
      throw new AnyViewError('PARSE_FAILED', 'This font file could not be read.', { cause });
    }
    document.fonts.add(face);

    const specimen = h('div', { class: 'av-font' });
    for (const sample of SAMPLES) {
      specimen.appendChild(
        h('p', { text: sample, style: { 'font-family': family, 'font-size': '28px' } }),
      );
    }
    for (const size of SIZES) {
      specimen.appendChild(
        h(
          'p',
          { style: { 'font-family': family, 'font-size': `${size}px` } },
          h('small', { text: `${size}px` }),
          PANGRAM,
        ),
      );
    }
    context.container.appendChild(h('div', { class: 'av-scroll' }, specimen));
    return {
      printable: () => specimen,
      destroy(): void {
        document.fonts.delete(face);
      },
    };
  },
};
