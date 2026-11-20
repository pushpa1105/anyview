import type { Renderer } from '../core/types';
import { renderRtf } from '../formats/rtf';
import { createPaperView } from '../ui/paper';

export const renderer: Renderer = {
  async render(context) {
    const paper = renderRtf(await context.file.bytes(), context);
    return createPaperView(context, paper);
  },
};
