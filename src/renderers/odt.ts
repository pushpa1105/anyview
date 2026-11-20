import type { Renderer } from '../core/types';
import { renderOdt } from '../formats/odt';
import { createPaperView } from '../ui/paper';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const zip = openZip(await context.file.bytes());
    const paper = await renderOdt(zip, context);
    return createPaperView(context, paper);
  },
};
