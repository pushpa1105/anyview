import type { Renderer } from '../core/types';
import { openPresentation } from '../formats/pptx';
import { createSlideView } from '../ui/slides';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const zip = openZip(await context.file.bytes());
    const deck = await openPresentation(zip, context);
    return createSlideView(context, deck);
  },
};
