import type { Renderer } from '../core/types';
import { openOdpPresentation } from '../formats/odp';
import { createSlideView } from '../ui/slides';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const zip = openZip(await context.file.bytes());
    const deck = await openOdpPresentation(zip, context);
    return createSlideView(context, deck);
  },
};
