import type { Renderer } from '../core/types';
import { openOdsWorkbook } from '../formats/ods';
import { createSheetView } from '../ui/sheets';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const zip = openZip(await context.file.bytes());
    const sheets = await openOdsWorkbook(zip);
    return createSheetView(context, sheets);
  },
};
