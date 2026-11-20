import type { Renderer } from '../core/types';
import { openWorkbook } from '../formats/xlsx';
import { createSheetView } from '../ui/sheets';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const zip = openZip(await context.file.bytes());
    const sheets = await openWorkbook(zip, context);
    return createSheetView(context, sheets);
  },
};
