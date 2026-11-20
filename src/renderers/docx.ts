import type { Renderer } from '../core/types';
import { renderDocx } from '../formats/docx';
import { createPaperView } from '../ui/paper';
import { openZip } from '../util/zip';

export const renderer: Renderer = {
  async render(context) {
    const zip = openZip(await context.file.bytes());
    const document = await renderDocx(zip, context);
    if (context.options.tableGridlines) document.element.classList.add('av-doc-gridlines');
    const view = createPaperView(context, document.element, document);
    // Tab stops and page breaks need measurements, so they are computed once the pages are attached.
    await document.layout();
    context.notify('page');
    return {
      ...view,
      destroy(): void {
        document.destroy();
      },
    };
  },
};
