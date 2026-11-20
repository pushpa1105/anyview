# Custom renderers

A renderer turns a file into DOM inside the viewer. All built-in formats are renderers, and you can add your own to support new formats or to replace a built-in one.

## Anatomy

```ts
import type { RendererDefinition } from '@buffmomo/anyview';

export const myRenderer: RendererDefinition = {
  // A unique name.
  id: 'my-renderer',

  // What it handles. Use any combination of the three:
  formats: ['csv'], //            built-in format ids to take over
  extensions: ['dwg', 'dxf'], //  new file extensions (without the dot)
  mimeTypes: ['image/vnd.dwg'], // new MIME types

  // Called once, the first time a matching file is opened. Import heavy code here.
  load: async () => (await import('./my-renderer-impl')).renderer,
};
```

```ts
// my-renderer-impl.ts
import type { Renderer } from '@buffmomo/anyview';

export const renderer: Renderer = {
  async render(context) {
    const text = await context.file.text();
    const pre = document.createElement('pre');
    pre.textContent = text; // never innerHTML with file content
    context.container.appendChild(pre);

    return {
      destroy() {
        // remove listeners, stop timers, terminate workers…
      },
    };
  },
};
```

Register it when creating the viewer:

```ts
createViewer('#preview', { renderers: [myRenderer] });
```

Custom renderers are consulted before built-in ones, in the order given.

## The render context

| Member                                                | Description                                                                                                                                           |
| ----------------------------------------------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------- |
| `file.name`, `file.mimeType`, `file.size`, `file.url` | Metadata. `size` is `null` for a remote file that has not been downloaded yet.                                                                        |
| `file.blob()`, `file.bytes()`, `file.text()`          | The content, fetched on first use and cached. `text()` decodes UTF-8 or UTF-16 (by byte-order mark), falling back to Windows-1252.                    |
| `format`                                              | The detected format (`id`, `label`, `category`…). For new extensions this is a synthetic entry named after your renderer.                             |
| `language`                                            | A syntax-highlighting hint for text formats, or `null`.                                                                                               |
| `container`                                           | An empty element to fill. It is a flex column that takes all available space.                                                                         |
| `signal`                                              | An `AbortSignal` that fires when the file is closed or replaced. Pass it to `fetch` and check it after long awaits.                                   |
| `options`, `labels`                                   | The viewer's resolved options and UI texts.                                                                                                           |
| `objectUrl(blob)`                                     | Like `URL.createObjectURL`, but revoked automatically when the file is closed.                                                                        |
| `fileUrl()`                                           | A URL for the whole file: the original URL when the browser can load it directly, otherwise an object URL. Useful for `<img>`, `<video>` and similar. |
| `notify('zoom' \| 'page')`                            | Tell the viewer that zoom or page changed from inside the renderer (for example by scrolling), so the toolbar and events update.                      |
| `openChild(blob, name)`                               | Open a file contained in this one (an archive entry, an attachment) in a nested viewer with a Back button.                                            |

## What `render` can return

Everything is optional. Return nothing for a static view.

```ts
return {
  // Called when the file is closed. Release everything you created.
  destroy() {},

  // Enables the zoom buttons and viewer.setZoom(). 1 means 100%.
  zoom: {
    get: () => currentZoom,
    set: (value) => applyZoom(value),
  },

  // Enables page navigation and viewer.goToPage(). Pages start at 1.
  pages: {
    count: () => pageCount,
    current: () => currentPage,
    goTo: (page) => showPage(page),
  },

  // Printing: either take over completely…
  print() {},
  // …or provide the node that should be printed (default: the container's content).
  printable: () => someElement,
};
```

The zoom level returned by `zoom.get()` right after rendering is what the toolbar's reset button goes back to.

## Errors

Throw from `render` (or reject) to show the error card. Throw an `AnyViewError` to control the code and message; anything else is wrapped as `RENDER_FAILED`.

```ts
import { AnyViewError } from '@buffmomo/anyview';

throw new AnyViewError('PARSE_FAILED', 'This drawing uses a DWG version that is not supported.', {
  details: { version },
});
```

## Example: replacing the CSV renderer

```ts
import { createViewer, type RendererDefinition } from '@buffmomo/anyview';

const chartCsv: RendererDefinition = {
  id: 'csv-chart',
  formats: ['csv'],
  load: () => ({
    async render({ file, container }) {
      const rows = (await file.text())
        .trim()
        .split('\n')
        .map((line) => line.split(','));
      const list = document.createElement('ul');
      for (const [label, value] of rows.slice(1)) {
        const item = document.createElement('li');
        item.textContent = `${label}: ${'#'.repeat(Number(value) || 0)}`;
        list.appendChild(item);
      }
      container.appendChild(list);
    },
  }),
};

createViewer('#preview', { src: '/data/votes.csv', renderers: [chartCsv] });
```

## Example: adding legacy Excel through a library you choose

AnyView does not read `.xls`. If your application needs it, you can plug in a library of your choice without AnyView depending on it:

```ts
const legacyExcel: RendererDefinition = {
  id: 'xls',
  extensions: ['xls'],
  load: async () => {
    const XLSX = await import('your-spreadsheet-library');
    return {
      async render({ file, container }) {
        const workbook = XLSX.read(await file.bytes());
        // Build a <table> from the first sheet with createElement/textContent…
      },
    };
  },
};
```

## Guidelines

- **Build DOM, never inject markup.** Use `createElement` and `textContent`. If you must show HTML from the file, put it in an `<iframe sandbox>`.
- **Clean up in `destroy`.** The same viewer may open hundreds of files in a session.
- **Respect `signal`.** After each long `await`, return early if `signal.aborted` is true.
- **Load lazily.** Keep `load` as a dynamic `import()` so your code is only downloaded when needed.
- **Use the viewer's CSS variables** (`var(--av-fg)`, `var(--av-surface)`, `var(--av-border)`…) so your renderer follows the light and dark themes.
