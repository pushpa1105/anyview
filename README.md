# @buffmomo/anyview

[![npm version](https://img.shields.io/npm/v/@buffmomo/anyview)](https://www.npmjs.com/package/@buffmomo/anyview)
[![CI](https://github.com/pushpa1105/anyview/actions/workflows/ci.yml/badge.svg)](https://github.com/pushpa1105/anyview/actions/workflows/ci.yml)
[![npm downloads](https://img.shields.io/npm/dm/@buffmomo/anyview)](https://www.npmjs.com/package/@buffmomo/anyview)
[![types included](https://img.shields.io/npm/types/@buffmomo/anyview)](https://www.npmjs.com/package/@buffmomo/anyview)
[![zero dependencies](https://img.shields.io/badge/dependencies-0-brightgreen)](package.json)
[![license](https://img.shields.io/npm/l/@buffmomo/anyview)](LICENSE)

**[Try the playground](https://pushpa1105.github.io/anyview/)** · [Formats](docs/formats.md) · [Security](docs/security.md) · [Changelog](CHANGELOG.md)

**One viewer for every file.** AnyView previews documents, spreadsheets, presentations, PDFs, images, video, audio, code and more, entirely in the browser. It works with plain JavaScript, as a web component, and has thin wrappers for React and Vue.

- **Zero runtime dependencies.** The ZIP reader, XML parser, DOCX/XLSX/PPTX renderers, Markdown parser and syntax highlighter are all written in this package.
- **Private by design.** Files are parsed on the user's device. Nothing is uploaded to a conversion service.
- **Small.** About 22 kB gzipped to start; each format's code (3–22 kB) is downloaded only when a file of that type is opened.
- **Safe with untrusted files.** File content is never inserted as markup into your page; HTML is shown in a locked-down sandbox.

```js
import { createViewer } from '@buffmomo/anyview';

createViewer('#preview', { src: '/files/quarterly-report.xlsx' });
```

## Contents

- [Supported formats](#supported-formats)
- [Installation](#installation)
- [Requirements](#requirements)
- [Quick start](#quick-start) (JavaScript, web component, React, Vue)
- [Sources](#sources)
- [Options](#options)
- [Viewer API](#viewer-api)
- [Events](#events)
- [Error handling](#error-handling)
- [TypeScript](#typescript)
- [Advanced usage](#advanced-usage) (PDF.js, custom renderers, theming, translation, authentication)
- [Security](#security)
- [Limits and known limitations](#limits-and-known-limitations)
- [Troubleshooting](#troubleshooting)
- [Contributing](#contributing)
- [Author and contributors](#author-and-contributors)
- [License](#license)

## Supported formats

| Kind                | Formats                                                             | What you get                                                                                                           |
| ------------------- | ------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------- |
| Documents           | DOCX, ODT, RTF                                                      | DOCX laid out in pages like Word: headers, footers, footnotes, table styles, floating pictures, shapes, tab stops      |
| Spreadsheets        | XLSX, ODS, CSV, TSV                                                 | XLSX with fonts, borders, table styles, conditional formatting, pictures and charts; smooth scrolling for 100 000 rows |
| Presentations       | PPTX, ODP                                                           | Slides with text, preset and custom shapes, pictures, tables, charts, backgrounds and theme colours                    |
| PDF                 | PDF                                                                 | The browser's built-in viewer, or [PDF.js](#pdf-rendering-with-pdfjs) if you provide it                                |
| Images              | PNG, JPEG, GIF, WebP, AVIF, BMP, ICO, SVG                           | Zoom and scroll                                                                                                        |
| Media               | MP4, WebM, OGV, MOV, MP3, WAV, OGG, M4A, AAC, FLAC, Opus            | Native player (subject to the codecs the browser supports)                                                             |
| Text and code       | TXT, LOG, JSON, XML, YAML, TOML, INI and 60+ source-code extensions | Line numbers and syntax highlighting; minified JSON is pretty-printed                                                  |
| Web                 | HTML, Markdown                                                      | HTML in a sandbox; Markdown with tables, task lists and highlighted code blocks                                        |
| Books and notebooks | EPUB, Jupyter (`.ipynb`)                                            | Chapter navigation; notebook cells with outputs                                                                        |
| Email               | EML                                                                 | Headers, HTML or text body, attachments you can open                                                                   |
| Archives            | ZIP, JAR                                                            | File listing; click a file to preview it inside the archive                                                            |
| Fonts               | TTF, OTF, WOFF, WOFF2                                               | Specimen sheet                                                                                                         |

Files AnyView recognises but cannot display (legacy `.doc`, `.xls`, `.ppt`, Outlook `.msg`, HEIC, TIFF) and unknown files show a clear message with a **Download** button instead of a blank box.

More detail per format: [docs/formats.md](docs/formats.md).

## Installation

```bash
npm install @buffmomo/anyview
```

```bash
pnpm add @buffmomo/anyview
yarn add @buffmomo/anyview
```

No other packages are required. React and Vue are optional peer dependencies, needed only if you import the matching wrapper.

Without a build step, load it as an ES module from any npm CDN:

```html
<script type="module">
  import { createViewer } from 'https://esm.sh/@buffmomo/anyview';
  createViewer('#preview', { src: '/files/report.pdf' });
</script>
```

## Requirements

- **Browsers:** Chrome and Edge 80+, Firefox 113+, Safari 16.4+. The floor is set by `DecompressionStream`, which Office, OpenDocument, EPUB and ZIP files need. Other formats also work in older browsers that support ES2022.
- **Module formats:** ESM and CommonJS builds with TypeScript declarations for both.
- **Server-side rendering:** importing the package on a server is safe (nothing touches `window` at import time). Create the viewer in the browser, for example inside `useEffect` or `onMounted`; the React and Vue wrappers already do this.
- **Node.js:** 18 or newer for tooling. AnyView is a browser library and does not render in Node.

## Quick start

The viewer fills its container, so **give the container a height**.

### Plain JavaScript

```html
<div id="preview" style="height: 600px"></div>

<script type="module">
  import { createViewer } from '@buffmomo/anyview';

  const viewer = createViewer('#preview', { src: '/files/report.docx' });

  viewer.on('load', ({ format, file }) => console.log(`Opened ${file.name} as ${format.label}`));
  viewer.on('error', (error) => console.warn(error.code, error.message));
</script>
```

Open a file the user picked:

```js
const viewer = createViewer('#preview');

document.querySelector('input[type=file]').addEventListener('change', (event) => {
  const [file] = event.target.files;
  if (file) viewer.load(file).catch(() => {}); // the viewer already shows the error
});
```

### Web component

Works in plain HTML and in any framework (Angular, Svelte, Solid, Lit…).

```html
<script type="module">
  import '@buffmomo/anyview/element';
</script>

<any-view src="/files/slides.pptx" theme="dark" style="height: 600px"></any-view>
```

| Attribute | Meaning                                             |
| --------- | --------------------------------------------------- |
| `src`     | URL of the file                                     |
| `name`    | File name override                                  |
| `type`    | Force a format (`docx`, `.csv`, `application/pdf`…) |
| `theme`   | `light`, `dark` or `auto` (default)                 |
| `toolbar` | `false` hides the toolbar                           |

| Property  | Meaning                                                                                              |
| --------- | ---------------------------------------------------------------------------------------------------- |
| `source`  | Any [source](#sources): a URL, `File`, `Blob` or bytes                                               |
| `options` | Any [option](#options) that cannot be an attribute (`renderers`, `pdfjs`, `labels`, `fetchOptions`…) |
| `viewer`  | The underlying [viewer](#viewer-api) (read-only)                                                     |

Events: `av-loadstart`, `av-load`, `av-error`, `av-pagechange`, `av-zoomchange`. The payload is in `event.detail`.

```js
const element = document.querySelector('any-view');
element.source = fileFromAnInput;
element.addEventListener('av-error', (event) => console.warn(event.detail.code));
```

### React

```tsx
import { useState } from 'react';
import { AnyView } from '@buffmomo/anyview/react';

export function Preview() {
  const [file, setFile] = useState<File | null>(null);
  return (
    <>
      <input type="file" onChange={(event) => setFile(event.target.files?.[0] ?? null)} />
      <AnyView
        src={file}
        style={{ height: 600 }}
        onLoad={({ format }) => console.log('Opened a', format.label)}
        onError={(error) => console.warn(error.code)}
      />
    </>
  );
}
```

`src`, `name`, `type` and `theme` are reactive. Other options (including `tableGridlines`) are read when the component mounts; change the component's `key` to apply new ones. A `ref` gives you `{ viewer }` for calling the [viewer API](#viewer-api). Callbacks: `onLoadStart`, `onLoad`, `onError`, `onPageChange`, `onZoomChange`.

### Vue 3

```vue
<script setup lang="ts">
import { ref } from 'vue';
import { AnyView } from '@buffmomo/anyview/vue';

const file = ref<File | null>(null);
const pick = (event: Event) => (file.value = (event.target as HTMLInputElement).files?.[0] ?? null);
</script>

<template>
  <input type="file" @change="pick" />
  <AnyView :src="file" style="height: 600px" @load="({ format }) => console.log(format.label)" />
</template>
```

Props: `src`, `name`, `type`, `theme`, `toolbar`, `options`. Events: `loadstart`, `load`, `error`, `pagechange`, `zoomchange`. A template ref exposes `viewer`.

Runnable examples for each setup are in [examples/](examples/).

## Sources

Anywhere a file is expected you can pass:

| Source                     | Notes                                                                                                                                                                                                         |
| -------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| URL (`string` or `URL`)    | Relative URLs resolve against the page. Images, video and audio are streamed by the browser; other formats are fetched (the server must allow it through CORS when on another origin). `data:` URLs work too. |
| `File`                     | For example from `<input type="file">` or drag and drop.                                                                                                                                                      |
| `Blob`                     | Pass a `name` or `type` so the format can be identified, or let AnyView detect it from the content.                                                                                                           |
| `ArrayBuffer`, typed array | Same as `Blob`.                                                                                                                                                                                               |

**Format detection** goes in this order: the `type` you give, the file extension, the MIME type, then the file's first bytes. A DOCX served without an extension as `application/octet-stream` is still opened as a document.

## Options

```ts
createViewer(container, options?)
```

`container` is an element or a CSS selector.

| Option           | Type                                     | Default         | Description                                                                                                                                                                                                                   |
| ---------------- | ---------------------------------------- | --------------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `src`            | [source](#sources)                       | none            | File to open immediately.                                                                                                                                                                                                     |
| `name`           | `string`                                 | from the source | File name, used for detection, the toolbar and downloads.                                                                                                                                                                     |
| `type`           | `string`                                 | detected        | Force a format: an id (`"docx"`), extension (`".docx"`) or MIME type.                                                                                                                                                         |
| `toolbar`        | `boolean \| ToolbarOptions`              | `true`          | `false` hides the toolbar. An object switches individual controls: `fileName`, `zoom`, `pagination`, `download`, `print`, `fullscreen`.                                                                                       |
| `theme`          | `'light' \| 'dark' \| 'auto'`            | `'auto'`        | `auto` follows the operating system.                                                                                                                                                                                          |
| `maxFileSize`    | `number` (bytes)                         | 200 MB          | Larger files are rejected with `FILE_TOO_LARGE`.                                                                                                                                                                              |
| `fetchOptions`   | `RequestInit`                            | none            | Passed to `fetch` for URL sources (headers, credentials…).                                                                                                                                                                    |
| `renderers`      | `RendererDefinition[]`                   | `[]`            | [Custom renderers](#custom-renderers); they take priority over built-in ones.                                                                                                                                                 |
| `pdfjs`          | PDF.js module or a function returning it | none            | See [PDF rendering with PDF.js](#pdf-rendering-with-pdfjs).                                                                                                                                                                   |
| `tableGridlines` | `boolean`                                | `true`          | Word files: draw dashed gridlines on table cell edges that have no border, as Word does on screen with "View Gridlines" on. They are never printed. `false` shows only the borders stored in the document, as Word prints it. |
| `labels`         | `Partial<Labels>`                        | English         | Override any UI text. See [Translation](#translation).                                                                                                                                                                        |
| `className`      | `string`                                 | none            | Extra class on the root element.                                                                                                                                                                                              |
| `styleNonce`     | `string`                                 | none            | CSP nonce for the injected `<style>` element.                                                                                                                                                                                 |

Invalid options throw an `AnyViewError` with the code `INVALID_OPTIONS` straight away.

## Viewer API

`createViewer` returns a `Viewer`:

| Member                                                | Description                                                                                                                                                                                |
| ----------------------------------------------------- | ------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------ |
| `load(source, { name?, type? })`                      | Opens a file, replacing the current one. Returns a promise that rejects with an `AnyViewError` on failure. If a newer `load()` or `destroy()` supersedes it, the promise resolves quietly. |
| `clear()`                                             | Closes the file and shows the empty state.                                                                                                                                                 |
| `on(event, handler)`                                  | Subscribes to an [event](#events). Returns an unsubscribe function.                                                                                                                        |
| `off(event, handler)`                                 | Unsubscribes.                                                                                                                                                                              |
| `getZoom()` / `setZoom(n)` / `zoomIn()` / `zoomOut()` | Zoom, where `1` is 100% (range 0.1–10). `getZoom()` is `null` for files that cannot be zoomed.                                                                                             |
| `getPage()` / `goToPage(n)`                           | Page, slide or chapter navigation. `getPage()` returns `{ page, pageCount }` or `null`.                                                                                                    |
| `download()`                                          | Saves the current file.                                                                                                                                                                    |
| `print()`                                             | Prints the current file only, not the surrounding page.                                                                                                                                    |
| `toggleFullscreen()`                                  | Enters or leaves full screen.                                                                                                                                                              |
| `setTheme(theme)`                                     | Switches theme.                                                                                                                                                                            |
| `destroy()`                                           | Removes the viewer and frees everything it holds. Call it when the viewer is no longer needed.                                                                                             |
| `state`                                               | `'idle'`, `'loading'`, `'ready'`, `'error'` or `'destroyed'`.                                                                                                                              |
| `format`                                              | The detected format (`{ id, label, category, extensions, mimeTypes }`) or `null`.                                                                                                          |
| `element`                                             | The viewer's root element.                                                                                                                                                                 |

Also exported: `detectFormat`, `sniffFormat`, `getFormatByExtension`, `getFormatByMimeType`, `getFormatById`, `FORMATS`, `SUPPORTED_FORMAT_IDS`, `DEFAULT_LABELS`, `ERROR_CODES` and `VERSION`.

```js
import { getFormatByExtension, SUPPORTED_FORMAT_IDS } from '@buffmomo/anyview';

// Decide whether to show a "Preview" button for an attachment.
const format = getFormatByExtension('pptx');
const canPreview = format !== null && SUPPORTED_FORMAT_IDS.includes(format.id);
```

## Events

| Event        | Payload                                      | When                                        |
| ------------ | -------------------------------------------- | ------------------------------------------- |
| `loadstart`  | `{ source }`                                 | A load begins.                              |
| `load`       | `{ format, file: { name, size, mimeType } }` | The file is displayed.                      |
| `error`      | `AnyViewError`                               | A load failed.                              |
| `pagechange` | `{ page, pageCount }`                        | The current page, slide or chapter changed. |
| `zoomchange` | `{ zoom }`                                   | The zoom level changed.                     |
| `destroy`    | none                                         | The viewer was destroyed.                   |

## Error handling

Every failure is an `AnyViewError` with a stable `code`, a readable `message`, optional `details` and the original `cause`. The viewer also shows the error in place of the file, with a download button where that makes sense, so users are never left with a blank area.

| Code                 | Meaning                                                                                |
| -------------------- | -------------------------------------------------------------------------------------- |
| `INVALID_SOURCE`     | The source is not a URL, `File`, `Blob` or bytes, or the URL is empty or malformed.    |
| `INVALID_OPTIONS`    | A bad option, a missing container, or no browser environment.                          |
| `UNSUPPORTED_FORMAT` | The file type is unknown or has no renderer. `details.format` holds the detected id.   |
| `FETCH_FAILED`       | The download failed: network error, CORS, or an HTTP error (`details.status`).         |
| `FILE_TOO_LARGE`     | The file, or something inside an archive, exceeds a size limit.                        |
| `PARSE_FAILED`       | The file is corrupt or not what its name claims.                                       |
| `RENDER_FAILED`      | The file was read but could not be displayed (for example an unsupported video codec). |
| `ABORTED`            | The operation was cancelled.                                                           |
| `DESTROYED`          | A method was called on a destroyed viewer.                                             |

```js
import { createViewer, isAnyViewError } from '@buffmomo/anyview';

const viewer = createViewer('#preview');

try {
  await viewer.load(url);
} catch (error) {
  if (isAnyViewError(error) && error.code === 'FETCH_FAILED' && error.details.status === 401) {
    redirectToLogin();
  }
}
```

A file passed through the `src` option is loaded right after creation; its errors are delivered through the `error` event (there is no promise to reject).

## TypeScript

Types ship with the package. Events, options and error codes are fully typed:

```ts
import {
  createViewer,
  type AnyViewErrorCode,
  type Viewer,
  type ViewerOptions,
} from '@buffmomo/anyview';

const options: ViewerOptions = {
  theme: 'dark',
  toolbar: { print: false, fullscreen: false },
  maxFileSize: 50 * 1024 * 1024,
};

const viewer: Viewer = createViewer(document.querySelector<HTMLElement>('#preview')!, options);

viewer.on('pagechange', ({ page, pageCount }) => {
  document.title = `Page ${page} of ${pageCount}`;
});

viewer.on('error', (error) => {
  const code: AnyViewErrorCode = error.code; // 'FETCH_FAILED' | 'PARSE_FAILED' | …
  console.warn(code);
});
```

## Advanced usage

### PDF rendering with PDF.js

By default PDFs are shown with the browser's built-in PDF viewer. That needs no extra code and works well on desktop, but some mobile browsers do not display PDFs inline. For identical rendering everywhere, install [PDF.js](https://mozilla.github.io/pdf.js/) in your app and hand it to AnyView. It stays your dependency, not AnyView's, and is loaded only when a PDF is opened:

```bash
npm install pdfjs-dist
```

```js
import { createViewer } from '@buffmomo/anyview';

createViewer('#preview', {
  src: '/files/contract.pdf',
  pdfjs: async () => {
    const pdfjs = await import('pdfjs-dist');
    pdfjs.GlobalWorkerOptions.workerSrc = new URL(
      'pdfjs-dist/build/pdf.worker.mjs',
      import.meta.url,
    ).toString();
    return pdfjs;
  },
});
```

With PDF.js, pages are drawn lazily as they scroll into view, and the toolbar gains page navigation and zoom. The adapter is tested against PDF.js 6 and also accepts the older API in which the document object has its own `destroy()` method.

How to obtain the worker URL depends on your bundler; the `new URL(…, import.meta.url)` form above is the one the PDF.js ecosystem recommends for webpack 5 and Vite. If it does not resolve in your setup, see the PDF.js documentation for your bundler.

### Custom renderers

Add a format AnyView does not know, or replace a built-in renderer:

```ts
import { createViewer, type RendererDefinition } from '@buffmomo/anyview';

const geojson: RendererDefinition = {
  id: 'geojson-map',
  extensions: ['geojson'],
  // Loaded only when a matching file is opened.
  load: async () => {
    const { drawMap } = await import('./draw-map');
    return {
      async render(context) {
        const data = JSON.parse(await context.file.text());
        const map = drawMap(context.container, data);
        return {
          destroy: () => map.remove(),
          zoom: { get: () => map.zoom, set: (value) => map.setZoom(value) },
        };
      },
    };
  },
};

createViewer('#preview', { src: '/data/regions.geojson', renderers: [geojson] });
```

A renderer receives a `context` with the file (`name`, `blob()`, `bytes()`, `text()`), an empty `container` to fill, an `AbortSignal`, and helpers such as `objectUrl(blob)` (revoked automatically). It may return `zoom`, `pages`, `print` and `destroy` so the toolbar can drive it. To override a built-in renderer, list its format ids: `{ id: 'my-csv', formats: ['csv'], load }`. Full guide: [docs/custom-renderers.md](docs/custom-renderers.md).

### Theming

All styles live under the `.av-root` class and use CSS custom properties, which you can override:

```css
.av-root {
  --av-accent: #7c3aed;
  --av-bg: #faf5ff;
  --av-surface: #ffffff;
  --av-fg: #1e1b4b;
  --av-muted: #6b7280;
  --av-border: #e5e7eb;
}
```

Use `[data-theme='dark']` to target the dark theme. Syntax colours are `--av-tok-comment`, `--av-tok-string`, `--av-tok-number`, `--av-tok-keyword`, `--av-tok-function`, `--av-tok-tag` and `--av-tok-attr`.

### Translation

Pass any of the texts in `DEFAULT_LABELS`:

```js
createViewer('#preview', {
  labels: {
    loading: 'Chargement…',
    download: 'Télécharger',
    print: 'Imprimer',
    errorTitle: 'Impossible d’afficher ce fichier',
    pageOf: '{page} sur {count}',
  },
});
```

### Files behind authentication

```js
createViewer('#preview', {
  src: 'https://api.example.com/files/42',
  name: 'contract.docx', // helps detection when the URL has no extension
  fetchOptions: { headers: { Authorization: `Bearer ${token}` } },
});
```

With `fetchOptions`, every format (including images and video) is downloaded through `fetch` so that your headers apply.

## Security

AnyView is built to display files you do not trust.

- **No markup injection.** Rendered output is built with `createElement` and `textContent`. The package never assigns file content to `innerHTML`; a lint rule in this repository enforces it.
- **Sandboxed HTML.** HTML files, EPUB chapters and HTML email bodies are shown in an `<iframe sandbox>` with every restriction on: no scripts, no forms, no popups, no access to your page or its cookies.
- **Inert SVG.** SVG files are displayed through `<img>`, where their scripts cannot run.
- **Safe links.** Only `http`, `https`, `mailto` and `tel` links are kept, and they open in a new tab with `noopener noreferrer`.
- **Resource limits.** A maximum file size, a cap on archive entries, and a cap on decompressed size protect against "zip bombs".

Details and a Content-Security-Policy example: [docs/security.md](docs/security.md).

## Limits and known limitations

AnyView aims to show Office files the way Word, Excel and PowerPoint show them, without replacing those applications.

- **Office fidelity.** Word documents are paginated with the browser's text layout: with the document's fonts available, line and page breaks match Word in the documents this was tested with, but a browser is not Word and unusual documents can differ. Not reproduced: comments, text columns and equations in DOCX; frozen panes, rotated text and drawing shapes in XLSX; animations, SmartArt and shadow or 3D effects in PPTX. Formulas show the results stored in the file. The full list per format is in [docs/formats.md](docs/formats.md).
- **Legacy binary formats** (`.doc`, `.xls`, `.ppt`, `.msg`) are not supported.
- **PDF without PDF.js** depends on the browser's own viewer.
- **Video and audio** depend on the codecs the browser supports.
- **Large files.** Spreadsheets show up to 100 000 rows and 2 000 columns per sheet; text files show the first 5 MB; archives list up to 5 000 entries. A notice appears when something was cut.
- **Password-protected** Office files and encrypted ZIP entries cannot be opened.
- **Fonts.** Documents use fonts embedded in the file and fonts installed on the viewer's device. When a font is missing, the browser substitutes one and line and page breaks can differ from the original.

## Troubleshooting

**The viewer is invisible or only a thin strip.**
The container has no height. Set one (`height: 600px`, or a flex/grid cell with a defined size).

**"The file could not be downloaded (network error or blocked by CORS)".**
The file is on another origin that does not send `Access-Control-Allow-Origin`. Allow your origin on that server, serve the file through your own backend, or fetch it yourself and pass the `Blob`.

**A file opens as the wrong type, or "Preview is not available".**
The URL has no extension and the server sends a generic content type. Pass `name: 'report.docx'` or `type: 'docx'`.

**PDFs do not show on a phone.**
Some mobile browsers have no inline PDF viewer. Use the [PDF.js option](#pdf-rendering-with-pdfjs).

**Styles are missing under a strict Content-Security-Policy.**
AnyView injects one `<style>` element. Pass `styleNonce`, and allow `blob:` in `img-src`, `media-src` and `frame-src`. See [docs/security.md](docs/security.md).

**"DecompressionStream is not available".**
The browser is older than the [requirements](#requirements); Office, OpenDocument, EPUB and ZIP files need it.

**Next.js or Nuxt: "document is not defined".**
`createViewer` was called during server rendering. Use the React or Vue wrapper, or call it inside `useEffect` / `onMounted`.

**Jest (jsdom) cannot run the viewer.**
jsdom lacks several browser APIs AnyView relies on. Mock the component in unit tests and cover it with a browser-based test instead.

## Contributing

Bug reports, sample files that render badly, and pull requests are welcome. See [CONTRIBUTING.md](CONTRIBUTING.md) for setup, the project layout and the dependency policy.

```bash
npm install
npm test            # unit and integration tests
npm run test:browser  # real-browser check (needs Chrome or Edge)
npm run lint && npm run typecheck
```

## Author and contributors

Created and maintained by [pushpa1105](https://github.com/pushpa1105).

Contributors:

- [pushpa1105](https://github.com/pushpa1105)

## License

[MIT](LICENSE)
