# Security model

A file viewer handles content written by people you do not know. This page explains what AnyView does with that content and what remains your responsibility.

## Threats and how they are handled

### Script injection (XSS)

The classic attack on a viewer is a file whose text ends up parsed as HTML in the host page: a Markdown file containing `<img onerror=…>`, a spreadsheet cell containing `<script>`, an SVG with an embedded script.

- **No markup strings.** Every renderer builds its output with `document.createElement`, `textContent` and CSSOM style properties. File content is never passed to `innerHTML`, `outerHTML`, `insertAdjacentHTML` or `document.write`. An ESLint rule in this repository rejects those APIs, so a regression fails CI.
- **HTML is sandboxed.** HTML files, EPUB chapters and HTML email bodies must be rendered as HTML, so they go into an `<iframe>` with an empty `sandbox` attribute. That enables every restriction: scripts do not run, forms cannot be submitted, popups and top-level navigation are blocked, and the frame has an opaque origin with no access to your page, cookies or storage.
- **SVG is displayed through `<img>`.** Browsers do not execute scripts or load external resources for SVG used as an image.
- **EPUB chapters are cleaned as well as sandboxed.** Scripts, frames, forms, event-handler attributes and links are removed before the chapter reaches the frame.
- **Notebook HTML outputs are not rendered.** The plain-text representation is shown instead.

The repository's tests open hostile files in every text-based format and assert that no script element or event handler appears in the host document; the real-browser check confirms that scripts embedded in HTML and SVG samples never run.

### Dangerous links

Links found in documents are kept only if their scheme is `http`, `https`, `mailto` or `tel`. Control characters and whitespace that browsers ignore inside URLs are stripped before the check, so `java\nscript:` does not slip through. Links open in a new tab with `rel="noopener noreferrer nofollow"`. Image sources are limited to `http(s)`, `blob:` and base64 raster `data:` URLs.

### Style injection

Colours, sizes and fonts read from files are validated (hex colours, numeric lengths, font names limited to letters, digits, spaces and a few punctuation marks, always written inside quotes) and applied through `style.setProperty`, one property at a time. A value cannot introduce additional declarations. Shapes, charts and icons are SVG elements built with `createElementNS`; outlines are generated from numbers, never copied from the file as markup. Fonts embedded in Word documents are registered through the `FontFace` API under names private to that document and removed when the file is closed; the browser's font sanitiser checks them like any web font.

### Resource exhaustion

- `maxFileSize` (200 MB by default) rejects oversized files before parsing.
- ZIP archives are limited to 100 000 entries and 256 MB per decompressed entry; decompression stops as soon as the output exceeds the limit, which defuses "zip bombs".
- Spreadsheets materialise at most 100 000 rows and 2 000 columns per sheet, and padding declared with huge repeat counts is not expanded.
- Text display is capped at 5 MB; highlighting at 400 000 characters.
- Nested MIME parts are limited to 12 levels.

These limits bound memory use, but parsing a very large file still takes time on the main thread. If your users upload arbitrary files, pick a `maxFileSize` that suits your audience's devices.

### Privacy

AnyView makes no network requests of its own beyond fetching the URL you give it. There is no telemetry. Two things can cause requests to other servers, both initiated by the file's content:

- An HTML file, EPUB chapter or HTML email can reference remote images or stylesheets, which the sandboxed frame will load. This can reveal to the sender that the file was opened (a "tracking pixel"). If that matters to you, apply a Content-Security-Policy (below) or convert such files to text before display.
- Markdown images with `http(s)` URLs are loaded when the document is shown.

## Content-Security-Policy

AnyView works under a strict policy. It needs:

| Directive     | Value                             | Why                                                         |
| ------------- | --------------------------------- | ----------------------------------------------------------- |
| `style-src`   | a nonce (pass it as `styleNonce`) | One injected `<style>` element holds the viewer's CSS       |
| `img-src`     | `blob:` `data:`                   | Images extracted from documents; notebook image outputs     |
| `media-src`   | `blob:`                           | Video and audio opened from a `File` or with `fetchOptions` |
| `frame-src`   | `blob:`                           | The browser PDF viewer (not needed when you use PDF.js)     |
| `connect-src` | the origins your files come from  | Downloads through `fetch`                                   |

Example:

```
Content-Security-Policy:
  default-src 'self';
  style-src 'self' 'nonce-R4nd0m';
  img-src 'self' blob: data:;
  media-src 'self' blob:;
  frame-src 'self' blob:;
```

```js
createViewer('#preview', { styleNonce: 'R4nd0m' });
```

Notes:

- Sandboxed `srcdoc` frames inherit your page's policy. If your `style-src` does not include `'unsafe-inline'`, HTML files, EPUB chapters and HTML emails are shown without their own inline styles. That is usually an acceptable trade for a strict policy.
- Your `img-src` also applies inside those frames, so a policy without remote hosts blocks tracking pixels in HTML email.
- No `script-src` allowance is needed. AnyView does not use `eval`, inline scripts or workers. (PDF.js, if you add it, uses a worker from a URL you control.)

## What stays your responsibility

- **The built-in PDF viewer is the browser's.** AnyView hands it a blob that is explicitly typed `application/pdf`. Its behaviour (including JavaScript inside PDFs, which browsers confine to the PDF sandbox) is governed by the browser.
- **Downloads.** The download button saves the original file unchanged. Opening it in a desktop application carries the usual risks of that application.
- **Where files come from.** AnyView does not authenticate or authorise anything. Use `fetchOptions` or fetch the file yourself and pass a `Blob`.
- **Custom renderers.** Code you register runs with your page's privileges. Follow the same rule: build DOM nodes, do not inject markup.

## Reporting a vulnerability

Please report security issues privately to the maintainers rather than opening a public issue. Include a sample file that demonstrates the problem.
