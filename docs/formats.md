# Supported formats in detail

This page lists what each built-in renderer shows and what it leaves out. "Format id" is the value of `viewer.format.id` and what you pass as the `type` option.

## How a format is chosen

1. The `type` option, if given: a format id (`docx`), an extension (`.docx`) or a MIME type.
2. The file name's extension.
3. The MIME type (`Blob.type` or the `Content-Type` response header). Generic types such as `application/octet-stream`, `application/zip` and `text/plain` are treated as inconclusive.
4. The content: magic numbers for binary formats, the package layout for ZIP-based formats (DOCX, XLSX, PPTX, ODF, EPUB), and the leading text for HTML, SVG, XML, JSON, notebooks and email.

You can call the same logic yourself:

```js
import { detectFormat, sniffFormat } from '@buffmomo/anyview';

const format = await detectFormat({
  name: file.name,
  mimeType: file.type,
  bytes: async () => new Uint8Array(await file.arrayBuffer()),
});
const fromBytes = sniffFormat(firstBytes); // FormatInfo | null
```

## Documents

### Word (`docx`): .docx, .docm, .dotx, .dotm

Word documents are laid out the way Word lays them out: as pages. The aim is that a document looks the same as it does in Word, with the same line breaks and the same number of pages when the fonts are available.

Shown:

- **Pages:** page size, orientation and margins per section; automatic page breaks between lines (with widow and orphan control), explicit page breaks, "page break before", "keep with next" and "keep lines together"; headers and footers (default, first page, odd and even) with live `PAGE` and `NUMPAGES` fields; footnotes at the bottom of the page that refers to them and endnotes at the end; table rows split across pages with repeated header rows. The toolbar navigates by page.
- **Text:** fonts (including theme fonts and fonts embedded in the document), size, colour (including theme colours with tints and shades), highlight and shading, bold, italic, underline styles and colours, single and double strike, super/subscript, small and all capitals, character spacing, raised and lowered text, text borders.
- **Paragraphs:** alignment, indents, hanging indents, spacing before and after, line spacing (single, multiple, exact, at least) computed from the font's own metrics, borders and shading, drop caps and simple text frames, right-to-left paragraphs, contextual spacing.
- **Tab stops:** left, centre, right and decimal stops, default stops, absolute position tabs and leaders (dots, dashes, lines), so tables of contents line up.
- **Styles:** paragraph, character and table styles with inheritance, document defaults, and Word's rule that table styles win over the default paragraph style inside tables.
- **Lists:** multi-level numbering, bullets (including picture bullets and Symbol/Wingdings bullets), right-aligned numbers, lists that continue each other, restarts.
- **Tables:** table styles with conditional formatting (header row, total row, first and last column, banded rows and columns, corner cells), per-side borders with conflict resolution between neighbouring cells, shading, cell margins, merged cells, fixed column widths, row heights, nested tables, percentage widths, indents, floating tables with text wrapped around them, vertical text.
- **Pictures and shapes:** inline pictures; floating pictures wrapped by text (square, tight), placed above or below text, or positioned on the page in front of or behind text; shapes with preset or custom outlines, solid, gradient and picture fills, lines with dashes and arrow heads; text boxes; groups; legacy VML pictures, text boxes and horizontal rules.
- **Fields and links:** hyperlinks, links to bookmarks (clicking scrolls to the target), tables of contents, `HYPERLINK` fields, cached results of other fields.

Not shown: comments, tracked deletions (insertions are shown), multiple text columns, page borders and watermarks stored as page backgrounds, charts and SmartArt inside documents, EMF/WMF images (a placeholder keeps their space), equations (shown as linear text), text wrapped tightly around irregular shapes (treated as square), and line numbering.

Fidelity notes:

- Pagination uses the browser's text layout. With the document's fonts available (installed, or embedded in the file) the page breaks match Word in the documents this was tested with. When a font is missing the browser substitutes one, lines break differently and page breaks can move. Office "cloud fonts" such as Aptos are only available to a browser if they are installed on the device.
- Word does not load embedded fonts while a document is open in Protected View or read-only mode, so such a document can look different in Word itself until editing is enabled. AnyView always uses the embedded fonts.
- Layout runs once after loading. Without a layout engine (for example in a server-side test environment) the document is split at explicit page breaks only.

### OpenDocument Text (`odt`): .odt, .ott

Shown: headings, paragraph and text styles with inheritance, nested lists, tables with spans, images, hyperlinks, footnote markers, page size and margins.

Not shown: footnote bodies, headers and footers, tracked changes, drawing shapes.

### Rich Text (`rtf`): .rtf

Shown: bold, italic, underline, strike, size, text and highlight colours, super/subscript, alignment, indents, spacing, simple tables, Unicode and code-page text, embedded PNG and JPEG pictures.

Not shown: fonts, styles, lists as structured lists (their text is shown), WMF/EMF pictures, headers and footers.

### PDF (`pdf`): .pdf

Without the `pdfjs` option, the PDF is displayed by the browser's own viewer in a frame, which brings its own toolbar, search and printing. With the `pdfjs` option, pages are drawn to canvases, lazily, with AnyView's page navigation and zoom. Text selection and search are not provided in that mode.

## Spreadsheets

All spreadsheet formats use the same grid: only the visible cells exist in the DOM, so large sheets stay fast. Zoom scales the grid; printing produces a table of the first 2 000 rows.

### Excel (`xlsx`): .xlsx, .xlsm, .xltx, .xltm

Workbooks are drawn with their own formatting on a white sheet, as Excel shows them, in both the light and the dark viewer theme.

Shown:

- **Cells:** all visible worksheets as tabs; text, rich text, numbers, booleans, errors and cached formula results; fonts, sizes, bold, italic, underline, strike and colours; solid, pattern and gradient fills; borders per side with their styles and colours; horizontal and vertical alignment, wrapping and indents; theme colours with tints, indexed colours and custom palettes.
- **Number formats:** thousands separators, decimals, percentages, scientific notation, currency and accounting formats (symbol on the left, number on the right), colours such as `[Red]`, negative-number sections, dates, times and elapsed time; the 1900 and 1904 date systems.
- **Layout:** column widths, row heights, hidden rows and columns, merged cells, text that overflows into empty neighbouring cells, and sheets with gridlines switched off.
- **Tables:** table styles (built-in and custom) with header and total rows, banding, first and last column, and filter buttons.
- **Conditional formatting:** icon sets, data bars, colour scales, and rules on values, text, blanks, errors, duplicates, top/bottom and above/below average; formula rules where the formula can be evaluated.
- **Pictures and charts:** pictures positioned over the cells; bar, column, line, area, pie, doughnut and scatter charts (including stacked and combined charts) drawn from the values saved in the file.

Not shown: frozen panes, rotated text, shapes and text boxes in drawings, hyperlinks, data validation, comments, sparklines, pivot table interactivity, hidden sheets, fraction formats (shown as decimals). Radar, stock and bubble charts are approximated; surface charts and secondary axes are not drawn. Built-in table styles are close approximations of Excel's.

Formulas are not recalculated: the value saved in the file is displayed.

### OpenDocument Spreadsheet (`ods`): .ods, .ots

Shown: every table as a tab, cell text as formatted by the producing application, merged cells. Cell styles are not applied.

### CSV and TSV (`csv`, `tsv`): .csv, .tsv, .tab

RFC 4180 parsing: quoted fields, doubled quotes, line breaks inside fields. The delimiter (comma, semicolon, tab or pipe) is detected from the first lines; `.tsv` always uses tabs. Encoding is UTF-8 or UTF-16 (by byte-order mark), with Windows-1252 as a fallback.

## Presentations

Slides are shown as a vertical list scaled to the available width, built as they scroll into view. The toolbar navigates by slide.

### PowerPoint (`pptx`): .pptx, .pptm, .ppsx, .ppsm, .potx, .potm

Shown: slide size; text with size, bold, italic, underline, strike, colour, font, super/subscript and capitals; paragraph alignment, indents, line spacing, bullets and automatic numbering; placeholder position and text style inherited from the layout and master; auto-fit font scaling; more than fifty preset shapes (rectangles, ellipses, triangles, arrows, stars, polygons, flowchart shapes, brackets, connectors and more) and custom shapes drawn from their outlines, with solid, gradient and picture fills, outlines, dashes and arrow heads; rotation and flips; pictures, including cropped pictures; groups; tables with cell borders, fills, margins and vertical alignment; bar, column, line, area, pie, doughnut and scatter charts drawn from the values saved in the file; solid, gradient and picture backgrounds; theme colours with tint, shade, luminance and transparency adjustments; decorative shapes from the layout and master.

Not shown: animations and transitions, SmartArt (a labelled placeholder marks its position), table styles referenced by id (explicit cell formatting is used), shadows, reflections and 3D effects, text fitted to the inside of non-rectangular shapes, speaker notes, embedded video.

### OpenDocument Presentation (`odp`): .odp, .otp

Shown: page size, text boxes with basic text styles, pictures, rectangles, ellipses and lines, page background colour. Master-page artwork is not shown.

## Images (`image`, `svg`)

PNG, JPEG, GIF, WebP, AVIF, BMP, ICO and APNG are decoded by the browser. Large images are scaled down to fit at first; zoom is relative to the natural size. SVG is displayed through `<img>`, so scripts and external references inside it are inactive.

## Video and audio (`video`, `audio`)

The browser's native player is used. URLs are streamed without downloading the whole file first (unless `fetchOptions` is set). Whether a file plays depends on the codecs the browser supports; an unsupported file produces a `RENDER_FAILED` error with a clear message.

## Text and code (`text`, `json`, `xml`)

Shown with line numbers. Syntax highlighting covers JavaScript, TypeScript, JSON, HTML/XML, CSS/SCSS/Less, Python, Java, C, C++, C#, Go, Rust, PHP, Ruby, Swift, Kotlin, Scala, Dart, Lua, R, shell, PowerShell, batch, SQL, YAML, TOML, INI, GraphQL, Dockerfile, Makefile, TeX and diffs. The highlighter is a lexer (comments, strings, numbers, keywords, function calls), not a full parser.

Single-line (minified) JSON up to 2 MB is re-indented; numbers are kept exactly as written. Files larger than 400 000 characters are shown without highlighting, and only the first 5 MB of text is displayed.

## Markdown (`markdown`): .md, .markdown, .mdx

CommonMark's everyday features plus GFM tables, task lists, strikethrough and autolinks: headings, emphasis, inline code, fenced and indented code blocks (highlighted when a language is given), block quotes, nested lists, links (inline and reference style), images and rules. Raw HTML in the source is shown as text, not interpreted. Files longer than 1.5 million characters are shown as plain text. Relative links and images resolve against the file's URL when it has one.

## HTML (`html`): .html, .htm, .xhtml

Rendered in a fully sandboxed frame: the page is displayed, but its scripts do not run and it cannot reach your application. Relative resources resolve against the file's URL when it has one.

## EPUB (`epub`)

Chapters follow the book's reading order; the toolbar moves between them. Each chapter is shown in a sandboxed frame with its images and stylesheets inlined. Scripts, forms and links between chapters are removed. The table of contents, fonts embedded in the book and DRM-protected books are not supported.

## Jupyter notebooks (`ipynb`)

Markdown cells are rendered, code cells are highlighted, and outputs are shown: text streams, results, PNG/JPEG/GIF images, Markdown and error tracebacks. HTML and JavaScript outputs are not executed; their plain-text form is shown instead. Notebook formats 3 and 4 are supported.

## Email (`eml`)

Subject, sender, recipients and date; the HTML body in a sandboxed frame, or the plain-text body; attachments listed with their size and opened in a nested preview. Encoded headers, nested multipart messages, base64 and quoted-printable encodings are handled. Images referenced by `cid:` are listed as attachments rather than placed in the body.

## Archives (`zip`): .zip, .jar

Files are listed alphabetically with their size. Clicking one extracts it in memory and opens it in a nested viewer with a **Back** button. Encrypted entries and compression methods other than "stored" and "deflate" cannot be extracted. ZIP64 archives are supported.

## Fonts (`font`): .ttf, .otf, .woff, .woff2

A specimen sheet: alphabets, digits, punctuation and a sample sentence at several sizes.

## Recognised but not displayed

| Format id                  | Files                       | Why                                    |
| -------------------------- | --------------------------- | -------------------------------------- |
| `doc`, `xls`, `ppt`, `ole` | Office 97–2003 binary files | Would need large binary-format parsers |
| `msg`                      | Outlook messages            | Binary compound format                 |
| `heic`                     | HEIC/HEIF images            | No browser decoder outside Safari      |
| `tiff`                     | TIFF images                 | No browser decoder outside Safari      |

These produce an `UNSUPPORTED_FORMAT` error and a download button. You can add support for any of them with a [custom renderer](custom-renderers.md).
