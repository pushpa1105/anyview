# Changelog

All notable changes to this project are documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

## Unreleased

### Added

- **DOCX:** documents are laid out in pages like Word, with headers and footers (including page
  numbers), footnotes on their page, endnotes, sections with their own page size, and page
  navigation in the toolbar. Also new: theme fonts and colours, embedded fonts, tab stops with
  leaders, table styles with conditional formatting and per-side borders, floating tables and
  pictures, drop caps, shapes with preset and custom outlines, groups, picture bullets, bookmarks
  and internal links, and horizontal rules.
- **XLSX:** fonts, fills, borders, alignment, theme and indexed colours, row heights, hidden rows
  and columns, text overflow, accounting formats, table styles, conditional formatting (icon
  sets, data bars, colour scales and rules), pictures and charts.
- **PPTX:** more than fifty preset shapes and custom shapes drawn from their outlines, dashed
  lines and arrow heads, charts, cropped pictures, and table cell borders, margins and alignment.
- `scripts/shot.mjs`, a developer tool that screenshots one file in a real browser for comparison
  with the original application.

### Changed

- Spreadsheets are always drawn on a white sheet with their own colours, also in the dark theme.
- Hidden spreadsheet columns are no longer drawn (they were 8 px wide).
- The DOCX renderer now reports pages (`viewer.getPage()`, `viewer.goToPage()`).

### Fixed

- Border colours of Word tables and paragraphs given as plain hex values were ignored.
- Word tables: numbers in cells are aligned on a decimal tab stop, percentage widths are measured
  the way Word measures them (also for tables inside cells), and a banded row or column no longer
  removes the outer border of the table.

## 0.1.0 - 2026-11-20

First release.

### Added

- `createViewer(container, options)`: a framework-agnostic viewer with a toolbar (zoom, page
  navigation, download, print, full screen), light, dark and automatic themes, translatable
  labels, typed events and typed errors.
- `<any-view>` custom element (`@buffmomo/anyview/element`), a React component
  (`@buffmomo/anyview/react`) and a Vue 3 component (`@buffmomo/anyview/vue`).
- Sources: URLs, `File`, `Blob`, `ArrayBuffer` and typed arrays. Format detection by explicit
  type, extension, MIME type and file content.
- Built-in renderers, each loaded on demand:
  - Documents: DOCX, ODT, RTF.
  - Spreadsheets: XLSX, ODS, CSV, TSV, on a virtualized grid with sheet tabs.
  - Presentations: PPTX, ODP.
  - PDF through the browser's viewer, or through PDF.js when the application supplies it.
  - Images (including SVG), video and audio.
  - Text and source code with syntax highlighting; JSON pretty-printing.
  - Markdown, HTML (sandboxed), EPUB, Jupyter notebooks, email (EML), ZIP archives and fonts.
- Nested previews for files inside archives and email attachments.
- Custom renderers for new formats or to replace built-in ones.
- A download fallback for recognised but unsupported formats (legacy Office, MSG, HEIC, TIFF).
- Zero runtime dependencies: ZIP reading (with ZIP64), XML parsing, spreadsheet number
  formatting, Markdown parsing and syntax highlighting are implemented in the package.
- Security measures for untrusted files: no markup injection, sandboxed frames for HTML
  content, URL and style validation, and limits on file, archive and sheet sizes.
- ESM and CommonJS builds with TypeScript declarations.
