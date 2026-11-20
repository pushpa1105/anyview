# Contributing to AnyView

Thank you for helping. Bug reports with a sample file are especially valuable: most rendering problems can only be fixed with a file that shows them. Please remove confidential content first.

## Setup

Requirements: Node.js 20 or newer and npm. For the browser check you also need Chrome or Edge installed.

```bash
git clone <your fork>
cd anyview
npm install
npm test
```

## Commands

| Command                 | What it does                                                                                                                                        |
| ----------------------- | --------------------------------------------------------------------------------------------------------------------------------------------------- |
| `npm run build`         | Builds `dist/` (ESM, CommonJS and type declarations) with tsup                                                                                      |
| `npm run dev`           | Rebuilds on change                                                                                                                                  |
| `npm test`              | Runs unit and integration tests (Vitest with a simulated DOM)                                                                                       |
| `npm run test:coverage` | The same with a coverage report; fails below 90% lines                                                                                              |
| `npm run test:browser`  | Builds, then opens every sample file in real Chrome or Edge and saves screenshots to `tests/.generated/screenshots/`                                |
| `npm run site`          | Assembles the documentation site and playground into `dist-site/` (run `npm test` and `npm run build` first); preview it with `npx serve dist-site` |
| `npm run lint`          | ESLint, including the "no markup injection" security rule                                                                                           |
| `npm run typecheck`     | TypeScript in strict mode over sources and tests                                                                                                    |
| `npm run format`        | Prettier                                                                                                                                            |
| `npm run check:package` | Validates the published package with publint and Are The Types Wrong                                                                                |

Before opening a pull request, run:

```bash
npm run format && npm run lint && npm run typecheck && npm run test:coverage && npm run test:browser
```

## Project layout

```
src/
  index.ts        public entry point
  element.ts      <any-view> custom element
  react.ts        React wrapper          vue.ts   Vue wrapper
  core/           viewer, format table and detection, source loading, registry, errors, labels
  renderers/      one small module per format; each is a lazily loaded chunk
  formats/        parsers that turn a file format into DOM or a data model (docx, xlsx, pptx, odf, rtf, epub, eml)
  ui/             shared views: grid, sheet tabs, slides, paper, toolbar, print, styles, DOM helpers
  util/           zip, xml, csv, markdown, highlight, number formats, text decoding, URL checks
tests/
  unit/           tests per module, using fixtures built in memory (tests/helpers/fixtures.ts)
  integration/    every format through the public API, plus real files written by third-party libraries
  browser/        pages used by the real-browser check (scripts/browser-check.mjs)
examples/         runnable usage examples
docs/             format details, security model, custom-renderer guide
```

How a file gets on screen: `core/viewer.ts` wraps the source (`core/source.ts`), detects the format (`core/formats.ts`), picks a renderer (`core/registry.ts`), imports it on demand and calls `render(context)`.

## Rules for code

### 1. No runtime dependencies

`package.json` has no `dependencies`, and that is a feature users rely on. Parsers are written here. A runtime dependency will only be considered when a format is truly out of reach by hand, and even then the preferred route is the one used for PDF.js: the application supplies the library and AnyView adapts to it. Development dependencies are fine.

### 2. Never inject markup

File content is untrusted. Build output with `createElement`, `textContent` and `style.setProperty` (the `h()` helper in `src/ui/dom.ts` does this). `innerHTML`, `outerHTML`, `insertAdjacentHTML` and `document.write` are blocked by ESLint. Content that must be rendered as HTML goes into `createSandboxedFrame()`.

Validate anything from a file before it reaches a style or attribute: colours must match a hex pattern, URLs go through `safeLinkUrl` / `safeImageUrl`, lengths must be numbers.

### 3. Fail clearly

Throw `AnyViewError` with the right code and a message a user can understand. Never render an empty box for a file that could not be read.

### 4. Stay lazy and bounded

Renderers are loaded with dynamic `import()`; do not import a renderer or a format parser from `core/`. Cap anything whose size comes from the file (rows, entries, repeat counts) and say so in a notice when something is cut.

### 5. Clean up

Whatever a renderer creates (object URLs via `context.objectUrl`, observers, listeners, timers) must be released in `destroy()`. Check `context.signal` after long awaits.

## Adding a format

1. Add it to the table in `src/core/formats.ts` (id, label, category, extensions, MIME types) and, if it has a recognisable signature, to `sniffFormat`.
2. Write the renderer in `src/renderers/<name>.ts` exporting `renderer`. Put substantial parsing in `src/formats/` or `src/util/`.
3. Register it in `src/core/registry.ts` with a dynamic import.
4. Add a fixture to `tests/helpers/fixtures.ts` and tests: unit tests for the parser, a case in `tests/integration/all-formats.test.ts`, and a check in `scripts/browser-check.mjs`.
5. Document it in `README.md` and `docs/formats.md`, and add a line to `CHANGELOG.md`.

## Tests

- Unit tests run in happy-dom, which has no layout engine: sizes are zero and images never load. Tests therefore assert structure and styles, and dispatch `load` events by hand where a renderer waits for one. Layout and decoding are covered by the browser check.
- Fixtures are built in memory, so the repository contains no binary test files. The integration suite additionally generates real DOCX, XLSX and PPTX files with independent libraries (`docx`, `exceljs`, `pptxgenjs`) to make sure the parsers do not only understand their own fixtures.
- A bug fix should come with a test that fails without it.

## Commits and releases

- Keep pull requests focused; describe what changed and why.
- Update `CHANGELOG.md` under "Unreleased".
- The project follows semantic versioning. Error codes, format ids, option names and CSS custom properties are public API.

## Code of conduct

Be respectful and constructive. Assume good intent, and keep discussion about the code.
