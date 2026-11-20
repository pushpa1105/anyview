# Examples

| Folder                                     | Shows                                                                      |
| ------------------------------------------ | -------------------------------------------------------------------------- |
| [vanilla/](vanilla/index.html)             | `createViewer` with a file picker, URL box, drag and drop and theme switch |
| [web-component/](web-component/index.html) | The `<any-view>` element: attributes, the `source` property and events     |
| [react/](react/FilePreview.tsx)            | The React component with callbacks and a ref to the viewer                 |
| [vue/](vue/FilePreview.vue)                | The Vue 3 component with events and a template ref                         |

## Running the HTML examples

The two HTML examples load the built package from `dist/`, so build first and serve the repository root with any static file server (ES modules do not load from `file://`):

```bash
npm install
npm run build
npx serve .
```

Then open `http://localhost:3000/examples/vanilla/` or `http://localhost:3000/examples/web-component/`.

Both accept a `?src=` query parameter to open a file straight away. After `npm test` has generated sample files you can try, for example:

```
http://localhost:3000/examples/vanilla/?src=/tests/.generated/real.docx
http://localhost:3000/examples/vanilla/?src=/tests/.generated/real.xlsx
http://localhost:3000/examples/vanilla/?src=/tests/.generated/real.pptx
```

## Using the React and Vue examples

Copy the component file into an existing React (17 or newer) or Vue (3.2 or newer) project that has the package installed:

```bash
npm install @buffmomo/anyview
```

Both examples open `/files/sample.pdf` when "Open a URL" is clicked; change that path to a file your app serves.
