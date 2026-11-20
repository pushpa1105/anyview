# Markdown in AnyView

AnyView renders Markdown with its own parser. Nothing in this file is inserted
into the page as raw HTML.

## What is supported

- **Bold**, _italic_, ~~strikethrough~~ and `inline code`
- [Links](https://github.com/pushpa1105/anyview) and images
- Lists, including task lists:
  - [x] Headings
  - [x] Tables
  - [ ] Your own content

> Block quotes work too.

| Format     | Extension | Laid out in pages |
| ---------- | --------- | ----------------- |
| Word       | `.docx`   | Yes               |
| Excel      | `.xlsx`   | No, a grid        |
| PowerPoint | `.pptx`   | One slide each    |

```js
import { createViewer } from '@buffmomo/anyview';

createViewer('#preview', { src: '/files/notes.md' });
```
