import { formatLabel } from '../core/labels';
import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { highlightToFragment, isHighlightable } from '../util/highlight';
import { formatBytes } from '../util/text';

/** Characters shown at most; beyond this the browser itself becomes sluggish. */
const MAX_DISPLAY_CHARS = 5_000_000;
/** Highlighting creates one element per token, so it is skipped for big files. */
const MAX_HIGHLIGHT_CHARS = 400_000;
const MAX_PRETTY_JSON_CHARS = 2_000_000;

/**
 * Re-indents single-line (minified) JSON. Works on the text itself instead of
 * `JSON.parse` + `JSON.stringify`, so large integers and number formatting are
 * shown exactly as they appear in the file.
 */
export function prettyJson(source: string, indent = '  '): string {
  const text = source.trim();
  if (text.includes('\n') || text.length > MAX_PRETTY_JSON_CHARS) return source;
  try {
    JSON.parse(text);
  } catch {
    return source;
  }
  let output = '';
  let depth = 0;
  let i = 0;
  const newline = (): string => '\n' + indent.repeat(depth);
  while (i < text.length) {
    const char = text[i]!;
    if (char === '"') {
      let end = i + 1;
      while (end < text.length && text[end] !== '"') end += text[end] === '\\' ? 2 : 1;
      output += text.slice(i, end + 1);
      i = end + 1;
    } else if (char === '{' || char === '[') {
      const close = char === '{' ? '}' : ']';
      const rest = /^\s*/.exec(text.slice(i + 1, i + 64))![0].length;
      if (text[i + 1 + rest] === close) {
        output += char + close;
        i += rest + 2;
      } else {
        depth++;
        output += char + newline();
        i++;
      }
    } else if (char === '}' || char === ']') {
      depth--;
      output += newline() + char;
      i++;
    } else if (char === ',') {
      output += ',' + newline();
      i++;
    } else if (char === ':') {
      output += ': ';
      i++;
    } else if (char === ' ' || char === '\t' || char === '\r') {
      i++;
    } else {
      output += char;
      i++;
    }
  }
  return output;
}

export const renderer: Renderer = {
  async render(context) {
    const { file, format, language, container, labels } = context;
    let text = (await file.text()).replace(/\r\n?/g, '\n');
    if (format.id === 'json') text = prettyJson(text);

    if (text.length > MAX_DISPLAY_CHARS) {
      const total = text.length;
      text = text.slice(0, MAX_DISPLAY_CHARS);
      container.appendChild(
        h('div', {
          class: 'av-notice',
          text: formatLabel(labels.truncated, {
            shown: formatBytes(MAX_DISPLAY_CHARS),
            total: formatBytes(total),
          }),
        }),
      );
    }

    const lineCount = text.split('\n').length - (text.endsWith('\n') ? 1 : 0) || 1;
    let numbers = '';
    for (let line = 1; line <= lineCount; line++) numbers += line === 1 ? '1' : `\n${line}`;
    const gutter = h('pre', {
      class: 'av-code-gutter',
      text: numbers,
      attrs: { 'aria-hidden': 'true' },
    });
    const lines = h('pre', { class: 'av-code-lines' });
    if (isHighlightable(language) && text.length <= MAX_HIGHLIGHT_CHARS) {
      lines.appendChild(highlightToFragment(text, language));
    } else {
      lines.textContent = text;
    }
    const table = h('div', { class: 'av-code-table' }, gutter, lines);
    container.appendChild(h('div', { class: 'av-code' }, table));

    let zoom = 1;
    return {
      zoom: {
        get: () => zoom,
        set(value: number): void {
          zoom = value;
          table.style.fontSize = `${13 * zoom}px`;
        },
      },
      printable: () => lines,
    };
  },
};
