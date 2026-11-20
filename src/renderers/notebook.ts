import { AnyViewError } from '../core/errors';
import type { Renderer } from '../core/types';
import { h } from '../ui/dom';
import { highlightToFragment } from '../util/highlight';
import { renderMarkdown } from '../util/markdown';
import { safeImageUrl } from '../util/url';

type Json = null | boolean | number | string | Json[] | { [key: string]: Json };
type JsonObject = { [key: string]: Json };

function isObject(value: Json | undefined): value is JsonObject {
  return typeof value === 'object' && value !== null && !Array.isArray(value);
}

/** Notebook text fields are either a string or an array of lines. */
function joinText(value: Json | undefined): string {
  if (typeof value === 'string') return value;
  if (Array.isArray(value)) return value.filter((line) => typeof line === 'string').join('');
  return '';
}

// eslint-disable-next-line no-control-regex
const ANSI_ESCAPE = /\u001b\[[0-9;]*[A-Za-z]/g;

function renderOutput(output: JsonObject, parent: HTMLElement): void {
  const type = output['output_type'];
  if (type === 'stream') {
    parent.appendChild(
      h('pre', {
        class: output['name'] === 'stderr' ? 'av-nb-output av-nb-error' : 'av-nb-output',
        text: joinText(output['text']).replace(ANSI_ESCAPE, ''),
      }),
    );
    return;
  }
  if (type === 'error') {
    const lines = Array.isArray(output['traceback']) ? output['traceback'] : [];
    const traceback = lines.filter((line) => typeof line === 'string').join('\n');
    const summary = `${joinText(output['ename'])}: ${joinText(output['evalue'])}`;
    parent.appendChild(
      h('pre', {
        class: 'av-nb-output av-nb-error',
        text: (traceback || summary).replace(ANSI_ESCAPE, ''),
      }),
    );
    return;
  }
  const data = output['data'];
  if (!isObject(data)) return;
  for (const mimeType of ['image/png', 'image/jpeg', 'image/gif']) {
    const encoded = joinText(data[mimeType]).replace(/\s+/g, '');
    const source = encoded ? safeImageUrl(`data:${mimeType};base64,${encoded}`) : null;
    if (source) {
      parent.appendChild(h('img', { attrs: { src: source, alt: '' } }));
      return;
    }
  }
  if (typeof data['text/markdown'] !== 'undefined') {
    const block = h('div');
    block.appendChild(renderMarkdown(joinText(data['text/markdown'])));
    parent.appendChild(block);
    return;
  }
  // HTML output is not interpreted; the plain-text representation is shown instead.
  const text = joinText(data['text/plain']);
  if (text) parent.appendChild(h('pre', { class: 'av-nb-output', text }));
}

export const renderer: Renderer = {
  async render(context) {
    let notebook: Json;
    try {
      notebook = JSON.parse(await context.file.text()) as Json;
    } catch (cause) {
      throw new AnyViewError('PARSE_FAILED', 'This notebook is not valid JSON.', { cause });
    }
    if (!isObject(notebook)) {
      throw new AnyViewError('PARSE_FAILED', 'This file is not a Jupyter notebook.');
    }
    // nbformat 4 keeps cells at the top level; nbformat 3 nests them in worksheets.
    let cells = notebook['cells'];
    if (!Array.isArray(cells) && Array.isArray(notebook['worksheets'])) {
      cells = notebook['worksheets'].flatMap((sheet) =>
        isObject(sheet) && Array.isArray(sheet['cells']) ? sheet['cells'] : [],
      );
    }
    if (!Array.isArray(cells)) {
      throw new AnyViewError(
        'PARSE_FAILED',
        'This file is not a Jupyter notebook (it has no cells).',
      );
    }
    const metadata = isObject(notebook['metadata']) ? notebook['metadata'] : {};
    const languageInfo = isObject(metadata['language_info']) ? metadata['language_info'] : {};
    const kernel = isObject(metadata['kernelspec']) ? metadata['kernelspec'] : {};
    const language = joinText(languageInfo['name']) || joinText(kernel['language']) || 'python';

    const prose = h('div', { class: 'av-prose' });
    for (const cell of cells) {
      if (!isObject(cell)) continue;
      const source = joinText(cell['source'] ?? cell['input']);
      const block = h('div', { class: 'av-nb-cell' });
      if (cell['cell_type'] === 'markdown') {
        block.appendChild(renderMarkdown(source));
      } else if (cell['cell_type'] === 'code') {
        const code = h('pre', { class: 'av-nb-code' });
        code.appendChild(highlightToFragment(source, language));
        block.appendChild(code);
        if (Array.isArray(cell['outputs'])) {
          for (const output of cell['outputs']) if (isObject(output)) renderOutput(output, block);
        }
      } else if (source) {
        block.appendChild(h('pre', { class: 'av-nb-output', text: source }));
      }
      prose.appendChild(block);
    }
    context.container.appendChild(h('div', { class: 'av-scroll' }, prose));

    let zoom = 1;
    return {
      zoom: {
        get: () => zoom,
        set(value: number): void {
          zoom = value;
          prose.style.setProperty('zoom', String(value));
        },
      },
      printable: () => prose,
    };
  },
};
