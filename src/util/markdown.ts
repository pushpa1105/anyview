import { append, externalLink, h } from '../ui/dom';
import { highlightToFragment } from './highlight';
import { safeImageUrl, safeLinkUrl } from './url';
import { decodeEntities } from './xml';

/**
 * A Markdown renderer covering CommonMark's everyday features plus GFM tables,
 * task lists, strikethrough and autolinks. It builds DOM nodes directly, so raw
 * HTML in the source is shown as text instead of being interpreted.
 */

export interface MarkdownOptions {
  /** Base URL for relative links and images. */
  baseUrl?: string | null;
}

interface LinkReference {
  url: string;
  title: string | undefined;
}

interface State {
  baseUrl: string | null;
  references: Map<string, LinkReference>;
}

const HEADING = /^ {0,3}(#{1,6})(?:[ \t]+(.*?))?(?:[ \t]+#+)?[ \t]*$/;
const FENCE = /^( {0,3})(`{3,}|~{3,})\s*([^\s`]*)[^`]*$/;
const RULE = /^ {0,3}([-*_])(?:[ \t]*\1){2,}[ \t]*$/;
const LIST_ITEM = /^( {0,3})([-*+]|\d{1,9}[.)])( +|$)(.*)$/;
const BLOCKQUOTE = /^ {0,3}>[ ]?(.*)$/;
const SETEXT = /^ {0,3}(=+|-+)[ \t]*$/;
const TABLE_DIVIDER = /^ {0,3}\|?[ \t]*:?-+:?[ \t]*(\|[ \t]*:?-+:?[ \t]*)*\|?[ \t]*$/;
const REFERENCE = /^ {0,3}\[([^\]]+)\]:\s*<?([^\s>]+)>?(?:\s+["'(]([^"')]*)["')])?\s*$/;
const BLANK = /^\s*$/;

function normaliseReference(label: string): string {
  return label.trim().replace(/\s+/g, ' ').toLowerCase();
}

function splitTableRow(line: string): string[] {
  const trimmed = line.trim().replace(/^\|/, '').replace(/\|$/, '');
  const cells: string[] = [];
  let current = '';
  for (let i = 0; i < trimmed.length; i++) {
    const char = trimmed[i]!;
    if (char === '\\' && trimmed[i + 1] === '|') {
      current += '|';
      i++;
    } else if (char === '|') {
      cells.push(current.trim());
      current = '';
    } else {
      current += char;
    }
  }
  cells.push(current.trim());
  return cells;
}

/** True when `line` starts a new block and therefore ends a paragraph. */
function startsBlock(line: string): boolean {
  return (
    HEADING.test(line) ||
    FENCE.test(line) ||
    RULE.test(line) ||
    BLOCKQUOTE.test(line) ||
    /^( {0,3})([-*+]|\d{1,9}[.)]) +\S/.test(line)
  );
}

function parseBlocks(lines: string[], state: State, parent: Node): void {
  let i = 0;
  while (i < lines.length) {
    const line = lines[i]!;
    if (BLANK.test(line)) {
      i++;
      continue;
    }

    const fence = FENCE.exec(line);
    if (fence) {
      const marker = fence[2]!;
      const indent = fence[1]!.length;
      const body: string[] = [];
      i++;
      while (i < lines.length) {
        const current = lines[i]!;
        const close = /^ {0,3}(`{3,}|~{3,})[ \t]*$/.exec(current);
        if (close && close[1]![0] === marker[0] && close[1]!.length >= marker.length) break;
        body.push(current.replace(new RegExp(`^ {0,${indent}}`), ''));
        i++;
      }
      i++;
      const language = fence[3] ?? '';
      const code = h('code', language ? { attrs: { 'data-language': language } } : null);
      code.appendChild(highlightToFragment(body.join('\n'), language));
      parent.appendChild(h('pre', null, code));
      continue;
    }

    const heading = HEADING.exec(line);
    if (heading) {
      const element = h(`h${heading[1]!.length}` as 'h1');
      parseInline(heading[2] ?? '', state, element);
      parent.appendChild(element);
      i++;
      continue;
    }

    if (RULE.test(line)) {
      parent.appendChild(h('hr'));
      i++;
      continue;
    }

    if (BLOCKQUOTE.test(line)) {
      const inner: string[] = [];
      while (i < lines.length && !BLANK.test(lines[i]!)) {
        const quoted = BLOCKQUOTE.exec(lines[i]!);
        // Lazy continuation: a paragraph line without ">" still belongs to the quote.
        if (!quoted && startsBlock(lines[i]!)) break;
        inner.push(quoted ? quoted[1]! : lines[i]!);
        i++;
      }
      const quote = h('blockquote');
      parseBlocks(inner, state, quote);
      parent.appendChild(quote);
      continue;
    }

    const item = LIST_ITEM.exec(line);
    if (item) {
      i = parseList(lines, i, state, parent);
      continue;
    }

    if (/^(?: {4}|\t)/.test(line)) {
      const body: string[] = [];
      while (i < lines.length && (/^(?: {4}|\t)/.test(lines[i]!) || BLANK.test(lines[i]!))) {
        body.push(lines[i]!.replace(/^(?: {4}|\t)/, ''));
        i++;
      }
      while (body.length > 0 && BLANK.test(body[body.length - 1]!)) body.pop();
      parent.appendChild(h('pre', null, h('code', { text: body.join('\n') })));
      continue;
    }

    if (line.includes('|') && i + 1 < lines.length && TABLE_DIVIDER.test(lines[i + 1]!)) {
      const header = splitTableRow(line);
      const divider = splitTableRow(lines[i + 1]!);
      if (header.length === divider.length) {
        const aligns = divider.map((cell) =>
          cell.startsWith(':') && cell.endsWith(':')
            ? 'center'
            : cell.endsWith(':')
              ? 'right'
              : cell.startsWith(':')
                ? 'left'
                : undefined,
        );
        const makeRow = (cells: string[], tag: 'th' | 'td'): HTMLTableRowElement => {
          const row = h('tr');
          for (let col = 0; col < header.length; col++) {
            const cell = h(tag, { style: { 'text-align': aligns[col] } });
            parseInline(cells[col] ?? '', state, cell);
            row.appendChild(cell);
          }
          return row;
        };
        const tbody = h('tbody');
        i += 2;
        while (i < lines.length && !BLANK.test(lines[i]!) && lines[i]!.includes('|')) {
          tbody.appendChild(makeRow(splitTableRow(lines[i]!), 'td'));
          i++;
        }
        parent.appendChild(h('table', null, h('thead', null, makeRow(header, 'th')), tbody));
        continue;
      }
    }

    // Paragraph (possibly a setext heading).
    const paragraph: string[] = [line];
    i++;
    let setext = 0;
    while (i < lines.length && !BLANK.test(lines[i]!)) {
      const current = lines[i]!;
      const underline = SETEXT.exec(current);
      if (underline) {
        setext = underline[1]![0] === '=' ? 1 : 2;
        i++;
        break;
      }
      if (startsBlock(current)) break;
      paragraph.push(current);
      i++;
    }
    const element = setext ? h(`h${setext}` as 'h1') : h('p');
    parseInline(paragraph.map((text) => text.replace(/^ {1,3}/, '')).join('\n'), state, element);
    parent.appendChild(element);
  }
}

function parseList(lines: string[], start: number, state: State, parent: Node): number {
  const first = LIST_ITEM.exec(lines[start]!)!;
  const ordered = /\d/.test(first[2]!);
  const bullet = ordered ? first[2]!.slice(-1) : first[2]!;
  const list = h(ordered ? 'ol' : 'ul');
  if (ordered) {
    const startNumber = parseInt(first[2]!, 10);
    if (startNumber !== 1) list.setAttribute('start', String(startNumber));
  }
  let i = start;
  while (i < lines.length) {
    const match = LIST_ITEM.exec(lines[i]!);
    if (!match) break;
    const sameKind = ordered
      ? /\d/.test(match[2]!) && match[2]!.endsWith(bullet)
      : match[2] === bullet;
    if (!sameKind) break;
    // Content of an item is indented to where its text starts.
    const contentIndent = match[1]!.length + match[2]!.length + Math.min(match[3]!.length || 1, 4);
    const body: string[] = [match[4]!];
    i++;
    while (i < lines.length) {
      const current = lines[i]!;
      if (BLANK.test(current)) {
        // A blank line stays in the item only if indented content follows it.
        const next = lines[i + 1];
        if (
          next !== undefined &&
          !BLANK.test(next) &&
          /^\s/.test(next) &&
          indentOf(next) >= contentIndent
        ) {
          body.push('');
          i++;
          continue;
        }
        break;
      }
      if (indentOf(current) >= contentIndent) {
        body.push(current.slice(contentIndent));
      } else if (LIST_ITEM.test(current) || startsBlock(current)) {
        break;
      } else {
        body.push(current.trimStart());
      }
      i++;
    }
    const element = h('li');
    const task = /^\[([ xX])\][ \t]+/.exec(body[0]!);
    if (task) {
      body[0] = body[0]!.slice(task[0].length);
      element.className = 'av-task';
      element.appendChild(
        h('input', { attrs: { type: 'checkbox', disabled: true, checked: task[1] !== ' ' } }),
      );
    }
    const simple =
      body.every((text) => !BLANK.test(text)) && !body.slice(1).some(startsBlockOrList);
    if (simple) parseInline(body.join('\n'), state, element);
    else parseBlocks(body, state, element);
    list.appendChild(element);
    while (
      i < lines.length &&
      BLANK.test(lines[i]!) &&
      i + 1 < lines.length &&
      LIST_ITEM.test(lines[i + 1]!)
    ) {
      i++;
    }
  }
  parent.appendChild(list);
  return i;
}

function startsBlockOrList(line: string): boolean {
  return LIST_ITEM.test(line) || startsBlock(line);
}

function indentOf(line: string): number {
  let indent = 0;
  for (const char of line) {
    if (char === ' ') indent++;
    else if (char === '\t') indent += 4;
    else break;
  }
  return indent;
}

const ESCAPABLE = /[\\`*_{}[\]()#+\-.!|<>~"'&]/;
const AUTOLINK = /^<((?:https?:\/\/|mailto:)[^\s<>]+)>/;
const BARE_URL = /^https?:\/\/[^\s<>]*[^\s<>.,;:!?'")\]]/;
const ENTITY = /^&(?:#x[0-9a-fA-F]+|#\d+|[a-zA-Z]\w*);/;

/** Finds the `]` matching the `[` at `open`, honouring nesting, escapes and code spans. */
function findClosingBracket(text: string, open: number): number {
  let depth = 0;
  for (let i = open; i < text.length; i++) {
    const char = text[i];
    if (char === '\\') i++;
    else if (char === '`') {
      const end = text.indexOf('`', i + 1);
      if (end !== -1) i = end;
    } else if (char === '[') depth++;
    else if (char === ']' && --depth === 0) return i;
  }
  return -1;
}

/** Parses `(url "title")` starting at `start`; returns the pieces and the end index. */
function parseLinkTarget(
  text: string,
  start: number,
): { url: string; title: string | undefined; end: number } | null {
  if (text[start] !== '(') return null;
  let i = start + 1;
  while (text[i] === ' ' || text[i] === '\n') i++;
  let url: string;
  if (text[i] === '<') {
    const close = text.indexOf('>', i);
    if (close === -1) return null;
    url = text.slice(i + 1, close);
    i = close + 1;
  } else {
    let depth = 0;
    const from = i;
    for (; i < text.length; i++) {
      const char = text[i]!;
      if (char === '\\') i++;
      else if (char === '(') depth++;
      else if (char === ')') {
        if (depth === 0) break;
        depth--;
      } else if (char === ' ' || char === '\n') break;
    }
    url = text.slice(from, i);
  }
  while (text[i] === ' ' || text[i] === '\n') i++;
  let title: string | undefined;
  const quote = text[i];
  if (quote === '"' || quote === "'" || quote === '(') {
    const closer = quote === '(' ? ')' : quote;
    const close = text.indexOf(closer, i + 1);
    if (close === -1) return null;
    title = text.slice(i + 1, close);
    i = close + 1;
    while (text[i] === ' ' || text[i] === '\n') i++;
  }
  if (text[i] !== ')') return null;
  return { url: url.replace(/\\([()\\])/g, '$1'), title, end: i + 1 };
}

function parseInline(text: string, state: State, parent: Node): void {
  let buffer = '';
  const flush = (): void => {
    if (buffer) parent.appendChild(document.createTextNode(buffer));
    buffer = '';
  };
  let i = 0;
  while (i < text.length) {
    const char = text[i]!;

    if (char === '\\' && i + 1 < text.length) {
      const next = text[i + 1]!;
      if (next === '\n') {
        flush();
        parent.appendChild(h('br'));
        i += 2;
        continue;
      }
      if (ESCAPABLE.test(next)) {
        buffer += next;
        i += 2;
        continue;
      }
    }

    if (char === '`') {
      const run = /^`+/.exec(text.slice(i))![0];
      const end = text.indexOf(run, i + run.length);
      if (end !== -1 && text[end + run.length] !== '`') {
        flush();
        let code = text.slice(i + run.length, end).replace(/\n/g, ' ');
        if (/^ .* $/.test(code) && code.trim()) code = code.slice(1, -1);
        parent.appendChild(h('code', { text: code }));
        i = end + run.length;
        continue;
      }
      buffer += run;
      i += run.length;
      continue;
    }

    if (char === '!' && text[i + 1] === '[') {
      const close = findClosingBracket(text, i + 1);
      const target = close === -1 ? null : resolveTarget(text, i + 1, close, state);
      if (target) {
        flush();
        const alt = text.slice(i + 2, close).replace(/[*_`[\]]/g, '');
        const source = safeImageUrl(target.url, state.baseUrl);
        if (source) {
          parent.appendChild(
            h('img', { attrs: { src: source, alt, title: target.title, loading: 'lazy' } }),
          );
        } else {
          buffer += alt;
        }
        i = target.end;
        continue;
      }
    }

    if (char === '[') {
      const close = findClosingBracket(text, i);
      const target = close === -1 ? null : resolveTarget(text, i, close, state);
      if (target) {
        flush();
        const href = safeLinkUrl(target.url, state.baseUrl);
        const link = href ? externalLink(href) : h('span');
        if (target.title) link.title = target.title;
        parseInline(text.slice(i + 1, close), state, link);
        parent.appendChild(link);
        i = target.end;
        continue;
      }
    }

    if (char === '<') {
      const auto = AUTOLINK.exec(text.slice(i));
      if (auto) {
        flush();
        const href = safeLinkUrl(auto[1]);
        append(parent, href ? externalLink(href, auto[1]!.replace(/^mailto:/, '')) : auto[1]);
        i += auto[0].length;
        continue;
      }
      if (/^<br\s*\/?>/i.test(text.slice(i, i + 6))) {
        flush();
        parent.appendChild(h('br'));
        i += /^<br\s*\/?>/i.exec(text.slice(i, i + 6))![0].length;
        continue;
      }
    }

    if (char === 'h' && (i === 0 || /[\s(*_~]/.test(text[i - 1]!))) {
      const bare = BARE_URL.exec(text.slice(i));
      if (bare) {
        flush();
        const href = safeLinkUrl(bare[0]);
        append(parent, href ? externalLink(href, bare[0]) : bare[0]);
        i += bare[0].length;
        continue;
      }
    }

    if (char === '&') {
      const entity = ENTITY.exec(text.slice(i, i + 12));
      if (entity) {
        buffer += decodeEntities(entity[0]);
        i += entity[0].length;
        continue;
      }
    }

    if (char === '*' || char === '_' || char === '~') {
      const run = new RegExp(`^\\${char}+`).exec(text.slice(i))![0];
      const emphasis = parseEmphasis(text, i, char, run.length);
      if (emphasis) {
        flush();
        const element = h(emphasis.tag);
        parseInline(emphasis.inner, state, element);
        parent.appendChild(element);
        i = emphasis.end;
        continue;
      }
      buffer += run;
      i += run.length;
      continue;
    }

    if (char === '\n') {
      // Two trailing spaces before a newline make a hard line break.
      if (/ {2,}$/.test(buffer)) {
        buffer = buffer.replace(/ +$/, '');
        flush();
        parent.appendChild(h('br'));
      } else {
        buffer = buffer.replace(/ +$/, '') + '\n';
      }
      i++;
      continue;
    }

    buffer += char;
    i++;
  }
  flush();
}

function resolveTarget(
  text: string,
  open: number,
  close: number,
  state: State,
): { url: string; title: string | undefined; end: number } | null {
  const inline = parseLinkTarget(text, close + 1);
  if (inline) return inline;
  // Reference styles: [text][label], [text][] and [text].
  let label = text.slice(open + 1, close);
  let end = close + 1;
  if (text[end] === '[') {
    const labelClose = text.indexOf(']', end + 1);
    if (labelClose !== -1) {
      const explicit = text.slice(end + 1, labelClose);
      if (explicit.trim()) label = explicit;
      end = labelClose + 1;
    }
  }
  const reference = state.references.get(normaliseReference(label));
  return reference ? { ...reference, end } : null;
}

function parseEmphasis(
  text: string,
  start: number,
  char: string,
  runLength: number,
): { tag: 'em' | 'strong' | 'del'; inner: string; end: number } | null {
  if (char === '~' && runLength !== 2) return null;
  const size = char === '~' ? 2 : Math.min(runLength, 2);
  const after = text[start + runLength];
  if (after === undefined || /\s/.test(after)) return null;
  // Intra-word underscores (snake_case) are not emphasis.
  if (char === '_' && start > 0 && /\w/.test(text[start - 1]!)) return null;
  const delimiter = char.repeat(size);
  let search = start + runLength;
  for (;;) {
    const close = text.indexOf(delimiter, search);
    if (close === -1) return null;
    const before = text[close - 1]!;
    const following = text[close + size];
    const escaped = before === '\\';
    const valid =
      !escaped &&
      !/\s/.test(before) &&
      (char !== '_' || following === undefined || !/\w/.test(following)) &&
      // For single delimiters, do not stop on the first char of a longer run.
      !(size === 1 && char !== '~' && following === char && text[close + 2] !== char);
    if (valid && close > start + runLength - 1) {
      const innerStart = start + size;
      const inner = text.slice(innerStart, close);
      if (inner.length > 0) {
        return {
          tag: char === '~' ? 'del' : size === 2 ? 'strong' : 'em',
          inner,
          end: close + size,
        };
      }
    }
    search = close + 1;
  }
}

/** Renders Markdown source to a document fragment. */
export function renderMarkdown(source: string, options: MarkdownOptions = {}): DocumentFragment {
  const state: State = { baseUrl: options.baseUrl ?? null, references: new Map() };
  const lines: string[] = [];
  let inFence: string | null = null;
  for (const raw of stripBom(source).replace(/\r\n?/g, '\n').split('\n')) {
    const fence = /^ {0,3}(`{3,}|~{3,})/.exec(raw);
    if (fence) {
      if (inFence === null) inFence = fence[1]![0]!;
      else if (fence[1]![0] === inFence) inFence = null;
    }
    // Link reference definitions are collected up front and removed from the flow.
    const reference = inFence === null && !fence ? REFERENCE.exec(raw) : null;
    if (reference) {
      const key = normaliseReference(reference[1]!);
      if (!state.references.has(key)) {
        state.references.set(key, { url: reference[2]!, title: reference[3] });
      }
      continue;
    }
    lines.push(raw);
  }
  const fragment = document.createDocumentFragment();
  parseBlocks(lines, state, fragment);
  return fragment;
}

/** Removes a leading byte-order mark. */
function stripBom(text: string): string {
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}
