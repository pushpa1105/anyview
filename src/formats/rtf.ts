import { AnyViewError } from '../core/errors';
import { h, setStyles } from '../ui/dom';
import { bytesToBinaryString, decodeWith, toBlobPart } from '../util/text';

/**
 * RTF → DOM. Handles character formatting (bold, italic, underline, strike,
 * size, colours, super/subscript), paragraph alignment and indents, Unicode
 * and code-page text, simple tables and embedded PNG/JPEG pictures.
 */

export interface RtfHost {
  objectUrl(blob: Blob): string;
}

interface CharState {
  bold: boolean;
  italic: boolean;
  underline: boolean;
  strike: boolean;
  hidden: boolean;
  /** Font size in half-points. */
  size: number;
  color: number;
  background: number;
  vertical: '' | 'super' | 'sub';
  /** Number of fallback characters to skip after a `\u` escape. */
  unicodeSkip: number;
  /** What the enclosing group is: normal content or a destination with special handling. */
  destination: 'content' | 'skip' | 'colortbl' | 'pict';
}

interface ParaState {
  align: string | undefined;
  leftIndent: number;
  firstIndent: number;
  spaceBefore: number;
  spaceAfter: number;
  inTable: boolean;
}

/** Destinations whose content is never shown. */
const SKIPPED = new Set([
  'fonttbl',
  'stylesheet',
  'info',
  'header',
  'headerl',
  'headerr',
  'headerf',
  'footer',
  'footerl',
  'footerr',
  'footerf',
  'footnote',
  'generator',
  'object',
  'themedata',
  'colorschememapping',
  'latentstyles',
  'datastore',
  'xmlnstbl',
  'listtable',
  'listoverridetable',
  'rsidtbl',
  'fldinst',
  'nonshppict',
  'filetbl',
  'revtbl',
  'pgdsctbl',
  'bkmkstart',
  'bkmkend',
  'annotation',
  'xe',
  'tc',
  'comment',
  'operator',
  'author',
  'title',
  'subject',
  'keywords',
  'doccomm',
  'company',
]);

const SYMBOLS: Readonly<Record<string, string>> = {
  emdash: '—',
  endash: '–',
  lquote: '‘',
  rquote: '’',
  ldblquote: '“',
  rdblquote: '”',
  bullet: '•',
  emspace: '\u2003',
  enspace: '\u2002',
  tab: '\t',
  '~': '\u00a0',
  _: '\u2011',
};

const DEFAULT_CHAR: CharState = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  hidden: false,
  size: 24,
  color: 0,
  background: 0,
  vertical: '',
  unicodeSkip: 1,
  destination: 'content',
};

const DEFAULT_PARA: ParaState = {
  align: undefined,
  leftIndent: 0,
  firstIndent: 0,
  spaceBefore: 0,
  spaceAfter: 0,
  inTable: false,
};

/** Renders an RTF document (given as raw bytes) into a detached element. */
export function renderRtf(bytes: Uint8Array, host: RtfHost): HTMLElement {
  // RTF is 7-bit ASCII with escapes; Latin-1 keeps a 1:1 byte ↔ char mapping for the tokenizer.
  const source = bytesToBinaryString(bytes);
  if (!source.trimStart().startsWith('{\\rtf')) {
    throw new AnyViewError('PARSE_FAILED', 'This file is not a valid RTF document.');
  }

  const paper = h('div', {
    style: {
      width: '794px',
      'min-height': '1123px',
      padding: '76px',
      'font-family': '"Times New Roman", serif',
    },
  });
  const colors: (string | undefined)[] = [];
  let codePage = 'windows-1252';
  let char: CharState = { ...DEFAULT_CHAR };
  let para: ParaState = { ...DEFAULT_PARA };
  const stack: CharState[] = [];

  let paragraph: HTMLElement | null = null;
  let lastSpan: HTMLElement | null = null;
  let lastKey = '';
  let pendingBytes: number[] = [];
  let skipCount = 0;

  let table: HTMLElement | null = null;
  let rowCells: HTMLElement[] = [];
  let cellBlocks: HTMLElement[] = [];

  let colorParts = { red: 0, green: 0, blue: 0, defined: false };
  let picture = { hex: '', type: '', width: 0, height: 0 };

  const ensureParagraph = (): HTMLElement => {
    if (!paragraph) {
      const px = (twips: number): string | undefined =>
        twips ? `${(twips / 15).toFixed(1)}px` : undefined;
      paragraph = h('p', {
        style: {
          'text-align': para.align,
          'margin-left': px(para.leftIndent),
          'text-indent': px(para.firstIndent),
          'margin-top': px(para.spaceBefore),
          'margin-bottom': px(para.spaceAfter),
          'white-space': 'pre-wrap',
        },
      });
      lastSpan = null;
      lastKey = '';
    }
    return paragraph;
  };

  const flushBytes = (): void => {
    if (pendingBytes.length === 0) return;
    const text = decodeWith(new Uint8Array(pendingBytes), codePage);
    pendingBytes = [];
    emit(text);
  };

  const emit = (text: string): void => {
    if (text === '' || char.destination !== 'content' || char.hidden) return;
    const key = `${+char.bold}${+char.italic}${+char.underline}${+char.strike}|${char.size}|${char.color}|${char.background}|${char.vertical}`;
    const target = ensureParagraph();
    if (lastSpan && key === lastKey) {
      lastSpan.textContent += text;
      return;
    }
    const decoration = [char.underline && 'underline', char.strike && 'line-through']
      .filter(Boolean)
      .join(' ');
    lastSpan = h('span', {
      text,
      style: {
        'font-weight': char.bold ? 'bold' : undefined,
        'font-style': char.italic ? 'italic' : undefined,
        'text-decoration': decoration || undefined,
        'font-size': `${((char.vertical ? 0.7 : 1) * char.size) / 2}pt`,
        color: colors[char.color],
        'background-color': colors[char.background],
        'vertical-align': char.vertical || undefined,
      },
    });
    lastKey = key;
    target.appendChild(lastSpan);
  };

  const endParagraph = (force: boolean): void => {
    flushBytes();
    if (!paragraph && !force) return;
    const finished = ensureParagraph();
    paragraph = null;
    if (para.inTable) {
      cellBlocks.push(finished);
    } else {
      table = null;
      paper.appendChild(finished);
    }
  };

  const endCell = (): void => {
    endParagraph(false);
    const cell = h('td', { style: { border: '1px solid #a6a6a6' } });
    for (const block of cellBlocks) cell.appendChild(block);
    cellBlocks = [];
    rowCells.push(cell);
  };

  const endRow = (): void => {
    if (rowCells.length === 0) return;
    if (!table) {
      table = h('tbody');
      paper.appendChild(h('table', null, table));
    }
    const row = h('tr');
    for (const cell of rowCells) row.appendChild(cell);
    rowCells = [];
    table.appendChild(row);
  };

  const finishPicture = (): void => {
    const type = picture.type;
    const hex = picture.hex;
    const { width, height } = picture;
    picture = { hex: '', type: '', width: 0, height: 0 };
    if (!type || hex.length < 2) return;
    const data = new Uint8Array(hex.length >> 1);
    for (let i = 0; i < data.length; i++) data[i] = parseInt(hex.substr(i * 2, 2), 16);
    const image = h('img', {
      attrs: { alt: '', src: host.objectUrl(new Blob([toBlobPart(data)], { type })) },
    });
    setStyles(image, {
      width: width ? `${(width / 15).toFixed(0)}px` : undefined,
      height: height ? `${(height / 15).toFixed(0)}px` : undefined,
    });
    ensureParagraph().appendChild(image);
    lastSpan = null;
  };

  const control = (word: string, parameter: number | null, starred: boolean): void => {
    const on = parameter !== 0;
    if (char.destination === 'colortbl') {
      if (word === 'red') colorParts = { ...colorParts, red: parameter ?? 0, defined: true };
      else if (word === 'green')
        colorParts = { ...colorParts, green: parameter ?? 0, defined: true };
      else if (word === 'blue') colorParts = { ...colorParts, blue: parameter ?? 0, defined: true };
      return;
    }
    if (char.destination === 'pict') {
      if (word === 'pngblip') picture.type = 'image/png';
      else if (word === 'jpegblip') picture.type = 'image/jpeg';
      else if (word === 'picwgoal') picture.width = parameter ?? 0;
      else if (word === 'pichgoal') picture.height = parameter ?? 0;
      return;
    }
    if (char.destination === 'skip') return;

    if (word === 'colortbl') {
      char.destination = 'colortbl';
      return;
    }
    if (word === 'pict') {
      flushBytes();
      char.destination = 'pict';
      return;
    }
    if (SKIPPED.has(word) || (starred && word !== 'shppict')) {
      flushBytes();
      char.destination = 'skip';
      return;
    }
    if (word in SYMBOLS) {
      flushBytes();
      emit(SYMBOLS[word]!);
      return;
    }
    switch (word) {
      case 'par':
      case 'sect':
        endParagraph(true);
        break;
      case 'line':
        flushBytes();
        ensureParagraph().appendChild(h('br'));
        lastSpan = null;
        break;
      case 'pard':
        flushBytes();
        para = { ...DEFAULT_PARA };
        break;
      case 'plain':
        flushBytes();
        char = { ...DEFAULT_CHAR, unicodeSkip: char.unicodeSkip, destination: char.destination };
        break;
      case 'b':
        flushBytes();
        char.bold = on;
        break;
      case 'i':
        flushBytes();
        char.italic = on;
        break;
      case 'ul':
      case 'uldb':
      case 'ulw':
        flushBytes();
        char.underline = on;
        break;
      case 'ulnone':
        flushBytes();
        char.underline = false;
        break;
      case 'strike':
        flushBytes();
        char.strike = on;
        break;
      case 'v':
        flushBytes();
        char.hidden = on;
        break;
      case 'fs':
        flushBytes();
        if (parameter && parameter > 0) char.size = parameter;
        break;
      case 'cf':
        flushBytes();
        char.color = parameter ?? 0;
        break;
      case 'highlight':
      case 'cb':
        flushBytes();
        char.background = parameter ?? 0;
        break;
      case 'super':
        flushBytes();
        char.vertical = 'super';
        break;
      case 'sub':
        flushBytes();
        char.vertical = 'sub';
        break;
      case 'nosupersub':
        flushBytes();
        char.vertical = '';
        break;
      case 'qc':
        para.align = 'center';
        break;
      case 'qr':
        para.align = 'right';
        break;
      case 'qj':
        para.align = 'justify';
        break;
      case 'ql':
        para.align = undefined;
        break;
      case 'li':
        para.leftIndent = parameter ?? 0;
        break;
      case 'fi':
        para.firstIndent = parameter ?? 0;
        break;
      case 'sb':
        para.spaceBefore = parameter ?? 0;
        break;
      case 'sa':
        para.spaceAfter = parameter ?? 0;
        break;
      case 'intbl':
        para.inTable = true;
        break;
      case 'cell':
        endCell();
        break;
      case 'row':
        endRow();
        break;
      case 'ansicpg':
        if (parameter) codePage = `windows-${parameter}`;
        break;
      case 'uc':
        char.unicodeSkip = parameter ?? 1;
        break;
      case 'u': {
        flushBytes();
        let code = parameter ?? 0;
        if (code < 0) code += 65536;
        emit(String.fromCharCode(code));
        skipCount = char.unicodeSkip;
        break;
      }
    }
  };

  const length = source.length;
  let i = 0;
  let depth = 0;
  while (i < length) {
    const current = source[i]!;
    if (current === '{') {
      flushBytes();
      stack.push({ ...char });
      depth++;
      i++;
    } else if (current === '}') {
      flushBytes();
      if (char.destination === 'pict') {
        const outer = stack[stack.length - 1];
        if (outer && outer.destination !== 'pict') finishPicture();
      }
      char = stack.pop() ?? { ...DEFAULT_CHAR };
      depth--;
      i++;
      if (depth <= 0) break;
    } else if (current === '\\') {
      const next = source[i + 1] ?? '';
      if (/[a-zA-Z]/.test(next)) {
        let end = i + 1;
        while (end < length && /[a-zA-Z]/.test(source[end]!)) end++;
        const word = source.slice(i + 1, end);
        let parameter: number | null = null;
        const numeric = /^-?\d+/.exec(source.slice(end, end + 12));
        if (numeric) {
          parameter = Number(numeric[0]);
          end += numeric[0].length;
        }
        if (source[end] === ' ') end++;
        i = end;
        if (skipCount > 0 && word !== 'u') {
          skipCount--;
          continue;
        }
        if (word === 'bin' && parameter) {
          i += parameter;
          continue;
        }
        control(word, parameter, false);
      } else if (next === "'") {
        const byte = parseInt(source.substr(i + 2, 2), 16);
        i += 4;
        if (skipCount > 0) skipCount--;
        else if (Number.isFinite(byte) && char.destination === 'content') pendingBytes.push(byte);
      } else if (next === '*') {
        // "\*\word": a destination that readers may ignore when they do not know it.
        const match = /^\\\*\s*\\([a-zA-Z]+)(-?\d+)? ?/.exec(source.slice(i, i + 64));
        if (match) {
          i += match[0].length;
          control(match[1]!, match[2] ? Number(match[2]) : null, true);
        } else {
          i += 2;
        }
      } else if (next === '\n' || next === '\r') {
        endParagraph(true);
        i += 2;
      } else if (next === '{' || next === '}' || next === '\\') {
        flushBytes();
        emit(next);
        i += 2;
      } else if (next in SYMBOLS) {
        flushBytes();
        emit(SYMBOLS[next]!);
        i += 2;
      } else {
        i += 2;
      }
    } else if (current === '\r' || current === '\n') {
      i++;
    } else {
      // Plain run: take everything up to the next special character in one slice.
      let end = i + 1;
      while (end < length) {
        const code = source.charCodeAt(end);
        if (code === 0x7b || code === 0x7d || code === 0x5c || code === 0x0a || code === 0x0d)
          break;
        end++;
      }
      let text = source.slice(i, end);
      i = end;
      if (skipCount > 0) {
        const skipped = Math.min(skipCount, text.length);
        text = text.slice(skipped);
        skipCount -= skipped;
      }
      if (char.destination === 'colortbl') {
        // Each ";" closes one colour-table entry; an empty entry means "automatic".
        for (const symbol of text) {
          if (symbol !== ';') continue;
          colors.push(
            colorParts.defined
              ? `rgb(${colorParts.red & 255}, ${colorParts.green & 255}, ${colorParts.blue & 255})`
              : undefined,
          );
          colorParts = { red: 0, green: 0, blue: 0, defined: false };
        }
      } else if (char.destination === 'pict') {
        picture.hex += text.replace(/[^0-9a-fA-F]/g, '');
      } else if (text) {
        flushBytes();
        emit(text);
      }
    }
  }
  endParagraph(false);
  if (rowCells.length > 0 || cellBlocks.length > 0) {
    if (cellBlocks.length > 0) endCell();
    endRow();
  }
  return paper;
}
