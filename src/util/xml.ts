/**
 * A small, dependency-free XML parser tuned for document formats (OOXML, ODF,
 * EPUB). It is not a validating parser: DTDs and processing instructions are
 * skipped, and custom entities are left as-is.
 */

export interface XmlElement {
  /** Qualified name as written, e.g. `"w:p"`. */
  readonly name: string;
  /** Name without the namespace prefix, e.g. `"p"`. */
  readonly local: string;
  /** Attributes keyed by qualified name, e.g. `"r:id"`. */
  readonly attrs: Readonly<Record<string, string>>;
  readonly children: XmlNode[];
}

export type XmlNode = XmlElement | string;

export interface ParseXmlOptions {
  /**
   * Keep whitespace-only text between elements. Needed for mixed content such
   * as ODF paragraphs; off by default to keep large OOXML trees small.
   */
  preserveWhitespace?: boolean;
}

const NAMED_ENTITIES: Readonly<Record<string, string>> = {
  lt: '<',
  gt: '>',
  amp: '&',
  quot: '"',
  apos: "'",
  nbsp: '\u00a0',
};

/** Replaces character and predefined entity references. */
export function decodeEntities(text: string): string {
  if (!text.includes('&')) return text;
  return text.replace(/&(#x[0-9a-fA-F]+|#\d+|[a-zA-Z]\w*);/g, (match, body: string) => {
    if (body.startsWith('#')) {
      const code =
        body[1] === 'x' || body[1] === 'X'
          ? parseInt(body.slice(2), 16)
          : parseInt(body.slice(1), 10);
      if (!Number.isFinite(code) || code <= 0 || code > 0x10ffff) return '';
      try {
        return String.fromCodePoint(code);
      } catch {
        return '';
      }
    }
    return NAMED_ENTITIES[body] ?? match;
  });
}

const ATTRIBUTE = /([^\s=/<>]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g;
const WHITESPACE_ONLY = /^[ \t\r\n]*$/;

function localName(name: string): string {
  const colon = name.indexOf(':');
  return colon === -1 ? name : name.slice(colon + 1);
}

function parseAttributes(source: string): Record<string, string> {
  const attrs: Record<string, string> = {};
  if (!source.includes('=')) return attrs;
  ATTRIBUTE.lastIndex = 0;
  let match: RegExpExecArray | null;
  while ((match = ATTRIBUTE.exec(source)) !== null) {
    attrs[match[1]!] = decodeEntities(match[2] ?? match[3] ?? '');
  }
  return attrs;
}

/** Parses an XML document and returns its root element. Throws on malformed input. */
export function parseXml(source: string, options: ParseXmlOptions = {}): XmlElement {
  const keepWhitespace = options.preserveWhitespace === true;
  const stack: XmlElement[] = [];
  let root: XmlElement | null = null;
  let position = source.charCodeAt(0) === 0xfeff ? 1 : 0;
  const length = source.length;

  const pushText = (text: string): void => {
    const parent = stack[stack.length - 1];
    if (!parent || text === '') return;
    parent.children.push(text);
  };

  const closeElement = (element: XmlElement): void => {
    if (keepWhitespace) return;
    const { children } = element;
    if (!children.some((child) => typeof child !== 'string')) return;
    // Mixed content: indentation between child elements is not significant.
    let write = 0;
    for (const child of children) {
      if (typeof child === 'string' && WHITESPACE_ONLY.test(child)) continue;
      children[write++] = child;
    }
    children.length = write;
  };

  while (position < length) {
    const open = source.indexOf('<', position);
    if (open === -1) {
      pushText(decodeEntities(source.slice(position)));
      break;
    }
    if (open > position) pushText(decodeEntities(source.slice(position, open)));

    if (source.startsWith('<!--', open)) {
      const end = source.indexOf('-->', open + 4);
      if (end === -1) throw new Error('Malformed XML: unterminated comment');
      position = end + 3;
    } else if (source.startsWith('<![CDATA[', open)) {
      const end = source.indexOf(']]>', open + 9);
      if (end === -1) throw new Error('Malformed XML: unterminated CDATA section');
      pushText(source.slice(open + 9, end));
      position = end + 3;
    } else if (source.startsWith('<?', open)) {
      const end = source.indexOf('?>', open + 2);
      if (end === -1) throw new Error('Malformed XML: unterminated processing instruction');
      position = end + 2;
    } else if (source.startsWith('<!', open)) {
      // DOCTYPE, possibly with an internal subset in square brackets.
      let depth = 0;
      let cursor = open + 2;
      for (; cursor < length; cursor++) {
        const char = source[cursor];
        if (char === '[') depth++;
        else if (char === ']') depth--;
        else if (char === '>' && depth <= 0) break;
      }
      if (cursor >= length) throw new Error('Malformed XML: unterminated declaration');
      position = cursor + 1;
    } else if (source[open + 1] === '/') {
      const end = source.indexOf('>', open + 2);
      if (end === -1) throw new Error('Malformed XML: unterminated closing tag');
      const name = source.slice(open + 2, end).trim();
      const element = stack.pop();
      if (!element || element.name !== name) {
        throw new Error(`Malformed XML: unexpected closing tag </${name}>`);
      }
      closeElement(element);
      position = end + 1;
    } else {
      const end = findTagEnd(source, open + 1);
      if (end === -1) throw new Error('Malformed XML: unterminated tag');
      const selfClosing = source[end - 1] === '/';
      const body = source.slice(open + 1, selfClosing ? end - 1 : end);
      const nameEnd = body.search(/[\s/]/);
      const name = nameEnd === -1 ? body : body.slice(0, nameEnd);
      if (name === '') throw new Error('Malformed XML: empty tag name');
      const element: XmlElement = {
        name,
        local: localName(name),
        attrs: nameEnd === -1 ? {} : parseAttributes(body.slice(nameEnd)),
        children: [],
      };
      const parent = stack[stack.length - 1];
      if (parent) parent.children.push(element);
      else if (root === null) root = element;
      else throw new Error('Malformed XML: more than one root element');
      if (!selfClosing) stack.push(element);
      position = end + 1;
    }
  }

  if (stack.length > 0) {
    throw new Error(`Malformed XML: <${stack[stack.length - 1]!.name}> is never closed`);
  }
  if (root === null) throw new Error('Malformed XML: no root element');
  return root;
}

/** Finds the `>` that ends a start tag, skipping over quoted attribute values. */
function findTagEnd(source: string, from: number): number {
  let quote = '';
  for (let i = from; i < source.length; i++) {
    const char = source[i]!;
    if (quote) {
      if (char === quote) quote = '';
    } else if (char === '"' || char === "'") {
      quote = char;
    } else if (char === '>') {
      return i;
    }
  }
  return -1;
}

export function isElement(node: XmlNode | undefined): node is XmlElement {
  return typeof node === 'object';
}

/** Child elements, optionally filtered by local name. */
export function childElements(
  element: XmlElement | null | undefined,
  local?: string,
): XmlElement[] {
  if (!element) return [];
  const result: XmlElement[] = [];
  for (const child of element.children) {
    if (typeof child !== 'string' && (local === undefined || child.local === local)) {
      result.push(child);
    }
  }
  return result;
}

/** First child element with the given local name. */
export function child(element: XmlElement | null | undefined, local: string): XmlElement | null {
  if (!element) return null;
  for (const node of element.children) {
    if (typeof node !== 'string' && node.local === local) return node;
  }
  return null;
}

/** Follows a path of local names, e.g. `find(run, 'rPr', 'sz')`. */
export function find(element: XmlElement | null | undefined, ...path: string[]): XmlElement | null {
  let current: XmlElement | null = element ?? null;
  for (const local of path) {
    current = child(current, local);
    if (!current) return null;
  }
  return current;
}

/**
 * Reads an attribute. A prefixed name (`"w:val"`) is matched exactly first and
 * then by local name under any prefix, because producers are free to choose
 * prefixes. An unprefixed name only matches an unprefixed attribute.
 */
export function attr(element: XmlElement | null | undefined, name: string): string | undefined {
  if (!element) return undefined;
  const exact = element.attrs[name];
  if (exact !== undefined) return exact;
  const colon = name.indexOf(':');
  if (colon === -1) return undefined;
  const suffix = name.slice(colon);
  for (const key in element.attrs) {
    if (key.endsWith(suffix)) return element.attrs[key];
  }
  return undefined;
}

/** Concatenated text of an element and all its descendants. */
export function textOf(element: XmlElement | null | undefined): string {
  if (!element) return '';
  let text = '';
  for (const node of element.children) {
    text += typeof node === 'string' ? node : textOf(node);
  }
  return text;
}

/** All descendant elements with the given local name, in document order. */
export function descendants(element: XmlElement | null | undefined, local: string): XmlElement[] {
  const result: XmlElement[] = [];
  const walk = (current: XmlElement): void => {
    for (const node of current.children) {
      if (typeof node === 'string') continue;
      if (node.local === local) result.push(node);
      walk(node);
    }
  };
  if (element) walk(element);
  return result;
}
