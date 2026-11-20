import { base64ToBytes, bytesToBinaryString, decodeText, decodeWith } from '../util/text';

/**
 * A MIME (RFC 5322 / RFC 2045) message parser: folded headers, encoded words,
 * nested multipart bodies, base64 and quoted-printable transfer encodings.
 */

export interface EmailAttachment {
  name: string;
  mimeType: string;
  data: Uint8Array;
}

export interface EmailMessage {
  /** Decoded headers, keyed by lower-case name (first occurrence wins). */
  headers: Map<string, string>;
  text: string | null;
  html: string | null;
  attachments: EmailAttachment[];
}

interface ContentHeader {
  value: string;
  params: Map<string, string>;
}

const MAX_DEPTH = 12;

function latin1ToBytes(text: string): Uint8Array {
  const bytes = new Uint8Array(text.length);
  for (let i = 0; i < text.length; i++) bytes[i] = text.charCodeAt(i) & 0xff;
  return bytes;
}

function decodeQuotedPrintable(text: string, headerMode: boolean): Uint8Array {
  const source = (headerMode ? text.replace(/_/g, ' ') : text).replace(/=\r?\n/g, '');
  const bytes: number[] = [];
  for (let i = 0; i < source.length; i++) {
    const char = source[i]!;
    if (char === '=' && /^[0-9a-fA-F]{2}$/.test(source.substr(i + 1, 2))) {
      bytes.push(parseInt(source.substr(i + 1, 2), 16));
      i += 2;
    } else {
      bytes.push(char.charCodeAt(0) & 0xff);
    }
  }
  return new Uint8Array(bytes);
}

/** Decodes RFC 2047 encoded words such as `=?utf-8?B?SGVsbG8=?=`. */
export function decodeHeaderValue(raw: string): string {
  // Raw 8-bit header bytes are in practice UTF-8.
  let value = /[\u0080-\u00ff]/.test(raw) ? decodeText(latin1ToBytes(raw)) : raw;
  // Whitespace between two adjacent encoded words is not part of the text.
  value = value.replace(/(\?=)\s+(=\?)/g, '$1$2');
  return value.replace(
    /=\?([^?\s]+)\?([bBqQ])\?([^?]*)\?=/g,
    (_match, charset: string, encoding: string, body: string) => {
      const bytes =
        encoding.toLowerCase() === 'b' ? base64ToBytes(body) : decodeQuotedPrintable(body, true);
      return decodeWith(bytes, charset.split('*')[0]!);
    },
  );
}

function parseHeaders(block: string): Map<string, string> {
  const headers = new Map<string, string>();
  const unfolded = block.replace(/\r?\n[ \t]+/g, ' ');
  for (const line of unfolded.split(/\r?\n/)) {
    const colon = line.indexOf(':');
    if (colon <= 0) continue;
    const name = line.slice(0, colon).trim().toLowerCase();
    if (!headers.has(name)) headers.set(name, line.slice(colon + 1).trim());
  }
  return headers;
}

/** Parses `type/subtype; name="value"; other=value`, including RFC 2231 `name*=` parameters. */
function parseContentHeader(raw: string | undefined): ContentHeader {
  const params = new Map<string, string>();
  if (!raw) return { value: '', params };
  const pieces = raw.match(/(?:[^;"]|"(?:[^"\\]|\\.)*")+/g) ?? [];
  const value = (pieces.shift() ?? '').trim().toLowerCase();
  const continued = new Map<string, string[]>();
  for (const piece of pieces) {
    const equals = piece.indexOf('=');
    if (equals === -1) continue;
    let name = piece.slice(0, equals).trim().toLowerCase();
    let text = piece.slice(equals + 1).trim();
    if (text.startsWith('"') && text.endsWith('"'))
      text = text.slice(1, -1).replace(/\\(.)/g, '$1');
    const extended = name.endsWith('*');
    if (extended) name = name.slice(0, -1);
    const part = /^(.*)\*(\d+)$/.exec(name);
    if (part) {
      const list = continued.get(part[1]!) ?? [];
      list[Number(part[2])] = extended ? decodeExtended(text, list.length === 0) : text;
      continued.set(part[1]!, list);
    } else {
      params.set(name, extended ? decodeExtended(text, true) : decodeHeaderValue(text));
    }
  }
  for (const [name, list] of continued) params.set(name, list.join(''));
  return { value, params };
}

/** RFC 2231: `utf-8''na%C3%AFve.txt`. */
function decodeExtended(text: string, hasCharset: boolean): string {
  const match = hasCharset ? /^([^']*)'[^']*'(.*)$/.exec(text) : null;
  const encoded = match ? match[2]! : text;
  const bytes: number[] = [];
  for (let i = 0; i < encoded.length; i++) {
    if (encoded[i] === '%' && /^[0-9a-fA-F]{2}$/.test(encoded.substr(i + 1, 2))) {
      bytes.push(parseInt(encoded.substr(i + 1, 2), 16));
      i += 2;
    } else {
      bytes.push(encoded.charCodeAt(i) & 0xff);
    }
  }
  return decodeWith(new Uint8Array(bytes), match?.[1] || 'utf-8');
}

function parsePart(raw: string, message: EmailMessage, depth: number, isRoot: boolean): void {
  const separator = /\r?\n\r?\n/.exec(raw);
  const headerBlock = separator ? raw.slice(0, separator.index) : raw;
  const body = separator ? raw.slice(separator.index + separator[0].length) : '';
  const headers = parseHeaders(headerBlock);
  if (isRoot) {
    for (const [name, value] of headers) message.headers.set(name, decodeHeaderValue(value));
  }
  const type = parseContentHeader(headers.get('content-type') ?? 'text/plain');
  const disposition = parseContentHeader(headers.get('content-disposition'));

  if (type.value.startsWith('multipart/') && depth < MAX_DEPTH) {
    const boundary = type.params.get('boundary');
    if (boundary) {
      const sections = body.split(`--${boundary}`);
      // The first section is the preamble; the one after the closing "--" is the epilogue.
      for (let i = 1; i < sections.length; i++) {
        const section = sections[i]!;
        if (section.startsWith('--')) break;
        parsePart(
          section.replace(/^[ \t]*\r?\n/, '').replace(/\r?\n$/, ''),
          message,
          depth + 1,
          false,
        );
      }
      return;
    }
  }

  const encoding = (headers.get('content-transfer-encoding') ?? '').trim().toLowerCase();
  const data =
    encoding === 'base64'
      ? base64ToBytes(body)
      : encoding === 'quoted-printable'
        ? decodeQuotedPrintable(body, false)
        : latin1ToBytes(body);
  const fileName = disposition.params.get('filename') ?? type.params.get('name');
  const isAttachment = disposition.value === 'attachment' || fileName !== undefined;
  const charset = type.params.get('charset') ?? 'utf-8';

  if (!isAttachment && type.value === 'text/plain') {
    message.text ??= /^utf-?8$/i.test(charset) ? decodeText(data) : decodeWith(data, charset);
  } else if (!isAttachment && type.value === 'text/html') {
    message.html ??= /^utf-?8$/i.test(charset) ? decodeText(data) : decodeWith(data, charset);
  } else if (data.length > 0) {
    const extension = type.value === 'message/rfc822' ? 'eml' : (type.value.split('/')[1] ?? 'bin');
    message.attachments.push({
      name:
        fileName ??
        `attachment-${message.attachments.length + 1}.${extension.replace(/[^\w]+/g, '') || 'bin'}`,
      mimeType: type.value || 'application/octet-stream',
      data,
    });
  }
}

/** Parses a raw `.eml` message. */
export function parseEmail(bytes: Uint8Array): EmailMessage {
  const message: EmailMessage = { headers: new Map(), text: null, html: null, attachments: [] };
  // A binary string keeps bytes and characters 1:1 so binary parts survive string handling.
  parsePart(bytesToBinaryString(bytes), message, 0, true);
  return message;
}
