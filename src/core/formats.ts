import { looksLikeText } from '../util/text';
import { extensionOf } from '../util/url';
import { openZip } from '../util/zip';
import type { FormatCategory, FormatInfo } from './types';

function format(
  id: string,
  label: string,
  category: FormatCategory,
  extensions: string,
  mimeTypes = '',
): FormatInfo {
  return Object.freeze({
    id,
    label,
    category,
    extensions: Object.freeze(extensions.split(' ').filter(Boolean)),
    mimeTypes: Object.freeze(mimeTypes.split(' ').filter(Boolean)),
  });
}

/** Source-code extensions and the highlighter language they map to. */
const CODE_LANGUAGES: Readonly<Record<string, string>> = {
  js: 'javascript',
  mjs: 'javascript',
  cjs: 'javascript',
  jsx: 'javascript',
  ts: 'typescript',
  mts: 'typescript',
  cts: 'typescript',
  tsx: 'typescript',
  py: 'python',
  pyw: 'python',
  java: 'java',
  c: 'c',
  h: 'c',
  cpp: 'cpp',
  cc: 'cpp',
  cxx: 'cpp',
  hpp: 'cpp',
  cs: 'csharp',
  go: 'go',
  rs: 'rust',
  php: 'php',
  rb: 'ruby',
  swift: 'swift',
  kt: 'kotlin',
  kts: 'kotlin',
  scala: 'scala',
  dart: 'dart',
  lua: 'lua',
  r: 'r',
  pl: 'perl',
  sh: 'shell',
  bash: 'shell',
  zsh: 'shell',
  fish: 'shell',
  ps1: 'powershell',
  bat: 'batch',
  cmd: 'batch',
  sql: 'sql',
  css: 'css',
  scss: 'css',
  sass: 'css',
  less: 'css',
  yaml: 'yaml',
  yml: 'yaml',
  toml: 'toml',
  ini: 'ini',
  cfg: 'ini',
  conf: 'ini',
  properties: 'ini',
  env: 'ini',
  vue: 'html',
  svelte: 'html',
  astro: 'html',
  graphql: 'graphql',
  gql: 'graphql',
  dockerfile: 'dockerfile',
  makefile: 'makefile',
  mk: 'makefile',
  gradle: 'kotlin',
  tex: 'tex',
  diff: 'diff',
  patch: 'diff',
};

const PLAIN_TEXT_EXTENSIONS =
  'txt text log srt vtt sub lrc nfo gitignore gitattributes editorconfig';

/** Every format AnyView recognises. Formats without a renderer show the download fallback. */
export const FORMATS: readonly FormatInfo[] = Object.freeze([
  format('pdf', 'PDF document', 'document', 'pdf', 'application/pdf'),
  format(
    'image',
    'Image',
    'image',
    'png jpg jpeg jfif gif webp avif bmp ico apng',
    'image/png image/jpeg image/gif image/webp image/avif image/bmp image/x-icon image/vnd.microsoft.icon image/apng',
  ),
  format('svg', 'SVG image', 'image', 'svg', 'image/svg+xml'),
  format(
    'video',
    'Video',
    'video',
    'mp4 m4v webm ogv mov',
    'video/mp4 video/webm video/ogg video/quicktime',
  ),
  format(
    'audio',
    'Audio',
    'audio',
    'mp3 wav ogg oga m4a aac flac opus weba',
    'audio/mpeg audio/wav audio/x-wav audio/ogg audio/mp4 audio/aac audio/flac audio/webm audio/opus',
  ),
  format(
    'docx',
    'Word document',
    'document',
    'docx docm dotx dotm',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document application/vnd.openxmlformats-officedocument.wordprocessingml.template',
  ),
  format(
    'xlsx',
    'Excel workbook',
    'spreadsheet',
    'xlsx xlsm xltx xltm',
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet application/vnd.openxmlformats-officedocument.spreadsheetml.template',
  ),
  format(
    'pptx',
    'PowerPoint presentation',
    'presentation',
    'pptx pptm ppsx ppsm potx potm',
    'application/vnd.openxmlformats-officedocument.presentationml.presentation application/vnd.openxmlformats-officedocument.presentationml.slideshow',
  ),
  format(
    'odt',
    'OpenDocument text',
    'document',
    'odt ott',
    'application/vnd.oasis.opendocument.text',
  ),
  format(
    'ods',
    'OpenDocument spreadsheet',
    'spreadsheet',
    'ods ots',
    'application/vnd.oasis.opendocument.spreadsheet',
  ),
  format(
    'odp',
    'OpenDocument presentation',
    'presentation',
    'odp otp',
    'application/vnd.oasis.opendocument.presentation',
  ),
  format('csv', 'CSV table', 'spreadsheet', 'csv', 'text/csv application/csv'),
  format('tsv', 'TSV table', 'spreadsheet', 'tsv tab', 'text/tab-separated-values'),
  format('markdown', 'Markdown', 'document', 'md markdown mdown mkd mdx', 'text/markdown'),
  format('html', 'HTML page', 'document', 'html htm xhtml', 'text/html application/xhtml+xml'),
  format(
    'json',
    'JSON',
    'text',
    'json jsonc json5 geojson webmanifest map har',
    'application/json application/ld+json application/geo+json application/manifest+json',
  ),
  format(
    'xml',
    'XML',
    'text',
    'xml xsd xsl xslt rss atom plist kml gpx wsdl xaml csproj',
    'application/xml text/xml application/rss+xml application/atom+xml',
  ),
  format('rtf', 'Rich Text document', 'document', 'rtf', 'application/rtf text/rtf'),
  format('epub', 'EPUB e-book', 'document', 'epub', 'application/epub+zip'),
  format(
    'zip',
    'ZIP archive',
    'archive',
    'zip jar',
    'application/zip application/x-zip-compressed application/java-archive',
  ),
  format('ipynb', 'Jupyter notebook', 'document', 'ipynb', 'application/x-ipynb+json'),
  format('eml', 'Email message', 'document', 'eml', 'message/rfc822'),
  format(
    'font',
    'Font',
    'font',
    'ttf otf woff woff2',
    'font/ttf font/otf font/woff font/woff2 application/font-woff application/x-font-ttf',
  ),
  format(
    'text',
    'Text',
    'text',
    `${PLAIN_TEXT_EXTENSIONS} ${Object.keys(CODE_LANGUAGES).join(' ')}`,
    'text/plain text/css text/javascript application/javascript application/x-sh application/x-yaml text/yaml application/toml application/sql text/x-python',
  ),
  // Recognised so the user gets an accurate message, but not renderable in a browser.
  format('doc', 'Word 97–2003 document', 'document', 'doc dot', 'application/msword'),
  format('xls', 'Excel 97–2003 workbook', 'spreadsheet', 'xls xlt', 'application/vnd.ms-excel'),
  format(
    'ppt',
    'PowerPoint 97–2003 presentation',
    'presentation',
    'ppt pps pot',
    'application/vnd.ms-powerpoint',
  ),
  format('msg', 'Outlook message', 'document', 'msg', 'application/vnd.ms-outlook'),
  format('heic', 'HEIC image', 'image', 'heic heif', 'image/heic image/heif'),
  format('tiff', 'TIFF image', 'image', 'tif tiff', 'image/tiff'),
  format('ole', 'Legacy Microsoft Office file', 'document', '', ''),
]);

const BY_ID = new Map<string, FormatInfo>();
const BY_EXTENSION = new Map<string, FormatInfo>();
const BY_MIME = new Map<string, FormatInfo>();
for (const info of FORMATS) {
  BY_ID.set(info.id, info);
  for (const extension of info.extensions) BY_EXTENSION.set(extension, info);
  for (const mimeType of info.mimeTypes) BY_MIME.set(mimeType, info);
}

/** A generic format used for files AnyView cannot identify. */
export const UNKNOWN_FORMAT: FormatInfo = format('unknown', 'File', 'other', '');

export function getFormatById(id: string): FormatInfo | null {
  return BY_ID.get(id) ?? null;
}

export function getFormatByExtension(extension: string): FormatInfo | null {
  return BY_EXTENSION.get(extension.replace(/^\./, '').toLowerCase()) ?? null;
}

export function getFormatByMimeType(mimeType: string): FormatInfo | null {
  const essence = mimeType.split(';')[0]!.trim().toLowerCase();
  if (!essence || essence === 'application/octet-stream' || essence === 'binary/octet-stream') {
    return null;
  }
  const exact = BY_MIME.get(essence);
  if (exact) return exact;
  if (essence.endsWith('+json')) return BY_ID.get('json')!;
  if (essence.endsWith('+xml')) return BY_ID.get('xml')!;
  const family = essence.split('/')[0];
  if (family === 'video') return BY_ID.get('video')!;
  if (family === 'audio') return BY_ID.get('audio')!;
  if (family === 'text') return BY_ID.get('text')!;
  return null;
}

/** Resolves a user-supplied `type`: a format id, an extension (with or without dot) or a MIME type. */
export function resolveFormatHint(hint: string): FormatInfo | null {
  const value = hint.trim().toLowerCase();
  if (!value) return null;
  if (value.includes('/')) return getFormatByMimeType(value);
  return getFormatById(value) ?? getFormatByExtension(value);
}

/** Highlighter language for a file name or type hint, or `null`. */
export function languageFor(name: string, formatId: string, hint?: string): string | null {
  if (formatId === 'json' || formatId === 'xml') return formatId;
  if (formatId === 'html') return 'html';
  if (formatId === 'svg') return 'xml';
  const hinted = hint ? CODE_LANGUAGES[hint.replace(/^\./, '').toLowerCase()] : undefined;
  if (hinted) return hinted;
  const extension = extensionOf(name);
  if (CODE_LANGUAGES[extension]) return CODE_LANGUAGES[extension];
  const base = (name.split(/[\\/]/).pop() ?? '').toLowerCase();
  if (base === 'dockerfile' || base === 'makefile') return base;
  return null;
}

function startsWith(bytes: Uint8Array, signature: readonly number[], offset = 0): boolean {
  if (bytes.length < offset + signature.length) return false;
  for (let i = 0; i < signature.length; i++) {
    if (bytes[offset + i] !== signature[i]) return false;
  }
  return true;
}

function ascii(bytes: Uint8Array, start: number, end: number): string {
  let text = '';
  for (let i = start; i < Math.min(end, bytes.length); i++) text += String.fromCharCode(bytes[i]!);
  return text;
}

/** Works out which document format a ZIP container holds. */
function sniffZip(bytes: Uint8Array): FormatInfo {
  const zip = BY_ID.get('zip')!;
  try {
    const archive = openZip(bytes);
    if (archive.has('[Content_Types].xml')) {
      if (archive.has('word/document.xml')) return BY_ID.get('docx')!;
      if (archive.has('xl/workbook.xml')) return BY_ID.get('xlsx')!;
      if (archive.has('ppt/presentation.xml')) return BY_ID.get('pptx')!;
    }
    const mimetype = archive.get('mimetype');
    if (mimetype && mimetype.method === 0 && mimetype.size < 200) {
      // The ODF/EPUB "mimetype" entry is stored uncompressed, so it can be read synchronously.
      const start = mimetype.offset + 30 + 8 + extraLengthAt(bytes, mimetype.offset);
      const declared = ascii(bytes, start, start + mimetype.size).trim();
      const match = getFormatByMimeType(declared);
      if (match) return match;
    }
    if (archive.has('META-INF/container.xml')) return BY_ID.get('epub')!;
  } catch {
    // Not a readable archive: fall through and let the ZIP renderer report it.
  }
  return zip;
}

function extraLengthAt(bytes: Uint8Array, localHeaderOffset: number): number {
  const at = localHeaderOffset + 28;
  return at + 1 < bytes.length ? bytes[at]! | (bytes[at + 1]! << 8) : 0;
}

/** Identifies a format from the file's leading bytes ("magic numbers"). */
export function sniffFormat(bytes: Uint8Array): FormatInfo | null {
  const id = (formatId: string): FormatInfo => BY_ID.get(formatId)!;
  if (startsWith(bytes, [0x25, 0x50, 0x44, 0x46])) return id('pdf');
  if (startsWith(bytes, [0x89, 0x50, 0x4e, 0x47])) return id('image');
  if (startsWith(bytes, [0xff, 0xd8, 0xff])) return id('image');
  if (startsWith(bytes, [0x47, 0x49, 0x46, 0x38])) return id('image');
  if (startsWith(bytes, [0x42, 0x4d]) && bytes.length > 26 && bytes[6] === 0 && bytes[7] === 0) {
    return id('image');
  }
  if (startsWith(bytes, [0x00, 0x00, 0x01, 0x00]) && bytes.length > 6 && bytes[5] === 0) {
    return id('image');
  }
  if (startsWith(bytes, [0x52, 0x49, 0x46, 0x46])) {
    const kind = ascii(bytes, 8, 12);
    if (kind === 'WEBP') return id('image');
    if (kind === 'WAVE') return id('audio');
    if (kind === 'AVI ') return id('video');
  }
  if (ascii(bytes, 4, 8) === 'ftyp') {
    const brand = ascii(bytes, 8, 12);
    if (brand === 'avif' || brand === 'avis') return id('image');
    if (brand.startsWith('hei') || brand === 'mif1' || brand === 'msf1') return id('heic');
    if (brand.startsWith('M4A')) return id('audio');
    return id('video');
  }
  if (startsWith(bytes, [0x1a, 0x45, 0xdf, 0xa3])) return id('video');
  if (startsWith(bytes, [0x4f, 0x67, 0x67, 0x53])) return id('audio');
  if (startsWith(bytes, [0x49, 0x44, 0x33])) return id('audio');
  if (startsWith(bytes, [0xff, 0xfb]) || startsWith(bytes, [0xff, 0xf3])) return id('audio');
  if (startsWith(bytes, [0x66, 0x4c, 0x61, 0x43])) return id('audio');
  if (startsWith(bytes, [0x49, 0x49, 0x2a, 0x00]) || startsWith(bytes, [0x4d, 0x4d, 0x00, 0x2a])) {
    return id('tiff');
  }
  if (startsWith(bytes, [0x77, 0x4f, 0x46, 0x46]) || startsWith(bytes, [0x77, 0x4f, 0x46, 0x32])) {
    return id('font');
  }
  if (
    startsWith(bytes, [0x4f, 0x54, 0x54, 0x4f]) ||
    startsWith(bytes, [0x00, 0x01, 0x00, 0x00, 0x00])
  ) {
    return id('font');
  }
  if (startsWith(bytes, [0x50, 0x4b, 0x03, 0x04]) || startsWith(bytes, [0x50, 0x4b, 0x05, 0x06])) {
    return sniffZip(bytes);
  }
  if (startsWith(bytes, [0x7b, 0x5c, 0x72, 0x74, 0x66])) return id('rtf');
  if (startsWith(bytes, [0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])) {
    // Legacy Office compound file; the exact application cannot be told apart cheaply.
    return id('ole');
  }
  if (!looksLikeText(bytes)) return null;

  const head = new TextDecoder('utf-8').decode(bytes.subarray(0, 1024));
  const trimmed = head.trimStart();
  const lower = trimmed.slice(0, 256).toLowerCase();
  if (lower.startsWith('<!doctype html') || lower.startsWith('<html')) return id('html');
  if (lower.startsWith('<svg') || (lower.startsWith('<?xml') && lower.includes('<svg'))) {
    return id('svg');
  }
  if (lower.startsWith('<?xml')) return id('xml');
  if (/^[[{]/.test(trimmed)) {
    if (/"nbformat"\s*:/.test(head) || /"cells"\s*:\s*\[/.test(head)) return id('ipynb');
    return id('json');
  }
  if (/^(from|received|return-path|mime-version|delivered-to|message-id):/i.test(trimmed)) {
    return id('eml');
  }
  return id('text');
}

export interface DetectInput {
  /** Explicit `type` option. */
  hint?: string | undefined;
  name: string;
  mimeType: string;
  /** Lazily provides the file's bytes; only called when name and MIME type are inconclusive. */
  bytes?: () => Promise<Uint8Array>;
}

/**
 * Detection order: explicit hint → file extension → MIME type → content sniffing.
 * A generic ZIP or JSON answer is refined by sniffing so that, say, a `.zip`-less
 * DOCX served as `application/zip` still opens as a document.
 */
export async function detectFormat(input: DetectInput): Promise<FormatInfo> {
  if (input.hint) {
    const hinted = resolveFormatHint(input.hint);
    if (hinted) return hinted;
  }
  const byExtension = getFormatByExtension(extensionOf(input.name));
  if (byExtension) return byExtension;

  const byMime = getFormatByMimeType(input.mimeType);
  const ambiguous = byMime === null || byMime.id === 'zip' || byMime.id === 'text';
  if (byMime && !ambiguous) return byMime;

  if (input.bytes) {
    const sniffed = sniffFormat(await input.bytes());
    if (sniffed && !(sniffed.id === 'text' && byMime)) return sniffed;
  }
  return byMime ?? UNKNOWN_FORMAT;
}
