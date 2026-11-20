import { describe, expect, it } from 'vitest';
import { formatLabel } from '../../src/core/labels';
import {
  base64ToBytes,
  decodeText,
  decodeWith,
  formatBytes,
  looksLikeText,
  toBlobPart,
} from '../../src/util/text';
import { extensionOf, fileNameFromUrl, safeImageUrl, safeLinkUrl } from '../../src/util/url';

const bytes = (...values: number[]): Uint8Array => new Uint8Array(values);

describe('decodeText', () => {
  it('decodes plain UTF-8', () => {
    expect(decodeText(new TextEncoder().encode('héllo ✓'))).toBe('héllo ✓');
  });

  it('strips a UTF-8 byte-order mark', () => {
    expect(decodeText(bytes(0xef, 0xbb, 0xbf, 0x68, 0x69))).toBe('hi');
  });

  it('decodes UTF-16 little and big endian by BOM', () => {
    expect(decodeText(bytes(0xff, 0xfe, 0x68, 0x00, 0x69, 0x00))).toBe('hi');
    expect(decodeText(bytes(0xfe, 0xff, 0x00, 0x68, 0x00, 0x69))).toBe('hi');
  });

  it('falls back to Windows-1252 for invalid UTF-8', () => {
    expect(decodeText(bytes(0x63, 0x61, 0x66, 0xe9))).toBe('café');
  });

  it('decodeWith tolerates unknown charset labels', () => {
    expect(decodeWith(bytes(0xe9), 'not-a-charset')).toBe('é');
    expect(decodeWith(bytes(0xa4), 'iso-8859-15')).toBe('€');
  });
});

describe('looksLikeText', () => {
  it('accepts empty input, ASCII and UTF-16 with BOM', () => {
    expect(looksLikeText(bytes())).toBe(true);
    expect(looksLikeText(new TextEncoder().encode('hello\n\tworld'))).toBe(true);
    expect(looksLikeText(bytes(0xff, 0xfe, 0x68, 0x00))).toBe(true);
  });

  it('rejects NUL bytes and dense control characters', () => {
    expect(looksLikeText(bytes(0x68, 0x00, 0x69))).toBe(false);
    expect(looksLikeText(bytes(1, 2, 3, 4, 5, 6, 7, 8, 0x41))).toBe(false);
  });
});

describe('formatBytes', () => {
  it.each([
    [0, '0 B'],
    [1023, '1023 B'],
    [1024, '1.0 KB'],
    [1536, '1.5 KB'],
    [150 * 1024, '150 KB'],
    [5 * 1024 * 1024, '5.0 MB'],
    [3 * 1024 ** 3, '3.0 GB'],
  ])('formats %d as %s', (input, expected) => {
    expect(formatBytes(input)).toBe(expected);
  });

  it('returns an empty string for invalid sizes', () => {
    expect(formatBytes(-1)).toBe('');
    expect(formatBytes(Number.NaN)).toBe('');
  });
});

describe('base64ToBytes', () => {
  it('decodes, ignoring whitespace and padding', () => {
    expect(Array.from(base64ToBytes('aGVs\r\nbG8='))).toEqual([104, 101, 108, 108, 111]);
  });

  it('survives truncated input', () => {
    expect(base64ToBytes('aGVsb').length).toBeGreaterThan(0);
    expect(base64ToBytes('').length).toBe(0);
  });
});

describe('toBlobPart', () => {
  it('copies only the viewed range', () => {
    const source = new Uint8Array([1, 2, 3, 4, 5]).subarray(1, 4);
    expect(Array.from(new Uint8Array(toBlobPart(source)))).toEqual([2, 3, 4]);
  });
});

describe('safeLinkUrl', () => {
  it('allows http, https, mailto, tel and fragments', () => {
    expect(safeLinkUrl('https://example.com/a?b=1')).toBe('https://example.com/a?b=1');
    expect(safeLinkUrl('HTTP://EXAMPLE.COM')).toBe('HTTP://EXAMPLE.COM');
    expect(safeLinkUrl('mailto:a@b.c')).toBe('mailto:a@b.c');
    expect(safeLinkUrl('tel:+123')).toBe('tel:+123');
    expect(safeLinkUrl('#section')).toBe('#section');
  });

  it.each([
    'javascript:alert(1)',
    'JaVaScRiPt:alert(1)',
    '  javascript:alert(1)',
    'java\nscript:alert(1)',
    'java\tscript:alert(1)',
    '\u0001javascript:alert(1)',
    'data:text/html,<script>alert(1)</script>',
    'vbscript:msgbox(1)',
    'file:///etc/passwd',
    'blob:https://example.com/x',
  ])('rejects %j', (url) => {
    expect(safeLinkUrl(url)).toBeNull();
  });

  it('resolves relative URLs against a base and rejects unsafe bases', () => {
    expect(safeLinkUrl('page.html', 'https://example.com/docs/readme.md')).toBe(
      'https://example.com/docs/page.html',
    );
    expect(safeLinkUrl('page.html', 'javascript:void(0)')).toBeNull();
    expect(safeLinkUrl('page.html')).toBe('page.html');
  });

  it('returns null for empty input', () => {
    expect(safeLinkUrl('')).toBeNull();
    expect(safeLinkUrl('   ')).toBeNull();
    expect(safeLinkUrl(null)).toBeNull();
    expect(safeLinkUrl(undefined)).toBeNull();
  });
});

describe('safeImageUrl', () => {
  it('allows http(s), blob and base64 raster data URLs', () => {
    expect(safeImageUrl('https://example.com/a.png')).toBe('https://example.com/a.png');
    expect(safeImageUrl('blob:https://example.com/id')).toBe('blob:https://example.com/id');
    expect(safeImageUrl('data:image/png;base64,AAAA')).toBe('data:image/png;base64,AAAA');
  });

  it('rejects scriptable or unknown sources', () => {
    expect(safeImageUrl('javascript:alert(1)')).toBeNull();
    expect(safeImageUrl('data:image/svg+xml;base64,AAAA')).toBeNull();
    expect(safeImageUrl('data:text/html;base64,AAAA')).toBeNull();
    expect(safeImageUrl('file:///x.png')).toBeNull();
    expect(safeImageUrl('')).toBeNull();
    expect(safeImageUrl(' ')).toBeNull();
  });

  it('resolves relative paths against a base', () => {
    expect(safeImageUrl('img/a.png', 'https://example.com/docs/readme.md')).toBe(
      'https://example.com/docs/img/a.png',
    );
    expect(safeImageUrl('a.png', 'file:///docs/readme.md')).toBeNull();
    expect(safeImageUrl('a.png')).toBe('a.png');
  });
});

describe('fileNameFromUrl / extensionOf', () => {
  it('extracts and decodes the last path segment', () => {
    expect(fileNameFromUrl('https://example.com/files/My%20Report.pdf?token=1#p2')).toBe(
      'My Report.pdf',
    );
    expect(fileNameFromUrl('https://example.com/')).toBe('');
    expect(fileNameFromUrl('/a/b/c.txt')).toBe('c.txt');
    expect(fileNameFromUrl('https://example.com/bad%E0%A4%A.txt')).toBe('bad%E0%A4%A.txt');
  });

  it('returns lower-case extensions', () => {
    expect(extensionOf('Report.PDF')).toBe('pdf');
    expect(extensionOf('archive.tar.gz')).toBe('gz');
    expect(extensionOf('C:\\docs\\a.docx')).toBe('docx');
    expect(extensionOf('noextension')).toBe('');
    expect(extensionOf('.gitignore')).toBe('');
    expect(extensionOf('trailing.')).toBe('');
  });
});

describe('formatLabel', () => {
  it('replaces known placeholders and keeps unknown ones', () => {
    expect(formatLabel('{page} / {count} {other}', { page: 2, count: 9 })).toBe('2 / 9 {other}');
  });
});
