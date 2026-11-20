import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  FORMATS,
  SUPPORTED_FORMAT_IDS,
  detectFormat,
  getFormatByExtension,
  getFormatById,
  getFormatByMimeType,
  isAnyViewError,
  sniffFormat,
} from '../../src';
import { languageFor, resolveFormatHint } from '../../src/core/formats';
import { findRenderer, matchCustomRenderer } from '../../src/core/registry';
import { createFile } from '../../src/core/source';
import type { Renderer, RendererDefinition } from '../../src';
import { mockFetch } from '../helpers/dom';
import { makeZip, PNG_1X1 } from '../helpers/zip';

const ascii = (text: string): Uint8Array => new TextEncoder().encode(text);
const bytes = (...values: number[]): Uint8Array => new Uint8Array(values);
const padded = (values: number[], length = 64): Uint8Array => {
  const data = new Uint8Array(length);
  data.set(values);
  return data;
};

describe('format table', () => {
  it('has unique ids and extensions', () => {
    const ids = FORMATS.map((format) => format.id);
    expect(new Set(ids).size).toBe(ids.length);
    const extensions = FORMATS.flatMap((format) => format.extensions);
    expect(new Set(extensions).size).toBe(extensions.length);
  });

  it('only advertises renderers for known formats', () => {
    for (const id of SUPPORTED_FORMAT_IDS) expect(getFormatById(id)).not.toBeNull();
    expect(SUPPORTED_FORMAT_IDS).toContain('docx');
    expect(SUPPORTED_FORMAT_IDS).not.toContain('doc');
  });

  it('looks formats up by id, extension and MIME type', () => {
    expect(getFormatById('pdf')?.label).toBe('PDF document');
    expect(getFormatById('nope')).toBeNull();
    expect(getFormatByExtension('.DOCX')?.id).toBe('docx');
    expect(getFormatByExtension('xyz')).toBeNull();
    expect(getFormatByMimeType('application/pdf; charset=binary')?.id).toBe('pdf');
    expect(getFormatByMimeType('TEXT/HTML')?.id).toBe('html');
    expect(getFormatByMimeType('application/vnd.api+json')?.id).toBe('json');
    expect(getFormatByMimeType('image/svg+xml')?.id).toBe('svg');
    expect(getFormatByMimeType('application/soap+xml')?.id).toBe('xml');
    expect(getFormatByMimeType('video/x-matroska')?.id).toBe('video');
    expect(getFormatByMimeType('audio/x-custom')?.id).toBe('audio');
    expect(getFormatByMimeType('text/x-anything')?.id).toBe('text');
    expect(getFormatByMimeType('application/octet-stream')).toBeNull();
    expect(getFormatByMimeType('application/x-unknown')).toBeNull();
    expect(getFormatByMimeType('')).toBeNull();
  });

  it('resolves user hints', () => {
    expect(resolveFormatHint('docx')?.id).toBe('docx');
    expect(resolveFormatHint('.XLSX')?.id).toBe('xlsx');
    expect(resolveFormatHint('jpg')?.id).toBe('image');
    expect(resolveFormatHint('application/pdf')?.id).toBe('pdf');
    expect(resolveFormatHint('  ')).toBeNull();
    expect(resolveFormatHint('nonsense')).toBeNull();
  });

  it('maps files to highlighter languages', () => {
    expect(languageFor('app.ts', 'text')).toBe('typescript');
    expect(languageFor('data.json', 'json')).toBe('json');
    expect(languageFor('feed.rss', 'xml')).toBe('xml');
    expect(languageFor('index.html', 'html')).toBe('html');
    expect(languageFor('logo.svg', 'svg')).toBe('xml');
    expect(languageFor('Dockerfile', 'text')).toBe('dockerfile');
    expect(languageFor('notes.txt', 'text')).toBeNull();
    expect(languageFor('blob', 'text', '.py')).toBe('python');
    expect(languageFor('', 'text')).toBeNull();
  });
});

describe('sniffFormat', () => {
  it.each([
    ['pdf', ascii('%PDF-1.7\n')],
    ['image', padded([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a])],
    ['image', padded([0xff, 0xd8, 0xff, 0xe0])],
    ['image', ascii('GIF89a........')],
    ['image', padded([0x42, 0x4d, 0x10, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00, 0x00])],
    ['image', padded([0x00, 0x00, 0x01, 0x00, 0x01, 0x00])],
    ['image', ascii('RIFF\u0001\u0001\u0001\u0001WEBPVP8 ')],
    ['audio', ascii('RIFF\u0001\u0001\u0001\u0001WAVEfmt ')],
    ['video', ascii('RIFF\u0001\u0001\u0001\u0001AVI LIST')],
    ['image', ascii('\u0001\u0001\u0001\u0001ftypavif....')],
    ['heic', ascii('\u0001\u0001\u0001\u0001ftypheic....')],
    ['audio', ascii('\u0001\u0001\u0001\u0001ftypM4A ....')],
    ['video', ascii('\u0001\u0001\u0001\u0001ftypisom....')],
    ['video', padded([0x1a, 0x45, 0xdf, 0xa3])],
    ['audio', ascii('OggS........')],
    ['audio', ascii('ID3\u0003........')],
    ['audio', padded([0xff, 0xfb, 0x90])],
    ['audio', ascii('fLaC........')],
    ['tiff', padded([0x49, 0x49, 0x2a, 0x00])],
    ['tiff', padded([0x4d, 0x4d, 0x00, 0x2a])],
    ['font', ascii('wOFF........')],
    ['font', ascii('wOF2........')],
    ['font', ascii('OTTO........')],
    ['font', padded([0x00, 0x01, 0x00, 0x00, 0x00, 0x0c])],
    ['rtf', ascii('{\\rtf1\\ansi hello}')],
    ['ole', padded([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1])],
    ['html', ascii('  <!DOCTYPE html><html></html>')],
    ['html', ascii('<html lang="en">')],
    ['svg', ascii('<svg xmlns="http://www.w3.org/2000/svg"/>')],
    ['svg', ascii('<?xml version="1.0"?>\n<svg/>')],
    ['xml', ascii('<?xml version="1.0"?><root/>')],
    ['json', ascii('{"a": 1}')],
    ['json', ascii('[1, 2]')],
    ['ipynb', ascii('{"cells": [], "nbformat": 4}')],
    ['eml', ascii('From: a@b.c\r\nSubject: hi\r\n\r\nbody')],
    ['eml', ascii('Received: from x\r\n')],
    ['text', ascii('just some words')],
    ['text', bytes()],
  ])('recognises %s', (expected, data) => {
    expect(sniffFormat(data)?.id).toBe(expected);
  });

  it('returns null for unknown binary data', () => {
    expect(sniffFormat(padded([0x01, 0x02, 0x00, 0x03]))).toBeNull();
  });

  it('tells ZIP-based document formats apart', () => {
    const office = (main: string): Uint8Array =>
      makeZip({ '[Content_Types].xml': '<Types/>', [main]: '<x/>' });
    expect(sniffFormat(office('word/document.xml'))?.id).toBe('docx');
    expect(sniffFormat(office('xl/workbook.xml'))?.id).toBe('xlsx');
    expect(sniffFormat(office('ppt/presentation.xml'))?.id).toBe('pptx');
    expect(sniffFormat(office('other/thing.xml'))?.id).toBe('zip');

    const odf = (mimetype: string): Uint8Array =>
      makeZip({ mimetype, 'content.xml': '<x/>' }, { stored: ['mimetype'] });
    expect(sniffFormat(odf('application/vnd.oasis.opendocument.text'))?.id).toBe('odt');
    expect(sniffFormat(odf('application/vnd.oasis.opendocument.spreadsheet'))?.id).toBe('ods');
    expect(sniffFormat(odf('application/vnd.oasis.opendocument.presentation'))?.id).toBe('odp');
    expect(sniffFormat(odf('application/epub+zip'))?.id).toBe('epub');

    expect(sniffFormat(makeZip({ 'META-INF/container.xml': '<c/>' }))?.id).toBe('epub');
    expect(sniffFormat(makeZip({ 'readme.txt': 'hi' }))?.id).toBe('zip');
    // A corrupt archive is still reported as ZIP so the renderer can explain the problem.
    expect(sniffFormat(padded([0x50, 0x4b, 0x03, 0x04]))?.id).toBe('zip');
    expect(sniffFormat(padded([0x50, 0x4b, 0x05, 0x06], 22))?.id).toBe('zip');
  });
});

describe('detectFormat', () => {
  const never = (): Promise<Uint8Array> => Promise.reject(new Error('bytes should not be read'));

  it('prefers the explicit hint', async () => {
    const format = await detectFormat({
      hint: 'csv',
      name: 'a.txt',
      mimeType: 'text/plain',
      bytes: never,
    });
    expect(format.id).toBe('csv');
  });

  it('falls through an unknown hint to the extension', async () => {
    expect((await detectFormat({ hint: '???', name: 'a.md', mimeType: '', bytes: never })).id).toBe(
      'markdown',
    );
  });

  it('uses the extension before the MIME type', async () => {
    const format = await detectFormat({
      name: 'report.docx',
      mimeType: 'application/octet-stream',
      bytes: never,
    });
    expect(format.id).toBe('docx');
  });

  it('uses a specific MIME type when there is no extension', async () => {
    expect(
      (await detectFormat({ name: 'download', mimeType: 'application/pdf', bytes: never })).id,
    ).toBe('pdf');
  });

  it('refines generic ZIP and text MIME types by sniffing', async () => {
    const docx = makeZip({ '[Content_Types].xml': '<Types/>', 'word/document.xml': '<x/>' });
    expect(
      (
        await detectFormat({
          name: 'x',
          mimeType: 'application/zip',
          bytes: () => Promise.resolve(docx),
        })
      ).id,
    ).toBe('docx');
    expect(
      (
        await detectFormat({
          name: 'x',
          mimeType: 'text/plain',
          bytes: () => Promise.resolve(ascii('{"a":1}')),
        })
      ).id,
    ).toBe('json');
    expect(
      (
        await detectFormat({
          name: 'x',
          mimeType: 'text/plain',
          bytes: () => Promise.resolve(ascii('hello')),
        })
      ).id,
    ).toBe('text');
  });

  it('sniffs when nothing else is known', async () => {
    expect(
      (await detectFormat({ name: '', mimeType: '', bytes: () => Promise.resolve(PNG_1X1) })).id,
    ).toBe('image');
  });

  it('returns the unknown format as a last resort', async () => {
    expect((await detectFormat({ name: 'mystery', mimeType: '' })).id).toBe('unknown');
    expect(
      (
        await detectFormat({
          name: 'mystery',
          mimeType: '',
          bytes: () => Promise.resolve(padded([1, 0, 2, 0])),
        })
      ).id,
    ).toBe('unknown');
  });
});

describe('renderer registry', () => {
  const renderer: Renderer = { render: () => undefined };
  const define = (
    definition: Partial<RendererDefinition> & { id: string },
  ): RendererDefinition => ({
    load: () => renderer,
    ...definition,
  });

  it('finds built-in renderers by format id', () => {
    expect(findRenderer('docx', [])?.id).toBe('docx');
    expect(findRenderer('svg', [])?.id).toBe('image');
    expect(findRenderer('doc', [])).toBeNull();
  });

  it('lets custom renderers override built-ins', () => {
    const byFormat = define({ id: 'my-pdf', formats: ['pdf'] });
    expect(findRenderer('pdf', [byFormat])).toBe(byFormat);
    // Reusing a built-in id inherits that renderer's formats.
    const byId = define({ id: 'text' });
    expect(findRenderer('json', [byId])).toBe(byId);
    expect(findRenderer('docx', [byId])?.id).toBe('docx');
  });

  it('routes new extensions and MIME types to custom renderers', () => {
    const custom = define({ id: 'dwg', extensions: ['.DWG'], mimeTypes: ['image/vnd.dwg'] });
    const byExtension = matchCustomRenderer([custom], 'plan.dwg', '');
    expect(byExtension?.definition).toBe(custom);
    expect(byExtension?.format).toMatchObject({ id: 'dwg', label: 'DWG', category: 'other' });
    expect(matchCustomRenderer([custom], 'file', 'image/vnd.dwg')?.definition).toBe(custom);
    expect(matchCustomRenderer([custom], 'plan.pdf', 'application/pdf')).toBeNull();
    expect(matchCustomRenderer([], 'plan.dwg', '')).toBeNull();
  });

  it('reuses a known format when a custom renderer names one', () => {
    const custom = define({ id: 'log', formats: ['text'], extensions: ['applog'] });
    expect(matchCustomRenderer([custom], 'a.applog', '')?.format.id).toBe('text');
  });
});

describe('createFile', () => {
  const options = (overrides: Partial<Parameters<typeof createFile>[1]> = {}) => ({
    maxFileSize: 1024 * 1024,
    signal: new AbortController().signal,
    ...overrides,
  });

  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it('wraps File, Blob, ArrayBuffer and typed arrays', async () => {
    const file = createFile(
      new File(['hello'], 'a.txt', { type: 'Text/Plain; charset=utf-8' }),
      options(),
    );
    expect(file).toMatchObject({
      name: 'a.txt',
      mimeType: 'text/plain',
      size: 5,
      url: null,
      direct: false,
    });
    expect(await file.text()).toBe('hello');
    expect((await file.blob()).size).toBe(5);

    expect(createFile(new Blob(['x']), options()).name).toBe('');
    expect(createFile(new Blob(['x']), options({ name: 'given.csv' })).name).toBe('given.csv');

    const buffer = new Uint8Array([104, 105]).buffer;
    expect(await createFile(buffer, options()).text()).toBe('hi');

    const view = new Uint8Array([0, 104, 105, 0]).subarray(1, 3);
    const fromView = createFile(view, options());
    expect(fromView.size).toBe(2);
    expect(Array.from(await fromView.bytes())).toEqual([104, 105]);
    expect(await createFile(new DataView(buffer), options()).text()).toBe('hi');
  });

  it('caches bytes and text', async () => {
    const file = createFile(new Blob(['abc']), options());
    expect(await file.bytes()).toBe(await file.bytes());
    expect(file.text()).toBe(file.text());
  });

  it('rejects unsupported and empty sources', () => {
    const codeOf = (source: unknown): string | false => {
      try {
        createFile(source as never, options());
      } catch (error) {
        return isAnyViewError(error) && error.code;
      }
      return false;
    };
    expect(codeOf(42)).toBe('INVALID_SOURCE');
    expect(codeOf(null)).toBe('INVALID_SOURCE');
    expect(codeOf({})).toBe('INVALID_SOURCE');
    expect(codeOf('')).toBe('INVALID_SOURCE');
    expect(codeOf('   ')).toBe('INVALID_SOURCE');
    expect(codeOf('http://[bad')).toBe('INVALID_SOURCE');
  });

  it('enforces the size limit for local data', async () => {
    const file = createFile(new Blob(['x'.repeat(100)]), options({ maxFileSize: 10 }));
    await expect(file.bytes()).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' });
    await expect(file.blob()).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' });
  });

  it('describes URL sources without fetching them', () => {
    const fetchMock = mockFetch(() => new Response('x'));
    const file = createFile('https://example.com/files/My%20Doc.pdf?sig=1', options());
    expect(file).toMatchObject({
      name: 'My Doc.pdf',
      mimeType: '',
      size: null,
      url: 'https://example.com/files/My%20Doc.pdf?sig=1',
      direct: true,
    });
    expect(createFile(new URL('https://example.com/a.png'), options()).name).toBe('a.png');
    expect(createFile('relative/file.csv', options()).url).toMatch(/relative\/file\.csv$/);
    expect(fetchMock).not.toHaveBeenCalled();
  });

  it('downloads on demand, once, and adopts response metadata', async () => {
    const fetchMock = mockFetch(
      () =>
        new Response('a,b', {
          headers: {
            'content-type': 'text/csv; charset=utf-8',
            'content-disposition': 'attachment; filename="export.csv"',
          },
        }),
    );
    const file = createFile('https://example.com/download?id=7', options());
    expect(file.name).toBe('download');
    expect(await file.text()).toBe('a,b');
    await file.bytes();
    await file.blob();
    expect(fetchMock).toHaveBeenCalledTimes(1);
    expect(file).toMatchObject({ name: 'export.csv', mimeType: 'text/csv', size: 3 });
  });

  it('decodes RFC 5987 file names and keeps an explicit name', async () => {
    mockFetch(
      () =>
        new Response('x', {
          headers: { 'content-disposition': "attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.pdf" },
        }),
    );
    const file = createFile('https://example.com/d', options());
    await file.blob();
    expect(file.name).toBe('résumé.pdf');

    const named = createFile('https://example.com/d', options({ name: 'mine.pdf' }));
    await named.blob();
    expect(named.name).toBe('mine.pdf');
  });

  it('passes fetch options through and marks such files as not directly usable', async () => {
    const fetchMock = mockFetch(() => new Response('x'));
    const headers = { Authorization: 'Bearer token' };
    const file = createFile('https://example.com/a.txt', options({ fetchOptions: { headers } }));
    expect(file.direct).toBe(false);
    await file.bytes();
    expect(fetchMock.mock.calls[0]![1]).toMatchObject({ headers });
  });

  it('reports HTTP errors, network errors and oversized downloads', async () => {
    mockFetch(() => new Response('nope', { status: 404 }));
    await expect(createFile('https://example.com/a', options()).bytes()).rejects.toMatchObject({
      code: 'FETCH_FAILED',
      details: { status: 404 },
    });

    mockFetch(() => {
      throw new TypeError('Failed to fetch');
    });
    await expect(createFile('https://example.com/a', options()).bytes()).rejects.toMatchObject({
      code: 'FETCH_FAILED',
    });

    mockFetch(() => new Response('x', { headers: { 'content-length': '99999999' } }));
    await expect(createFile('https://example.com/a', options()).bytes()).rejects.toMatchObject({
      code: 'FILE_TOO_LARGE',
    });

    mockFetch(() => new Response('x'.repeat(50)));
    await expect(
      createFile('https://example.com/a', options({ maxFileSize: 10 })).bytes(),
    ).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' });
  });

  it('maps aborted downloads to ABORTED', async () => {
    mockFetch(() => {
      throw new DOMException('aborted', 'AbortError');
    });
    await expect(createFile('https://example.com/a', options()).bytes()).rejects.toMatchObject({
      code: 'ABORTED',
    });
  });

  it('reads the MIME type of data URLs without fetching', () => {
    const file = createFile('data:text/plain;base64,aGk=', options());
    expect(file).toMatchObject({ name: '', mimeType: 'text/plain', direct: false });
  });
});
