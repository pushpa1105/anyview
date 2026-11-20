import { describe, expect, it, vi } from 'vitest';
import { isAnyViewError } from '../../src';
import {
  attr,
  child,
  childElements,
  decodeEntities,
  descendants,
  find,
  isElement,
  parseXml,
  textOf,
} from '../../src/util/xml';
import { inflateRaw, openZip, resolveZipPath } from '../../src/util/zip';
import { makeZip } from '../helpers/zip';

describe('parseXml', () => {
  it('parses elements, attributes and text', () => {
    const root = parseXml('<?xml version="1.0"?><a x="1" y=\'two\'><b>hi</b><c/></a>');
    expect(root.name).toBe('a');
    expect(root.attrs).toEqual({ x: '1', y: 'two' });
    expect(childElements(root).map((element) => element.name)).toEqual(['b', 'c']);
    expect(textOf(child(root, 'b'))).toBe('hi');
  });

  it('splits namespace prefixes into local names', () => {
    const root = parseXml('<w:p xmlns:w="urn:w"><w:r w:val="1" r:id="rId2"/></w:p>');
    expect(root.local).toBe('p');
    const run = child(root, 'r')!;
    expect(run.name).toBe('w:r');
    expect(attr(run, 'w:val')).toBe('1');
    expect(attr(run, 'r:id')).toBe('rId2');
  });

  it('matches prefixed attributes under any prefix but not unprefixed ones', () => {
    const root = parseXml('<a ns0:val="x" id="plain" rel:id="linked"/>');
    expect(attr(root, 'w:val')).toBe('x');
    expect(attr(root, 'id')).toBe('plain');
    expect(attr(root, 'r:id')).toBe('linked');
    expect(attr(root, 'val')).toBeUndefined();
    expect(attr(root, 'missing')).toBeUndefined();
    expect(attr(null, 'x')).toBeUndefined();
  });

  it('decodes entities in text and attributes', () => {
    const root = parseXml('<a t="&lt;&amp;&quot;&#65;&#x42;">&gt;&apos;&nbsp;&unknown;</a>');
    expect(root.attrs['t']).toBe('<&"AB');
    expect(textOf(root)).toBe(">' &unknown;");
    expect(decodeEntities('no entities')).toBe('no entities');
    expect(decodeEntities('&#0;&#x110000;&#99999999999;')).toBe('');
  });

  it('keeps CDATA verbatim and skips comments, PIs and doctypes', () => {
    const root = parseXml(
      '<!DOCTYPE a [<!ENTITY x "y">]><!-- c --><a><?pi data?><![CDATA[<b>&amp;</b>]]><!-- c --></a>',
    );
    expect(textOf(root)).toBe('<b>&amp;</b>');
  });

  it('drops indentation between elements unless asked to preserve it', () => {
    const source = '<a>\n  <b> x </b>\n  <c> </c>\n</a>';
    const compact = parseXml(source);
    expect(compact.children).toHaveLength(2);
    expect(textOf(child(compact, 'b'))).toBe(' x ');
    // Whitespace-only text is significant when it is the only content.
    expect(textOf(child(compact, 'c'))).toBe(' ');
    const preserved = parseXml(source, { preserveWhitespace: true });
    expect(preserved.children.length).toBe(5);
  });

  it('handles ">" inside attribute values and a leading BOM', () => {
    const root = parseXml('﻿<a expr="1 > 0" other=\'a>b\'>ok</a>');
    expect(root.attrs['expr']).toBe('1 > 0');
    expect(root.attrs['other']).toBe('a>b');
  });

  it.each([
    ['<a><b></a>', 'unexpected closing tag'],
    ['<a>', 'never closed'],
    ['', 'no root element'],
    ['<a/><b/>', 'more than one root'],
    ['<a><!-- oops</a>', 'unterminated comment'],
    ['<a><![CDATA[oops</a>', 'unterminated CDATA'],
    ['<a><?pi oops</a>', 'unterminated processing instruction'],
    ['<a attr="x"', 'unterminated tag'],
    ['<a></a', 'unterminated closing tag'],
    ['<!DOCTYPE a', 'unterminated declaration'],
    ['< ></>', 'empty tag name'],
  ])('rejects malformed input %j', (source, message) => {
    expect(() => parseXml(source)).toThrow(message);
  });

  it('provides traversal helpers', () => {
    const root = parseXml('<a><b><c n="1"/><c n="2"/></b><d><c n="3"/></d>text</a>');
    expect(find(root, 'b', 'c')?.attrs['n']).toBe('1');
    expect(find(root, 'b', 'missing')).toBeNull();
    expect(find(null, 'b')).toBeNull();
    expect(descendants(root, 'c').map((element) => element.attrs['n'])).toEqual(['1', '2', '3']);
    expect(descendants(null, 'c')).toEqual([]);
    expect(childElements(null)).toEqual([]);
    expect(childElements(root, 'd')).toHaveLength(1);
    expect(child(null, 'b')).toBeNull();
    expect(textOf(null)).toBe('');
    expect(isElement(root.children[0])).toBe(true);
    expect(isElement(root.children[2])).toBe(false);
  });
});

describe('openZip', () => {
  it('reads stored and deflated entries', async () => {
    const zip = openZip(makeZip({ 'a.txt': 'alpha', 'dir/b.txt': 'beta '.repeat(500) }));
    expect(zip.entries.map((entry) => entry.name)).toEqual(['a.txt', 'dir/b.txt']);
    expect(await zip.readText('a.txt')).toBe('alpha');
    expect(await zip.readText('dir/b.txt')).toBe('beta '.repeat(500));
    const stored = openZip(makeZip({ 'a.txt': 'alpha' }, { compress: false }));
    expect(await stored.readText('a.txt')).toBe('alpha');
    expect(stored.entries[0]!.compressedSize).toBe(5);
  });

  it('looks entries up leniently', async () => {
    const zip = openZip(makeZip({ 'Word/Document.xml': '<x/>' }));
    expect(zip.has('word/document.xml')).toBe(true);
    expect(zip.has('/Word/Document.xml')).toBe(true);
    expect(zip.has('nope')).toBe(false);
    expect(await zip.read('nope')).toBeNull();
    expect(await zip.readText('nope')).toBeNull();
  });

  it('flags directories', () => {
    const zip = openZip(makeZip({ 'folder/': '', 'folder/file.txt': 'x' }, { compress: false }));
    expect(zip.entries[0]!.isDirectory).toBe(true);
    expect(zip.entries[1]!.isDirectory).toBe(false);
  });

  it('finds the directory when the archive has a trailing comment', async () => {
    const base = makeZip({ 'a.txt': 'alpha' });
    const withComment = new Uint8Array(base.length + 10);
    withComment.set(base);
    // Declare a 10-byte archive comment in the end-of-central-directory record.
    new DataView(withComment.buffer).setUint16(base.length - 2, 10, true);
    expect(await openZip(withComment).readText('a.txt')).toBe('alpha');
  });

  const expectCode = (run: () => unknown, code: string): void => {
    try {
      run();
    } catch (error) {
      expect(isAnyViewError(error) && error.code).toBe(code);
      return;
    }
    throw new Error('expected an error');
  };

  it('rejects files that are not ZIP archives', () => {
    expectCode(() => openZip(new Uint8Array(4)), 'PARSE_FAILED');
    expectCode(() => openZip(new Uint8Array(64)), 'PARSE_FAILED');
  });

  it('rejects archives with a corrupt central directory', () => {
    const data = makeZip({ 'a.txt': 'alpha' });
    const view = new DataView(data.buffer);
    const directoryOffset = view.getUint32(data.length - 6, true);
    view.setUint32(directoryOffset, 0xdeadbeef, true);
    expectCode(() => openZip(data), 'PARSE_FAILED');

    const truncated = makeZip({ 'a.txt': 'alpha' });
    new DataView(truncated.buffer).setUint32(truncated.length - 6, 0x7fffffff, true);
    expectCode(() => openZip(truncated), 'PARSE_FAILED');
  });

  it('enforces the entry-count limit', () => {
    const data = makeZip({ a: '1', b: '2', c: '3' });
    expectCode(() => openZip(data, { maxEntries: 2 }), 'FILE_TOO_LARGE');
  });

  it('refuses entries above the size limit, encrypted entries and unknown methods', async () => {
    const big = openZip(makeZip({ 'a.txt': 'x'.repeat(1000) }), { maxEntrySize: 100 });
    await expect(big.read('a.txt')).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' });

    const encrypted = openZip(makeZip({ 'a.txt': 'secret' }, { encrypted: true }));
    await expect(encrypted.read('a.txt')).rejects.toMatchObject({ code: 'PARSE_FAILED' });

    const exotic = openZip(makeZip({ 'a.txt': 'data' }, { compress: false, method: 99 }));
    await expect(exotic.read('a.txt')).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('detects a bad local header and out-of-bounds data', async () => {
    const data = makeZip({ 'a.txt': 'alpha' }, { compress: false });
    new DataView(data.buffer).setUint32(0, 0, true);
    await expect(openZip(data).read('a.txt')).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('reads ZIP64 size fields from the extra data', async () => {
    // Hand-built archive: one stored entry whose sizes live in a ZIP64 extra field.
    const name = new TextEncoder().encode('a.txt');
    const content = new TextEncoder().encode('zip64');
    const local = new Uint8Array(30 + name.length + content.length);
    const localView = new DataView(local.buffer);
    localView.setUint32(0, 0x04034b50, true);
    localView.setUint16(26, name.length, true);
    local.set(name, 30);
    local.set(content, 30 + name.length);

    const extra = new Uint8Array(4 + 16);
    const extraView = new DataView(extra.buffer);
    extraView.setUint16(0, 0x0001, true);
    extraView.setUint16(2, 16, true);
    extraView.setBigUint64(4, BigInt(content.length), true);
    extraView.setBigUint64(12, BigInt(content.length), true);

    const central = new Uint8Array(46 + name.length + extra.length);
    const centralView = new DataView(central.buffer);
    centralView.setUint32(0, 0x02014b50, true);
    centralView.setUint32(20, 0xffffffff, true);
    centralView.setUint32(24, 0xffffffff, true);
    centralView.setUint16(28, name.length, true);
    centralView.setUint16(30, extra.length, true);
    centralView.setUint32(42, 0, true);
    central.set(name, 46);
    central.set(extra, 46 + name.length);

    const end = new Uint8Array(22);
    const endView = new DataView(end.buffer);
    endView.setUint32(0, 0x06054b50, true);
    endView.setUint16(10, 1, true);
    endView.setUint32(12, central.length, true);
    endView.setUint32(16, local.length, true);

    const archive = new Uint8Array(local.length + central.length + end.length);
    archive.set(local);
    archive.set(central, local.length);
    archive.set(end, local.length + central.length);
    const zip = openZip(archive);
    expect(zip.entries[0]!.size).toBe(5);
    expect(await zip.readText('a.txt')).toBe('zip64');
  });

  it('reads the ZIP64 end-of-central-directory record', async () => {
    const base = makeZip({ 'a.txt': 'alpha' }, { compress: false });
    const baseView = new DataView(base.buffer);
    const eocd = base.length - 22;
    const directorySize = baseView.getUint32(eocd + 12, true);
    const directoryOffset = baseView.getUint32(eocd + 16, true);

    // Insert a ZIP64 EOCD record and locator before the classic record, then
    // mark the classic fields as "see ZIP64".
    const record = new Uint8Array(56);
    const recordView = new DataView(record.buffer);
    recordView.setUint32(0, 0x06064b50, true);
    recordView.setBigUint64(32, 1n, true);
    recordView.setBigUint64(40, BigInt(directorySize), true);
    recordView.setBigUint64(48, BigInt(directoryOffset), true);
    const locator = new Uint8Array(20);
    const locatorView = new DataView(locator.buffer);
    locatorView.setUint32(0, 0x07064b50, true);
    locatorView.setBigUint64(8, BigInt(eocd), true);

    const archive = new Uint8Array(base.length + record.length + locator.length);
    archive.set(base.subarray(0, eocd));
    archive.set(record, eocd);
    archive.set(locator, eocd + record.length);
    archive.set(base.subarray(eocd), eocd + record.length + locator.length);
    const view = new DataView(archive.buffer);
    const newEocd = archive.length - 22;
    view.setUint16(newEocd + 10, 0xffff, true);
    view.setUint32(newEocd + 12, 0xffffffff, true);
    view.setUint32(newEocd + 16, 0xffffffff, true);

    expect(await openZip(archive).readText('a.txt')).toBe('alpha');

    // Without the locator the archive is invalid.
    const broken = archive.slice();
    new DataView(broken.buffer).setUint32(eocd + record.length, 0, true);
    expectCode(() => openZip(broken), 'PARSE_FAILED');
  });
});

describe('inflateRaw', () => {
  it('rejects corrupt data', async () => {
    await expect(inflateRaw(new Uint8Array([0xff, 0xff, 0xff, 0xff]))).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('stops when the output exceeds the limit', async () => {
    const data = makeZip({ a: 'x'.repeat(100_000) });
    const zip = openZip(data);
    const entry = zip.entries[0]!;
    const compressed = data.subarray(30 + 1, 30 + 1 + entry.compressedSize);
    await expect(inflateRaw(compressed, 1000)).rejects.toMatchObject({ code: 'FILE_TOO_LARGE' });
  });

  it('reports browsers without DecompressionStream', async () => {
    vi.stubGlobal('DecompressionStream', undefined);
    try {
      await expect(inflateRaw(new Uint8Array(1))).rejects.toMatchObject({ code: 'RENDER_FAILED' });
    } finally {
      vi.unstubAllGlobals();
    }
  });
});

describe('resolveZipPath', () => {
  it('resolves relative, parent and absolute targets', () => {
    expect(resolveZipPath('word/document.xml', 'media/image1.png')).toBe('word/media/image1.png');
    expect(resolveZipPath('ppt/slides/slide1.xml', '../media/a.png')).toBe('ppt/media/a.png');
    expect(resolveZipPath('ppt/slides/slide1.xml', '/ppt/theme/theme1.xml')).toBe(
      'ppt/theme/theme1.xml',
    );
    expect(resolveZipPath('', 'word/document.xml')).toBe('word/document.xml');
    expect(resolveZipPath('a/b.xml', './c.xml')).toBe('a/c.xml');
  });
});
