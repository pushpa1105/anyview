import { AnyViewError } from '../core/errors';
import { decodeText, toBlobPart } from './text';

/**
 * A minimal ZIP reader. Decompression uses the browser's native
 * `DecompressionStream('deflate-raw')`; only stored and deflated entries are
 * supported, which covers every office/EPUB container.
 */

export interface ZipEntry {
  readonly name: string;
  readonly size: number;
  readonly compressedSize: number;
  readonly isDirectory: boolean;
  /** @internal */ readonly method: number;
  /** @internal */ readonly offset: number;
  /** @internal */ readonly encrypted: boolean;
}

export interface ZipLimits {
  /** Maximum number of entries. Default 100 000. */
  maxEntries?: number;
  /** Maximum uncompressed size of one entry in bytes. Default 256 MB. */
  maxEntrySize?: number;
}

const EOCD_SIGNATURE = 0x06054b50;
const ZIP64_LOCATOR_SIGNATURE = 0x07064b50;
const ZIP64_EOCD_SIGNATURE = 0x06064b50;
const CENTRAL_SIGNATURE = 0x02014b50;
const LOCAL_SIGNATURE = 0x04034b50;

function invalid(message: string): AnyViewError {
  return new AnyViewError('PARSE_FAILED', `Invalid ZIP archive: ${message}`);
}

function readUint64(view: DataView, offset: number): number {
  const value = view.getBigUint64(offset, true);
  if (value > BigInt(Number.MAX_SAFE_INTEGER)) throw invalid('value out of range');
  return Number(value);
}

export class ZipArchive {
  readonly entries: readonly ZipEntry[];
  private readonly bytes: Uint8Array;
  private readonly byName = new Map<string, ZipEntry>();
  private readonly byLowerName = new Map<string, ZipEntry>();
  private readonly maxEntrySize: number;

  /** @internal Use {@link openZip}. */
  constructor(bytes: Uint8Array, entries: ZipEntry[], maxEntrySize: number) {
    this.bytes = bytes;
    this.entries = entries;
    this.maxEntrySize = maxEntrySize;
    for (const entry of entries) {
      this.byName.set(entry.name, entry);
      this.byLowerName.set(entry.name.toLowerCase(), entry);
    }
  }

  /** Looks an entry up by path; a leading slash and letter case are forgiven. */
  get(name: string): ZipEntry | null {
    const path = name.replace(/^\/+/, '');
    return this.byName.get(path) ?? this.byLowerName.get(path.toLowerCase()) ?? null;
  }

  has(name: string): boolean {
    return this.get(name) !== null;
  }

  /** Reads and decompresses an entry. Returns `null` when it does not exist. */
  async read(name: string): Promise<Uint8Array | null> {
    const entry = this.get(name);
    return entry ? this.readEntry(entry) : null;
  }

  /** Reads an entry as text. Returns `null` when it does not exist. */
  async readText(name: string): Promise<string | null> {
    const data = await this.read(name);
    return data ? decodeText(data) : null;
  }

  async readEntry(entry: ZipEntry): Promise<Uint8Array> {
    if (entry.encrypted) throw invalid(`"${entry.name}" is encrypted`);
    if (entry.size > this.maxEntrySize) {
      throw new AnyViewError(
        'FILE_TOO_LARGE',
        `"${entry.name}" is too large to extract (${entry.size} bytes).`,
        { details: { size: entry.size, limit: this.maxEntrySize } },
      );
    }
    const { bytes } = this;
    if (entry.offset + 30 > bytes.length) throw invalid('entry is out of bounds');
    const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
    if (view.getUint32(entry.offset, true) !== LOCAL_SIGNATURE) {
      throw invalid(`bad local header for "${entry.name}"`);
    }
    const nameLength = view.getUint16(entry.offset + 26, true);
    const extraLength = view.getUint16(entry.offset + 28, true);
    const start = entry.offset + 30 + nameLength + extraLength;
    const end = start + entry.compressedSize;
    if (end > bytes.length) throw invalid('entry is out of bounds');
    const data = bytes.subarray(start, end);
    if (entry.method === 0) return data;
    if (entry.method === 8) return inflateRaw(data, this.maxEntrySize);
    throw invalid(`unsupported compression method ${entry.method} for "${entry.name}"`);
  }
}

/** Inflates raw DEFLATE data, refusing to produce more than `limit` bytes. */
export async function inflateRaw(data: Uint8Array, limit = Infinity): Promise<Uint8Array> {
  if (typeof DecompressionStream === 'undefined') {
    throw new AnyViewError(
      'RENDER_FAILED',
      'This browser cannot decompress files (DecompressionStream is not available).',
    );
  }
  const inflater = new DecompressionStream('deflate-raw');
  const writer = inflater.writable.getWriter();
  // Write errors surface on the readable side, where they are handled below.
  writer.write(toBlobPart(data)).catch(() => undefined);
  writer.close().catch(() => undefined);
  const reader = inflater.readable.getReader();
  const chunks: Uint8Array[] = [];
  let total = 0;
  try {
    for (;;) {
      const { done, value } = await reader.read();
      if (done) break;
      const chunk = value as Uint8Array;
      total += chunk.length;
      if (total > limit) {
        throw new AnyViewError('FILE_TOO_LARGE', 'Compressed data expands beyond the size limit.', {
          details: { limit },
        });
      }
      chunks.push(chunk);
    }
  } catch (error) {
    await reader.cancel().catch(() => undefined);
    if (error instanceof AnyViewError) throw error;
    throw new AnyViewError('PARSE_FAILED', 'Invalid ZIP archive: corrupt compressed data', {
      cause: error,
    });
  }
  const result = new Uint8Array(total);
  let offset = 0;
  for (const chunk of chunks) {
    result.set(chunk, offset);
    offset += chunk.length;
  }
  return result;
}

function decodeName(bytes: Uint8Array, utf8: boolean): string {
  if (utf8) return new TextDecoder('utf-8').decode(bytes);
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/** Parses the central directory of a ZIP archive. Nothing is decompressed yet. */
export function openZip(bytes: Uint8Array, limits: ZipLimits = {}): ZipArchive {
  const maxEntries = limits.maxEntries ?? 100_000;
  const maxEntrySize = limits.maxEntrySize ?? 256 * 1024 * 1024;
  const view = new DataView(bytes.buffer, bytes.byteOffset, bytes.byteLength);
  const length = bytes.length;
  if (length < 22) throw invalid('file is too small');

  // The end-of-central-directory record sits at the end, before an optional comment.
  let eocd = -1;
  const lowest = Math.max(0, length - 22 - 0xffff);
  for (let i = length - 22; i >= lowest; i--) {
    if (view.getUint32(i, true) === EOCD_SIGNATURE) {
      eocd = i;
      break;
    }
  }
  if (eocd === -1) throw invalid('end of central directory not found');

  let count = view.getUint16(eocd + 10, true);
  let directorySize = view.getUint32(eocd + 12, true);
  let directoryOffset = view.getUint32(eocd + 16, true);

  if (count === 0xffff || directorySize === 0xffffffff || directoryOffset === 0xffffffff) {
    const locator = eocd - 20;
    if (locator < 0 || view.getUint32(locator, true) !== ZIP64_LOCATOR_SIGNATURE) {
      throw invalid('ZIP64 locator not found');
    }
    const eocd64 = readUint64(view, locator + 8);
    if (eocd64 + 56 > length || view.getUint32(eocd64, true) !== ZIP64_EOCD_SIGNATURE) {
      throw invalid('ZIP64 end of central directory not found');
    }
    count = readUint64(view, eocd64 + 32);
    directorySize = readUint64(view, eocd64 + 40);
    directoryOffset = readUint64(view, eocd64 + 48);
  }

  if (count > maxEntries) {
    throw new AnyViewError('FILE_TOO_LARGE', `Archive has too many entries (${count}).`, {
      details: { entries: count, limit: maxEntries },
    });
  }
  if (directoryOffset + directorySize > length) throw invalid('central directory is out of bounds');

  const entries: ZipEntry[] = [];
  let cursor = directoryOffset;
  for (let i = 0; i < count; i++) {
    if (cursor + 46 > length || view.getUint32(cursor, true) !== CENTRAL_SIGNATURE) {
      throw invalid('corrupt central directory');
    }
    const flags = view.getUint16(cursor + 8, true);
    const method = view.getUint16(cursor + 10, true);
    let compressedSize = view.getUint32(cursor + 20, true);
    let size = view.getUint32(cursor + 24, true);
    const nameLength = view.getUint16(cursor + 28, true);
    const extraLength = view.getUint16(cursor + 30, true);
    const commentLength = view.getUint16(cursor + 32, true);
    let offset = view.getUint32(cursor + 42, true);
    const nameStart = cursor + 46;
    const extraStart = nameStart + nameLength;
    if (extraStart + extraLength + commentLength > length)
      throw invalid('corrupt central directory');
    const name = decodeName(bytes.subarray(nameStart, extraStart), (flags & 0x800) !== 0);

    if (size === 0xffffffff || compressedSize === 0xffffffff || offset === 0xffffffff) {
      // ZIP64 extended information: only the overflowed fields are present, in order.
      let extra = extraStart;
      const extraEnd = extraStart + extraLength;
      while (extra + 4 <= extraEnd) {
        const id = view.getUint16(extra, true);
        const fieldSize = view.getUint16(extra + 2, true);
        if (id === 0x0001) {
          let field = extra + 4;
          if (size === 0xffffffff) {
            size = readUint64(view, field);
            field += 8;
          }
          if (compressedSize === 0xffffffff) {
            compressedSize = readUint64(view, field);
            field += 8;
          }
          if (offset === 0xffffffff) offset = readUint64(view, field);
          break;
        }
        extra += 4 + fieldSize;
      }
    }

    entries.push({
      name,
      size,
      compressedSize,
      method,
      offset,
      isDirectory: name.endsWith('/'),
      encrypted: (flags & 0x1) !== 0,
    });
    cursor = extraStart + extraLength + commentLength;
  }
  return new ZipArchive(bytes, entries, maxEntrySize);
}

/** Resolves a relative path found inside an archive, e.g. `("word/document.xml", "../media/a.png")`. */
export function resolveZipPath(fromFile: string, target: string): string {
  if (target.startsWith('/')) return target.slice(1);
  const parts = fromFile.split('/');
  parts.pop();
  for (const segment of target.split('/')) {
    if (segment === '..') parts.pop();
    else if (segment !== '.' && segment !== '') parts.push(segment);
  }
  return parts.join('/');
}
