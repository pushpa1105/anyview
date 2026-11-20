import { decodeText, toBlobPart } from '../util/text';
import { fileNameFromUrl } from '../util/url';
import { AnyViewError, toAnyViewError } from './errors';
import type { LoadedFile, ViewerSource } from './types';

export interface SourceOptions {
  name?: string | undefined;
  fetchOptions?: RequestInit | undefined;
  maxFileSize: number;
  signal: AbortSignal;
}

function tooLarge(size: number, limit: number): AnyViewError {
  return new AnyViewError(
    'FILE_TOO_LARGE',
    `The file is larger than the ${Math.round(limit / (1024 * 1024))} MB limit.`,
    { details: { size, limit } },
  );
}

function fileNameFromDisposition(header: string | null): string {
  if (!header) return '';
  const extended = /filename\*\s*=\s*(?:UTF-8|utf-8)''([^;]+)/.exec(header)?.[1];
  if (extended) {
    try {
      return decodeURIComponent(extended.trim());
    } catch {
      // Fall through to the plain parameter.
    }
  }
  const plain = /filename\s*=\s*(?:"([^"]*)"|([^;]+))/i.exec(header);
  return (plain?.[1] ?? plain?.[2] ?? '').trim();
}

function describe(value: unknown): string {
  if (value === null) return 'null';
  if (typeof value === 'object') return Object.prototype.toString.call(value);
  return typeof value;
}

/**
 * Wraps any supported source in a {@link LoadedFile}. Nothing is downloaded
 * here: remote content is fetched on first use and cached, so media files can
 * be streamed by the browser instead of being buffered in memory.
 */
export function createFile(source: ViewerSource, options: SourceOptions): LoadedFile {
  const { maxFileSize, signal } = options;

  if (typeof Blob !== 'undefined' && source instanceof Blob) {
    return fromBlob(
      source,
      options.name ?? (source as Partial<File>).name ?? '',
      null,
      maxFileSize,
    );
  }
  if (source instanceof ArrayBuffer) {
    return fromBlob(new Blob([source]), options.name ?? '', null, maxFileSize);
  }
  if (ArrayBuffer.isView(source)) {
    const view = new Uint8Array(source.buffer, source.byteOffset, source.byteLength);
    return fromBlob(new Blob([toBlobPart(view)]), options.name ?? '', null, maxFileSize);
  }
  if (typeof source === 'string' || (typeof URL !== 'undefined' && source instanceof URL)) {
    const text = String(source).trim();
    if (!text) throw new AnyViewError('INVALID_SOURCE', 'The source URL is empty.');
    let url: string;
    try {
      const base = typeof document !== 'undefined' ? document.baseURI : undefined;
      url = new URL(text, base).href;
    } catch (cause) {
      throw new AnyViewError('INVALID_SOURCE', `"${text}" is not a valid URL.`, { cause });
    }
    return fromUrl(url, options.name, options.fetchOptions, maxFileSize, signal);
  }
  throw new AnyViewError(
    'INVALID_SOURCE',
    `Unsupported source type ${describe(source)}. Pass a URL, File, Blob, ArrayBuffer or typed array.`,
  );
}

function fromBlob(blob: Blob, name: string, url: string | null, maxFileSize: number): LoadedFile {
  let bytes: Promise<Uint8Array> | null = null;
  let text: Promise<string> | null = null;
  const checked = (): Blob => {
    if (blob.size > maxFileSize) throw tooLarge(blob.size, maxFileSize);
    return blob;
  };
  const file: LoadedFile = {
    name,
    mimeType: blob.type.split(';')[0]!.trim().toLowerCase(),
    size: blob.size,
    url,
    direct: false,
    blob: () => Promise.resolve().then(checked),
    bytes: () =>
      (bytes ??= Promise.resolve().then(async () => new Uint8Array(await checked().arrayBuffer()))),
    text: () => (text ??= file.bytes().then((data) => decodeText(data))),
  };
  return file;
}

function fromUrl(
  url: string,
  nameOverride: string | undefined,
  fetchOptions: RequestInit | undefined,
  maxFileSize: number,
  signal: AbortSignal,
): LoadedFile {
  const isData = url.startsWith('data:');
  let download: Promise<LoadedFile> | null = null;
  let resolved: LoadedFile | null = null;

  const fetchFile = (): Promise<LoadedFile> =>
    (download ??= (async () => {
      let response: Response;
      try {
        response = await fetch(url, { ...fetchOptions, signal });
      } catch (cause) {
        throw toAnyViewError(
          cause,
          'FETCH_FAILED',
          'The file could not be downloaded (network error or blocked by CORS)',
        );
      }
      if (!response.ok) {
        throw new AnyViewError(
          'FETCH_FAILED',
          `The file could not be downloaded (HTTP ${response.status}).`,
          { details: { status: response.status, url: isData ? 'data:' : url } },
        );
      }
      const declared = Number(response.headers.get('content-length'));
      if (Number.isFinite(declared) && declared > maxFileSize)
        throw tooLarge(declared, maxFileSize);
      let blob: Blob;
      try {
        blob = await response.blob();
      } catch (cause) {
        throw toAnyViewError(cause, 'FETCH_FAILED', 'The download was interrupted');
      }
      const name =
        nameOverride ??
        (fileNameFromDisposition(response.headers.get('content-disposition')) || file.name);
      resolved = fromBlob(blob, name, url, maxFileSize);
      return resolved;
    })());

  const file: LoadedFile = {
    get name() {
      return resolved?.name ?? nameOverride ?? (isData ? '' : fileNameFromUrl(url));
    },
    get mimeType() {
      if (resolved) return resolved.mimeType;
      return isData ? (/^data:([^;,]+)/.exec(url)?.[1] ?? '').toLowerCase() : '';
    },
    get size() {
      return resolved?.size ?? null;
    },
    url,
    // Custom fetch options (auth headers…) cannot be applied by <img>/<video>, so
    // such files must go through fetch() and be shown from an object URL instead.
    direct: fetchOptions === undefined && !isData,
    blob: () => fetchFile().then((loaded) => loaded.blob()),
    bytes: () => fetchFile().then((loaded) => loaded.bytes()),
    text: () => fetchFile().then((loaded) => loaded.text()),
  };
  return file;
}
