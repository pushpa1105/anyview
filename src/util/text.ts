/**
 * Decodes bytes to text. Honours UTF-8/UTF-16 byte-order marks, then tries
 * strict UTF-8, then falls back to `fallbackEncoding` (Windows-1252 by default,
 * the usual encoding of legacy files that are not valid UTF-8).
 */
export function decodeText(bytes: Uint8Array, fallbackEncoding = 'windows-1252'): string {
  if (bytes.length >= 3 && bytes[0] === 0xef && bytes[1] === 0xbb && bytes[2] === 0xbf) {
    return new TextDecoder('utf-8').decode(bytes.subarray(3));
  }
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder('utf-16le').decode(bytes.subarray(2));
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder('utf-16be').decode(bytes.subarray(2));
  }
  try {
    return new TextDecoder('utf-8', { fatal: true }).decode(bytes);
  } catch {
    return decodeWith(bytes, fallbackEncoding);
  }
}

/** Decodes with a named charset, falling back to Windows-1252 for unknown labels. */
export function decodeWith(bytes: Uint8Array, encoding: string): string {
  try {
    return new TextDecoder(encoding).decode(bytes);
  } catch {
    return new TextDecoder('windows-1252').decode(bytes);
  }
}

/** Heuristic: text files do not contain NUL bytes or many control characters. */
export function looksLikeText(bytes: Uint8Array): boolean {
  const length = Math.min(bytes.length, 4096);
  if (length === 0) return true;
  // UTF-16 with BOM is text even though it is full of NUL bytes.
  if (length >= 2) {
    const bom = (bytes[0]! << 8) | bytes[1]!;
    if (bom === 0xfffe || bom === 0xfeff) return true;
  }
  let suspicious = 0;
  for (let i = 0; i < length; i++) {
    const byte = bytes[i]!;
    if (byte === 0) return false;
    if (byte < 0x09 || (byte > 0x0d && byte < 0x20 && byte !== 0x1b)) suspicious++;
  }
  return suspicious / length < 0.02;
}

/** `1536` → `"1.5 KB"`. */
export function formatBytes(bytes: number): string {
  if (!Number.isFinite(bytes) || bytes < 0) return '';
  const units = ['B', 'KB', 'MB', 'GB', 'TB'];
  let value = bytes;
  let unit = 0;
  while (value >= 1024 && unit < units.length - 1) {
    value /= 1024;
    unit++;
  }
  const digits = unit === 0 || value >= 100 ? 0 : 1;
  return `${value.toFixed(digits)} ${units[unit]!}`;
}

/** Decodes a base64 string (whitespace tolerated) to bytes. Invalid input yields what could be decoded. */
export function base64ToBytes(base64: string): Uint8Array {
  const clean = base64.replace(/[^A-Za-z0-9+/]/g, '');
  const usable = clean.length - (clean.length % 4 === 1 ? 1 : 0);
  let binary: string;
  try {
    binary = atob(clean.slice(0, usable));
  } catch {
    return new Uint8Array(0);
  }
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return bytes;
}

/** Copies a view into a standalone `ArrayBuffer`-backed part usable in a `Blob`. */
export function toBlobPart(bytes: Uint8Array): ArrayBuffer {
  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength) as ArrayBuffer;
}

/**
 * Maps each byte to the character with the same code (0–255), so that binary
 * data can be scanned as a string and turned back into identical bytes.
 * `TextDecoder('latin1')` cannot be used for this: browsers treat that label
 * as Windows-1252, which remaps bytes 0x80–0x9F.
 */
export function bytesToBinaryString(bytes: Uint8Array): string {
  let result = '';
  const chunk = 0x8000;
  for (let i = 0; i < bytes.length; i += chunk) {
    result += String.fromCharCode(...bytes.subarray(i, i + chunk));
  }
  return result;
}
