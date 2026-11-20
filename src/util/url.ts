const SCHEME = /^([a-zA-Z][a-zA-Z0-9+.-]*):/;
const LINK_SCHEMES = new Set(['http', 'https', 'mailto', 'tel']);
const SAFE_DATA_IMAGE = /^data:image\/(png|jpe?g|gif|webp|avif|bmp|x-icon);base64,/i;

function normalise(url: string): string {
  // Browsers ignore tabs/newlines inside URLs and leading control chars: so must we,
  // otherwise "java\nscript:" would slip through a naive scheme check.
  let start = 0;
  let end = url.length;
  // Leading and trailing spaces and control characters are stripped…
  while (start < end && url.charCodeAt(start) <= 0x20) start++;
  while (end > start && url.charCodeAt(end - 1) <= 0x20) end--;
  // …and tabs and line breaks are removed wherever they appear.
  return url.slice(start, end).replace(/[\t\n\r]/g, '');
}

function resolve(url: string, base: string | null | undefined): string | null {
  if (!base) return url;
  try {
    return new URL(url, base).href;
  } catch {
    return null;
  }
}

/**
 * Returns a URL that is safe to put in an `href`, or `null`.
 * Only http(s), mailto and tel are allowed; relative URLs are resolved against `base`.
 */
export function safeLinkUrl(url: string | null | undefined, base?: string | null): string | null {
  if (!url) return null;
  const clean = normalise(url);
  if (!clean) return null;
  const scheme = SCHEME.exec(clean)?.[1]?.toLowerCase();
  if (scheme) return LINK_SCHEMES.has(scheme) ? clean : null;
  if (clean.startsWith('#')) return clean;
  const resolved = resolve(clean, base);
  if (resolved === null) return null;
  // A base with an unsafe scheme must not turn a relative link into an unsafe one.
  const resolvedScheme = SCHEME.exec(resolved)?.[1]?.toLowerCase();
  if (resolvedScheme && !LINK_SCHEMES.has(resolvedScheme)) return null;
  return resolved;
}

/** Returns a URL that is safe to put in an `<img src>`, or `null`. */
export function safeImageUrl(url: string | null | undefined, base?: string | null): string | null {
  if (!url) return null;
  const clean = normalise(url);
  if (!clean) return null;
  const scheme = SCHEME.exec(clean)?.[1]?.toLowerCase();
  if (scheme === 'http' || scheme === 'https' || scheme === 'blob') return clean;
  if (scheme === 'data') return SAFE_DATA_IMAGE.test(clean) ? clean : null;
  if (scheme) return null;
  const resolved = resolve(clean, base);
  if (resolved === null) return null;
  const resolvedScheme = SCHEME.exec(resolved)?.[1]?.toLowerCase();
  if (resolvedScheme && !['http', 'https', 'blob'].includes(resolvedScheme)) return null;
  return resolved;
}

/** Last path segment of a URL, decoded; `""` when there is none. */
export function fileNameFromUrl(url: string): string {
  let path: string;
  try {
    path = new URL(url, 'http://localhost/').pathname;
  } catch {
    path = url.split(/[?#]/)[0] ?? '';
  }
  const segment = path.split('/').filter(Boolean).pop() ?? '';
  try {
    return decodeURIComponent(segment);
  } catch {
    return segment;
  }
}

/** Lower-case extension without the dot; `""` when there is none. */
export function extensionOf(name: string): string {
  const base = name.split(/[\\/]/).pop() ?? '';
  const dot = base.lastIndexOf('.');
  if (dot <= 0 || dot === base.length - 1) return '';
  return base.slice(dot + 1).toLowerCase();
}
