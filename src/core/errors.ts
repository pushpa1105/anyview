/**
 * Stable, machine-readable error codes. Codes are part of the public API:
 * new codes may be added in minor releases, existing ones are never renamed.
 */
export const ERROR_CODES = [
  'INVALID_SOURCE',
  'INVALID_OPTIONS',
  'UNSUPPORTED_FORMAT',
  'FETCH_FAILED',
  'FILE_TOO_LARGE',
  'PARSE_FAILED',
  'RENDER_FAILED',
  'ABORTED',
  'DESTROYED',
] as const;

export type AnyViewErrorCode = (typeof ERROR_CODES)[number];

export interface AnyViewErrorOptions {
  /** The underlying error, if this error wraps another one. */
  cause?: unknown;
  /** Extra structured context, e.g. `{ status: 404 }` or `{ format: 'docx' }`. */
  details?: Readonly<Record<string, unknown>>;
}

/** The single error type thrown and emitted by AnyView. */
export class AnyViewError extends Error {
  override readonly name = 'AnyViewError';
  readonly code: AnyViewErrorCode;
  readonly details: Readonly<Record<string, unknown>>;

  constructor(code: AnyViewErrorCode, message: string, options: AnyViewErrorOptions = {}) {
    super(message, options.cause === undefined ? undefined : { cause: options.cause });
    this.code = code;
    this.details = options.details ?? {};
  }
}

/** Type guard that also works across bundles/realms where `instanceof` can fail. */
export function isAnyViewError(value: unknown): value is AnyViewError {
  if (value instanceof AnyViewError) return true;
  if (typeof value !== 'object' || value === null) return false;
  const candidate = value as { name?: unknown; code?: unknown };
  return (
    candidate.name === 'AnyViewError' &&
    typeof candidate.code === 'string' &&
    (ERROR_CODES as readonly string[]).includes(candidate.code)
  );
}

/**
 * Normalises any thrown value into an `AnyViewError`. Existing AnyView errors
 * pass through untouched; aborts keep their meaning; everything else is wrapped
 * with `fallbackCode`.
 */
export function toAnyViewError(
  value: unknown,
  fallbackCode: AnyViewErrorCode,
  fallbackMessage: string,
): AnyViewError {
  if (isAnyViewError(value)) return value;
  if (value instanceof Error && value.name === 'AbortError') {
    return new AnyViewError('ABORTED', 'The operation was aborted.', { cause: value });
  }
  const reason = value instanceof Error && value.message ? `: ${value.message}` : '';
  return new AnyViewError(fallbackCode, `${fallbackMessage}${reason}`, { cause: value });
}
