import { describe, expect, it } from 'vitest';
import { AnyViewError, ERROR_CODES, isAnyViewError, toAnyViewError } from '../../src';

describe('AnyViewError', () => {
  it('carries code, message, details and cause', () => {
    const cause = new Error('boom');
    const error = new AnyViewError('FETCH_FAILED', 'Could not fetch', {
      cause,
      details: { status: 404 },
    });
    expect(error).toBeInstanceOf(Error);
    expect(error.name).toBe('AnyViewError');
    expect(error.code).toBe('FETCH_FAILED');
    expect(error.message).toBe('Could not fetch');
    expect(error.details).toEqual({ status: 404 });
    expect(error.cause).toBe(cause);
  });

  it('defaults details to an empty object and leaves cause unset', () => {
    const error = new AnyViewError('PARSE_FAILED', 'Bad file');
    expect(error.details).toEqual({});
    expect('cause' in error).toBe(false);
  });
});

describe('isAnyViewError', () => {
  it('accepts real instances', () => {
    expect(isAnyViewError(new AnyViewError('ABORTED', 'x'))).toBe(true);
  });

  it('accepts structurally identical errors from another bundle', () => {
    expect(isAnyViewError({ name: 'AnyViewError', code: 'RENDER_FAILED' })).toBe(true);
  });

  it.each([null, undefined, 'error', 42, new Error('x'), { name: 'AnyViewError', code: 'NOPE' }])(
    'rejects %o',
    (value) => {
      expect(isAnyViewError(value)).toBe(false);
    },
  );
});

describe('toAnyViewError', () => {
  it('passes AnyView errors through unchanged', () => {
    const original = new AnyViewError('FILE_TOO_LARGE', 'too big');
    expect(toAnyViewError(original, 'RENDER_FAILED', 'fallback')).toBe(original);
  });

  it('maps AbortError to ABORTED', () => {
    const abort = new DOMException('aborted', 'AbortError');
    const error = toAnyViewError(abort, 'FETCH_FAILED', 'fallback');
    expect(error.code).toBe('ABORTED');
    expect(error.cause).toBe(abort);
  });

  it('wraps other errors and appends their message', () => {
    const error = toAnyViewError(new TypeError('bad zip'), 'PARSE_FAILED', 'Could not parse');
    expect(error.code).toBe('PARSE_FAILED');
    expect(error.message).toBe('Could not parse: bad zip');
  });

  it('wraps non-error values without a suffix', () => {
    const error = toAnyViewError('nope', 'RENDER_FAILED', 'Could not render');
    expect(error.message).toBe('Could not render');
    expect(error.cause).toBe('nope');
  });
});

describe('ERROR_CODES', () => {
  it('has no duplicates', () => {
    expect(new Set(ERROR_CODES).size).toBe(ERROR_CODES.length);
  });
});
