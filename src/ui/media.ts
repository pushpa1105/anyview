import { AnyViewError } from '../core/errors';

/** Resolves when `element` fires `readyEvent`, rejects on `error` or abort. */
export function waitForLoad(
  element: HTMLElement,
  readyEvent: string,
  signal: AbortSignal,
  failureMessage: string,
): Promise<void> {
  return new Promise<void>((resolve, reject) => {
    const cleanup = (): void => {
      element.removeEventListener(readyEvent, onReady);
      element.removeEventListener('error', onError);
      signal.removeEventListener('abort', onAbort);
    };
    const onReady = (): void => {
      cleanup();
      resolve();
    };
    const onError = (): void => {
      cleanup();
      reject(new AnyViewError('RENDER_FAILED', failureMessage));
    };
    const onAbort = (): void => {
      cleanup();
      reject(new AnyViewError('ABORTED', 'The operation was aborted.'));
    };
    if (signal.aborted) {
      onAbort();
      return;
    }
    element.addEventListener(readyEvent, onReady);
    element.addEventListener('error', onError);
    signal.addEventListener('abort', onAbort);
  });
}
