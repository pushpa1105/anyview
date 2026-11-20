/**
 * The simulated DOM used by the unit tests (happy-dom) never loads iframe pages
 * or external stylesheets, and prints an error with a stack trace every time
 * one is attached. That is expected here, so those messages are dropped to keep
 * real errors visible. happy-dom holds its own console reference, which is why
 * the filter sits on the output stream rather than on `console.error`.
 */
const write = process.stderr.write.bind(process.stderr);

process.stderr.write = (chunk: unknown, ...rest: unknown[]): boolean => {
  if (typeof chunk === 'string' && chunk.includes('loading is disabled')) return true;
  return (write as (...values: unknown[]) => boolean)(chunk, ...rest);
};
