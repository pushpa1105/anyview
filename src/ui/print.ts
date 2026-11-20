import { CSS } from './styles';

const PRINT_CSS = `
html,body{margin:0;padding:0;background:#fff}
.av-root{display:block;height:auto;min-height:0;overflow:visible;background:#fff}
.av-root *{overflow:visible!important;max-height:none!important}
.av-prose,.av-paper{box-shadow:none;margin:0;max-width:none}
.av-page{box-shadow:none;break-after:page;margin:0 auto}
.av-doc{display:block;zoom:1!important}
.av-doc-page{break-after:page;margin:0 auto}
.av-code-gutter{position:static}
img{max-width:100%}
`;

/**
 * Prints a copy of `node` from a hidden same-origin iframe, so only the file
 * (not the surrounding application) ends up on paper.
 */
export function printNode(node: Node, title: string): void {
  const frame = document.createElement('iframe');
  frame.setAttribute('aria-hidden', 'true');
  frame.style.cssText =
    'position:fixed;right:0;bottom:0;width:0;height:0;border:0;visibility:hidden';
  document.body.appendChild(frame);
  const target = frame.contentDocument;
  const view = frame.contentWindow;
  if (!target || !view) {
    frame.remove();
    window.print();
    return;
  }
  target.title = title;
  const style = target.createElement('style');
  style.textContent = CSS + PRINT_CSS;
  (target.head ?? target.documentElement).appendChild(style);
  const root = target.createElement('div');
  root.className = 'av-root';
  root.setAttribute('data-theme', 'light');
  root.appendChild(target.importNode(node, true));
  target.body.appendChild(root);

  // A cloned <canvas> is blank, so each one is replaced by a snapshot of the original.
  const originals =
    node.nodeType === 1 ? Array.from((node as Element).querySelectorAll('canvas')) : [];
  Array.from(root.querySelectorAll('canvas')).forEach((copy, index) => {
    const original = originals[index];
    if (!original) return;
    try {
      const snapshot = target.createElement('img');
      snapshot.src = original.toDataURL('image/png');
      snapshot.style.cssText = original.style.cssText;
      copy.replaceWith(snapshot);
    } catch {
      // A tainted canvas cannot be exported; leave the blank copy in place.
    }
  });

  let done = false;
  const cleanup = (): void => {
    if (done) return;
    done = true;
    frame.remove();
  };
  const run = (): void => {
    view.addEventListener('afterprint', cleanup, { once: true });
    try {
      view.focus();
      view.print();
    } finally {
      // Fallback for browsers that never fire afterprint on iframes.
      setTimeout(cleanup, 60_000);
    }
  };

  const images = Array.from(root.querySelectorAll('img')).filter((image) => !image.complete);
  if (images.length === 0) {
    run();
    return;
  }
  let pending = images.length;
  let started = false;
  const start = (): void => {
    if (started) return;
    started = true;
    run();
  };
  const settle = (): void => {
    pending--;
    if (pending <= 0) start();
  };
  for (const image of images) {
    image.addEventListener('load', settle, { once: true });
    image.addEventListener('error', settle, { once: true });
  }
  setTimeout(start, 1500);
}
