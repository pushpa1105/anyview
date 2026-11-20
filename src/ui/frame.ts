import { h } from './dom';

function escapeAttribute(value: string): string {
  return value.replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/</g, '&lt;');
}

/**
 * Relative URLs inside a srcdoc document would resolve against the host page.
 * Point them at the file's own location instead, or nowhere when it has none.
 */
export function withBase(html: string, baseUrl: string | null): string {
  const base = `<base href="${escapeAttribute(baseUrl ?? 'about:invalid')}">`;
  const head = /<head(\s[^>]*)?>/i.exec(html);
  if (head) {
    const at = head.index + head[0].length;
    return html.slice(0, at) + base + html.slice(at);
  }
  return base + html;
}

/**
 * Creates an iframe for untrusted HTML. The empty `sandbox` attribute applies
 * every restriction: no scripts, no forms, no same-origin access, no popups and
 * no top-level navigation, so the document can be displayed but cannot act.
 */
export function createSandboxedFrame(
  html: string,
  title: string,
  baseUrl?: string | null,
): HTMLIFrameElement {
  const frame = h('iframe', {
    class: 'av-frame',
    attrs: { sandbox: '', title, referrerpolicy: 'no-referrer' },
  });
  frame.srcdoc = withBase(html, baseUrl ?? null);
  return frame;
}
