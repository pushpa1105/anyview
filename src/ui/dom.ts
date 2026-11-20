/**
 * DOM construction helpers. All content reaches the page through
 * `createElement`/`textContent`; nothing in this package parses markup strings
 * into the live document.
 */

export type Child = Node | string | number | null | undefined | false;

export interface ElementProps {
  class?: string | undefined;
  text?: string | undefined;
  /** CSS declarations keyed by property name (`"font-size"`). Empty values are skipped. */
  style?: Readonly<Record<string, string | number | null | undefined>> | undefined;
  attrs?: Readonly<Record<string, string | number | boolean | null | undefined>> | undefined;
  on?: Readonly<Record<string, (event: Event) => void>> | undefined;
}

export function setStyles(element: HTMLElement | SVGElement, styles: ElementProps['style']): void {
  if (!styles) return;
  for (const property in styles) {
    const value = styles[property];
    if (value === null || value === undefined || value === '') continue;
    element.style.setProperty(property, String(value));
  }
}

export function append(parent: Node, ...children: Child[]): void {
  for (const node of children) {
    if (node === null || node === undefined || node === false) continue;
    parent.appendChild(typeof node === 'object' ? node : document.createTextNode(String(node)));
  }
}

/** Creates an HTML element. */
export function h<K extends keyof HTMLElementTagNameMap>(
  tag: K,
  props?: ElementProps | null,
  ...children: Child[]
): HTMLElementTagNameMap[K] {
  const element = document.createElement(tag);
  if (props) {
    if (props.class) element.className = props.class;
    if (props.text !== undefined) element.textContent = props.text;
    setStyles(element, props.style);
    if (props.attrs) {
      for (const name in props.attrs) {
        const value = props.attrs[name];
        if (value === null || value === undefined || value === false) continue;
        element.setAttribute(name, value === true ? '' : String(value));
      }
    }
    if (props.on) {
      for (const type in props.on) element.addEventListener(type, props.on[type]!);
    }
  }
  append(element, ...children);
  return element;
}

export function clear(element: Element): void {
  while (element.firstChild) element.removeChild(element.firstChild);
}

const SVG_NS = 'http://www.w3.org/2000/svg';

/** 24×24 stroke icons, each a list of SVG path `d` strings. */
const ICON_PATHS = {
  download: ['M12 4v11', 'M7 11l5 5 5-5', 'M5 20h14'],
  print: [
    'M7 9V4h10v5',
    'M7 17H5a1 1 0 0 1-1-1v-6a1 1 0 0 1 1-1h14a1 1 0 0 1 1 1v6a1 1 0 0 1-1 1h-2',
    'M7 14h10v6H7z',
  ],
  fullscreen: ['M4 9V4h5', 'M20 9V4h-5', 'M4 15v5h5', 'M20 15v5h-5'],
  zoomIn: ['M12 6v12', 'M6 12h12'],
  zoomOut: ['M6 12h12'],
  previous: ['M15 5l-7 7 7 7'],
  next: ['M9 5l7 7-7 7'],
  back: ['M11 5l-7 7 7 7', 'M4 12h16'],
  file: ['M6 3h8l4 4v14H6z', 'M14 3v4h4'],
  warning: ['M12 4l9 16H3z', 'M12 10v4', 'M12 17v.5'],
} as const;

export type IconName = keyof typeof ICON_PATHS;

export function icon(name: IconName, size = 18): SVGSVGElement {
  const svg = document.createElementNS(SVG_NS, 'svg');
  svg.setAttribute('viewBox', '0 0 24 24');
  svg.setAttribute('width', String(size));
  svg.setAttribute('height', String(size));
  svg.setAttribute('fill', 'none');
  svg.setAttribute('stroke', 'currentColor');
  svg.setAttribute('stroke-width', '2');
  svg.setAttribute('stroke-linecap', 'round');
  svg.setAttribute('stroke-linejoin', 'round');
  svg.setAttribute('aria-hidden', 'true');
  for (const d of ICON_PATHS[name]) {
    const path = document.createElementNS(SVG_NS, 'path');
    path.setAttribute('d', d);
    svg.appendChild(path);
  }
  return svg;
}

/** An icon button with an accessible name. */
export function iconButton(name: IconName, label: string, onClick: () => void): HTMLButtonElement {
  const button = h('button', {
    class: 'av-btn',
    attrs: { type: 'button', title: label, 'aria-label': label },
    on: { click: onClick },
  });
  button.appendChild(icon(name));
  return button;
}

/** Opens a link in a new tab without giving the target access to this window. */
export function externalLink(href: string, ...children: Child[]): HTMLAnchorElement {
  return h(
    'a',
    { attrs: { href, target: '_blank', rel: 'noopener noreferrer nofollow' } },
    ...children,
  );
}

/** Runs `callback` at most once per animation frame. */
export function rafThrottle(callback: () => void): { (): void; cancel(): void } {
  let handle: number | null = null;
  const schedule =
    typeof requestAnimationFrame === 'function'
      ? requestAnimationFrame
      : (run: FrameRequestCallback): number => setTimeout(() => run(0), 16) as unknown as number;
  const cancel =
    typeof cancelAnimationFrame === 'function'
      ? cancelAnimationFrame
      : (id: number): void => clearTimeout(id);
  const throttled = (): void => {
    if (handle !== null) return;
    handle = schedule(() => {
      handle = null;
      callback();
    });
  };
  throttled.cancel = (): void => {
    if (handle !== null) cancel(handle);
    handle = null;
  };
  return throttled;
}

/** Observes element size changes where `ResizeObserver` exists; returns a disposer. */
export function observeResize(element: Element, callback: () => void): () => void {
  if (typeof ResizeObserver === 'undefined') return () => undefined;
  const observer = new ResizeObserver(() => callback());
  observer.observe(element);
  return () => observer.disconnect();
}
