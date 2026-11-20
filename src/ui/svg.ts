import { append, type Child } from './dom';

const SVG_NS = 'http://www.w3.org/2000/svg';

/** Creates an SVG element. Attributes with empty values are skipped. */
export function svg(
  tag: string,
  attrs?: Readonly<Record<string, string | number | null | undefined>> | null,
  ...children: Child[]
): SVGElement {
  const element = document.createElementNS(SVG_NS, tag);
  if (attrs) {
    for (const name in attrs) {
      const value = attrs[name];
      if (value === null || value === undefined || value === '') continue;
      element.setAttribute(name, String(value));
    }
  }
  append(element, ...children);
  return element;
}

/** Rounds a coordinate so generated markup stays short and stable. */
export const round = (value: number): number => Math.round(value * 100) / 100;
