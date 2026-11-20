import { AnyViewError } from '../core/errors';
import { attr, childElements, parseXml, type XmlElement, type XmlNode } from '../util/xml';
import type { ZipArchive } from '../util/zip';

/** Helpers shared by the OpenDocument formats (ODT, ODS, ODP). */

/** Reads `content.xml` (or another part) keeping whitespace, which is significant in ODF text. */
export async function readOdfPart(zip: ZipArchive, path: string): Promise<XmlElement | null> {
  const xml = await zip.readText(path);
  return xml === null ? null : parseXml(xml, { preserveWhitespace: true });
}

export async function requireOdfContent(zip: ZipArchive, kind: string): Promise<XmlElement> {
  const content = await readOdfPart(zip, 'content.xml');
  if (!content) {
    throw new AnyViewError(
      'PARSE_FAILED',
      `This is not a valid ${kind}: "content.xml" is missing.`,
    );
  }
  return content;
}

/**
 * Plain text of an ODF text node, expanding `<text:s>` (repeated spaces),
 * `<text:tab>` and `<text:line-break>`.
 */
export function odfText(node: XmlNode): string {
  if (typeof node === 'string') return node;
  switch (node.local) {
    case 's':
      return ' '.repeat(Math.min(Number(attr(node, 'text:c') ?? 1) || 1, 1000));
    case 'tab':
      return '\t';
    case 'line-break':
      return '\n';
    case 'note':
    case 'annotation':
      return '';
    default:
      return node.children.map(odfText).join('');
  }
}

export interface OdfStyle {
  parent: string | undefined;
  family: string | undefined;
  /** Merged `fo:*`/`style:*` properties from the style's property elements, by local name. */
  properties: Record<string, string>;
  /** For list styles: whether each level is numbered. */
  element: XmlElement;
}

/** Collects named styles from `<office:styles>` and `<office:automatic-styles>` blocks. */
export function collectOdfStyles(...roots: (XmlElement | null)[]): Map<string, OdfStyle> {
  const styles = new Map<string, OdfStyle>();
  for (const root of roots) {
    if (!root) continue;
    for (const block of childElements(root)) {
      if (block.local !== 'styles' && block.local !== 'automatic-styles') continue;
      for (const style of childElements(block)) {
        const name = attr(style, 'style:name');
        if (!name) continue;
        const properties: Record<string, string> = {};
        for (const group of childElements(style)) {
          if (!group.local.endsWith('-properties')) continue;
          for (const key in group.attrs) {
            properties[key.slice(key.indexOf(':') + 1)] = group.attrs[key]!;
          }
        }
        styles.set(name, {
          parent: attr(style, 'style:parent-style-name'),
          family: attr(style, 'style:family'),
          properties,
          element: style,
        });
      }
    }
  }
  return styles;
}

/** Resolves a style's properties including everything inherited from its parents. */
export function resolveOdfStyle(
  styles: Map<string, OdfStyle>,
  name: string | undefined,
): Record<string, string> {
  const chain: OdfStyle[] = [];
  const seen = new Set<string>();
  let current = name;
  while (current && !seen.has(current)) {
    seen.add(current);
    const style = styles.get(current);
    if (!style) break;
    chain.unshift(style);
    current = style.parent;
  }
  return Object.assign({}, ...chain.map((style) => style.properties)) as Record<string, string>;
}

const LENGTH = /^(-?\d*\.?\d+)(cm|mm|in|pt|pc|px)?$/;
const PX_PER_UNIT: Readonly<Record<string, number>> = {
  cm: 96 / 2.54,
  mm: 96 / 25.4,
  in: 96,
  pt: 96 / 72,
  pc: 16,
  px: 1,
};

/** `"2.5cm"` → pixels. Unknown or missing values → `undefined`. */
export function odfLength(value: string | undefined): number | undefined {
  const match = value ? LENGTH.exec(value.trim()) : null;
  if (!match) return undefined;
  return Number(match[1]) * (PX_PER_UNIT[match[2] ?? 'px'] ?? 1);
}

/** Accepts only plain colour values (`#rrggbb`), never arbitrary CSS. */
export function odfColor(value: string | undefined): string | undefined {
  return value && /^#[0-9a-fA-F]{6}$/.test(value) ? value : undefined;
}

/** Maps resolved text properties to CSS declarations. */
export function odfTextCss(properties: Record<string, string>): Record<string, string | undefined> {
  const size = properties['font-size'];
  const underline = properties['text-underline-style'];
  const strike = properties['text-line-through-style'];
  const decoration = [
    underline && underline !== 'none' ? 'underline' : '',
    strike && strike !== 'none' ? 'line-through' : '',
  ]
    .filter(Boolean)
    .join(' ');
  const position = properties['text-position'] ?? '';
  return {
    'font-weight': properties['font-weight'] === 'bold' ? 'bold' : undefined,
    'font-style': properties['font-style'] === 'italic' ? 'italic' : undefined,
    'text-decoration': decoration || undefined,
    color: odfColor(properties['color']),
    'background-color': odfColor(properties['background-color']),
    'font-size': size && /^\d*\.?\d+pt$/.test(size) ? size : undefined,
    'vertical-align': position.startsWith('super')
      ? 'super'
      : position.startsWith('sub')
        ? 'sub'
        : undefined,
  };
}

/** Maps resolved paragraph properties to CSS declarations. */
export function odfParagraphCss(
  properties: Record<string, string>,
): Record<string, string | undefined> {
  const align = properties['text-align'];
  const length = (key: string): string | undefined => {
    const px = odfLength(properties[key]);
    return px !== undefined && px !== 0 ? `${px.toFixed(1)}px` : undefined;
  };
  return {
    'text-align':
      align === 'center'
        ? 'center'
        : align === 'end' || align === 'right'
          ? 'right'
          : align === 'justify'
            ? 'justify'
            : undefined,
    'margin-left': length('margin-left'),
    'margin-right': length('margin-right'),
    'margin-top': length('margin-top'),
    'margin-bottom': length('margin-bottom'),
    'text-indent': length('text-indent'),
  };
}
