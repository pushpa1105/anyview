import { AnyViewError } from '../core/errors';
import { attr, childElements, parseXml, type XmlElement } from '../util/xml';
import { resolveZipPath, type ZipArchive } from '../util/zip';

/** Helpers shared by the Office Open XML formats (DOCX, XLSX, PPTX). */

export interface Relationship {
  id: string;
  type: string;
  /** Path inside the package, or the URL for external targets. */
  target: string;
  external: boolean;
}

/** Path of the relationships part that belongs to `partPath`. */
export function relsPathFor(partPath: string): string {
  const slash = partPath.lastIndexOf('/');
  const folder = slash === -1 ? '' : partPath.slice(0, slash + 1);
  const name = slash === -1 ? partPath : partPath.slice(slash + 1);
  return `${folder}_rels/${name}.rels`;
}

/** Reads the relationships of a part, keyed by relationship id. Missing file → empty map. */
export async function readRelationships(
  zip: ZipArchive,
  partPath: string,
): Promise<Map<string, Relationship>> {
  const relationships = new Map<string, Relationship>();
  const xml = await zip.readText(relsPathFor(partPath));
  if (!xml) return relationships;
  for (const element of childElements(parseXml(xml), 'Relationship')) {
    const id = element.attrs['Id'];
    const target = element.attrs['Target'];
    if (!id || target === undefined) continue;
    const external = element.attrs['TargetMode'] === 'External';
    relationships.set(id, {
      id,
      type: (element.attrs['Type'] ?? '').split('/').pop() ?? '',
      target: external ? target : resolveZipPath(partPath, target),
      external,
    });
  }
  return relationships;
}

/** Reads and parses an XML part; `null` when the part does not exist. */
export async function readXmlPart(zip: ZipArchive, path: string): Promise<XmlElement | null> {
  const xml = await zip.readText(path);
  return xml === null ? null : parseXml(xml);
}

/** Reads a part that must exist; throws a descriptive error otherwise. */
export async function requireXmlPart(
  zip: ZipArchive,
  path: string,
  kind: string,
): Promise<XmlElement> {
  const part = await readXmlPart(zip, path);
  if (!part) {
    throw new AnyViewError('PARSE_FAILED', `This is not a valid ${kind}: "${path}" is missing.`);
  }
  return part;
}

/** Boolean OOXML attribute/element value: absent `val` means true. */
export function isOn(element: XmlElement | null | undefined, name = 'w:val'): boolean {
  if (!element) return false;
  const value = attr(element, name);
  return (
    value === undefined ||
    (value !== '0' && value !== 'false' && value !== 'off' && value !== 'none')
  );
}

const IMAGE_TYPES: Readonly<Record<string, string>> = {
  png: 'image/png',
  jpg: 'image/jpeg',
  jpeg: 'image/jpeg',
  jfif: 'image/jpeg',
  gif: 'image/gif',
  bmp: 'image/bmp',
  webp: 'image/webp',
  svg: 'image/svg+xml',
  avif: 'image/avif',
  ico: 'image/x-icon',
};

/** MIME type for an embedded picture, or `null` for formats browsers cannot show (EMF, WMF, TIFF). */
export function imageMimeType(path: string): string | null {
  const extension = path.slice(path.lastIndexOf('.') + 1).toLowerCase();
  return IMAGE_TYPES[extension] ?? null;
}

/** `"FF0000"` → `"#FF0000"`; anything that is not a 6-digit hex colour → `undefined`. */
export function hexColor(value: string | undefined): string | undefined {
  if (!value) return undefined;
  const hex = value.length === 8 ? value.slice(2) : value;
  return /^[0-9a-fA-F]{6}$/.test(hex) ? `#${hex}` : undefined;
}

/** English Metric Units → CSS pixels. */
export function emuToPx(emu: number): number {
  return emu / 9525;
}
