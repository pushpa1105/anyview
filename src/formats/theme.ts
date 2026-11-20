import { attr, child, childElements, find, type XmlElement } from '../util/xml';
import type { ZipArchive } from '../util/zip';
import { hexColor, readXmlPart } from './ooxml';

/** Colours and fonts of an Office theme part (`theme1.xml`), shared by DOCX and XLSX. */
export interface OfficeTheme {
  /** Scheme colours by name (`dk1`, `lt1`, `dk2`, `lt2`, `accent1`…`accent6`, `hlink`, `folHlink`). */
  colors: Map<string, string>;
  /** Heading font. */
  majorFont: string | undefined;
  /** Body font. */
  minorFont: string | undefined;
}

export const EMPTY_THEME: OfficeTheme = {
  colors: new Map(),
  majorFont: undefined,
  minorFont: undefined,
};

/** Font names end up inside CSS `font-family`, so only plain names are accepted. */
export function safeFontName(name: string | undefined): string | undefined {
  return name && /^[\w .,'()&+-]{1,64}$/.test(name) ? name.replace(/'/g, '') : undefined;
}

export function parseTheme(root: XmlElement | null): OfficeTheme {
  const colors = new Map<string, string>();
  for (const entry of childElements(find(root, 'themeElements', 'clrScheme'))) {
    const value = child(entry, 'srgbClr') ?? child(entry, 'sysClr');
    const hex = hexColor(attr(value, 'lastClr')) ?? hexColor(attr(value, 'val'));
    if (hex) colors.set(entry.local, hex);
  }
  const fonts = find(root, 'themeElements', 'fontScheme');
  const typeface = (kind: string): string | undefined =>
    safeFontName(attr(find(fonts, kind, 'latin'), 'typeface'));
  return { colors, majorFont: typeface('majorFont'), minorFont: typeface('minorFont') };
}

export async function loadTheme(zip: ZipArchive, path: string | undefined): Promise<OfficeTheme> {
  if (!path) return EMPTY_THEME;
  return parseTheme(await readXmlPart(zip, path));
}

function channels(hex: string): [number, number, number] {
  return [
    parseInt(hex.slice(1, 3), 16),
    parseInt(hex.slice(3, 5), 16),
    parseInt(hex.slice(5, 7), 16),
  ];
}

function toHex(red: number, green: number, blue: number): string {
  return `#${[red, green, blue]
    .map((value) =>
      Math.max(0, Math.min(255, Math.round(value)))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

/** Mixes a colour with white: `amount` 1 keeps the colour, 0 gives white (Word `themeTint`). */
export function tintColor(hex: string, amount: number): string {
  const [r, g, b] = channels(hex);
  return toHex(
    r * amount + 255 * (1 - amount),
    g * amount + 255 * (1 - amount),
    b * amount + 255 * (1 - amount),
  );
}

/** Mixes a colour with black: `amount` 1 keeps the colour, 0 gives black (Word `themeShade`). */
export function shadeColor(hex: string, amount: number): string {
  const [r, g, b] = channels(hex);
  return toHex(r * amount, g * amount, b * amount);
}

/**
 * SpreadsheetML `tint`: −1…1, applied to the luminance of the colour in HLS
 * space. Negative values darken, positive values lighten.
 */
export function applySheetTint(hex: string, tint: number): string {
  if (!tint) return hex;
  const [r, g, b] = channels(hex).map((value) => value / 255) as [number, number, number];
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  let hue = 0;
  let saturation = 0;
  let luminance = (max + min) / 2;
  if (max !== min) {
    const delta = max - min;
    saturation = luminance > 0.5 ? delta / (2 - max - min) : delta / (max + min);
    if (max === r) hue = (g - b) / delta + (g < b ? 6 : 0);
    else if (max === g) hue = (b - r) / delta + 2;
    else hue = (r - g) / delta + 4;
    hue /= 6;
  }
  luminance = tint < 0 ? luminance * (1 + tint) : luminance * (1 - tint) + tint;
  if (saturation === 0) return toHex(luminance * 255, luminance * 255, luminance * 255);
  const q =
    luminance < 0.5
      ? luminance * (1 + saturation)
      : luminance + saturation - luminance * saturation;
  const p = 2 * luminance - q;
  const channel = (t: number): number => {
    const wrapped = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (wrapped < 1 / 6) return p + (q - p) * 6 * wrapped;
    if (wrapped < 1 / 2) return q;
    if (wrapped < 2 / 3) return p + (q - p) * (2 / 3 - wrapped) * 6;
    return p;
  };
  return toHex(channel(hue + 1 / 3) * 255, channel(hue) * 255, channel(hue - 1 / 3) * 255);
}
