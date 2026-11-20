import type { GridBorder, GridCellStyle } from '../ui/grid';
import { BUILTIN_FORMATS } from '../util/numfmt';
import { attr, child, childElements, descendants, type XmlElement } from '../util/xml';
import { hexColor } from './ooxml';
import { applySheetTint, safeFontName, type OfficeTheme } from './theme';

/** SpreadsheetML styles: fonts, fills, borders, cell formats, differential formats, table styles. */

export interface FontDef {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  name?: string;
  /** Size in points. */
  size?: number;
}

/** A side that is `null` is explicitly cleared; `undefined` means "not specified". */
export interface BorderDef {
  top?: GridBorder | null;
  right?: GridBorder | null;
  bottom?: GridBorder | null;
  left?: GridBorder | null;
  /** Inner lines of a region; only meaningful in differential formats. */
  horizontal?: GridBorder | null;
  vertical?: GridBorder | null;
}

export interface AlignDef {
  horizontal?: string;
  vertical?: string;
  wrap?: boolean;
  indent?: number;
}

export interface CellXf {
  numberFormat: string;
  font: FontDef;
  /** CSS background (a colour or gradient), when the cell is filled. */
  fill: string | undefined;
  border: BorderDef;
  align: AlignDef;
  /** Whether the font/fill/border were set on the cell rather than inherited from its style. */
  explicitFont: boolean;
  explicitFill: boolean;
  explicitBorder: boolean;
}

/** Differential format: only the listed properties change. */
export interface Dxf {
  font?: FontDef;
  /** `null` removes the fill. */
  fill?: string | null;
  border?: BorderDef;
  numberFormat?: string;
}

export interface TableStyleDef {
  /** Differential formats by element type (`wholeTable`, `headerRow`, `firstRowStripe`…). */
  elements: Map<string, { dxf: Dxf; size: number }>;
}

export interface StyleSheet {
  cellXfs: CellXf[];
  dxfs: Dxf[];
  tableStyles: Map<string, TableStyleDef>;
  defaultTableStyle: string | undefined;
  /** Font of the Normal style; column widths are measured in its digits. */
  defaultFont: FontDef;
  theme: OfficeTheme;
  /** Resolves a `<color>`-like element (rgb, theme + tint, indexed). */
  color(element: XmlElement | null): string | undefined;
  /** Parses a `<font>` or rich-text `<rPr>` element. */
  font(element: XmlElement | null): FontDef;
}

/** Excel's default 56-colour palette (indexes 8–63; 0–7 repeat 8–15). */
const INDEXED =
  '000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 000000 FFFFFF FF0000 00FF00 0000FF FFFF00 FF00FF 00FFFF 800000 008000 000080 808000 800080 008080 C0C0C0 808080 9999FF 993366 FFFFCC CCFFFF 660066 FF8080 0066CC CCCCFF 000080 FF00FF FFFF00 00FFFF 800080 800000 008080 0000FF 00CCFF CCFFFF CCFFCC FFFF99 99CCFF FF99CC CC99FF FFCC99 3366FF 33CCCC 99CC00 FFCC00 FF9900 FF6600 666699 969696 003366 339966 003300 333300 993300 993366 333399 333333'.split(
    ' ',
  );

/** Theme colour order used by `theme="n"` attributes. */
const THEME_ORDER = [
  'lt1',
  'dk1',
  'lt2',
  'dk2',
  'accent1',
  'accent2',
  'accent3',
  'accent4',
  'accent5',
  'accent6',
  'hlink',
  'folHlink',
];
const THEME_FALLBACK = [
  '#ffffff',
  '#000000',
  '#e7e6e6',
  '#44546a',
  '#4472c4',
  '#ed7d31',
  '#a5a5a5',
  '#ffc000',
  '#5b9bd5',
  '#70ad47',
  '#0563c1',
  '#954f72',
];

export function themeColorAt(theme: OfficeTheme, index: number, tint = 0): string | undefined {
  const name = THEME_ORDER[index];
  if (!name) return undefined;
  return applySheetTint((theme.colors.get(name) ?? THEME_FALLBACK[index]!).toLowerCase(), tint);
}

const BORDER_STYLES: Readonly<Record<string, [number, GridBorder['style']]>> = {
  thin: [1, 'solid'],
  medium: [2, 'solid'],
  thick: [3, 'solid'],
  double: [3, 'double'],
  hair: [1, 'dotted'],
  dotted: [1, 'dotted'],
  dashed: [1, 'dashed'],
  dashDot: [1, 'dashed'],
  dashDotDot: [1, 'dotted'],
  mediumDashed: [2, 'dashed'],
  mediumDashDot: [2, 'dashed'],
  mediumDashDotDot: [2, 'dotted'],
  slantDashDot: [2, 'dashed'],
};

/** Share of foreground colour in each pattern fill. */
const PATTERN_DENSITY: Readonly<Record<string, number>> = {
  darkGray: 0.75,
  mediumGray: 0.5,
  lightGray: 0.25,
  gray125: 0.125,
  gray0625: 0.0625,
};

function mix(foreground: string, background: string, share: number): string {
  const part = (hex: string, at: number): number => parseInt(hex.slice(at, at + 2), 16);
  return `#${[1, 3, 5]
    .map((at) =>
      Math.round(part(foreground, at) * share + part(background, at) * (1 - share))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

const flagOn = (element: XmlElement | null): boolean | undefined => {
  if (!element) return undefined;
  const value = attr(element, 'val');
  return value !== '0' && value !== 'false';
};

export function parseStyles(root: XmlElement | null, theme: OfficeTheme): StyleSheet {
  const customIndexed = childElements(
    child(child(root, 'colors'), 'indexedColors'),
    'rgbColor',
  ).map((entry) => hexColor(attr(entry, 'rgb')));

  const color = (element: XmlElement | null): string | undefined => {
    if (!element) return undefined;
    const rgb = hexColor(attr(element, 'rgb'));
    if (rgb) return rgb;
    const tint = Number(attr(element, 'tint') ?? 0) || 0;
    const themeIndex = attr(element, 'theme');
    if (themeIndex !== undefined) return themeColorAt(theme, Number(themeIndex), tint);
    const indexed = attr(element, 'indexed');
    if (indexed !== undefined) {
      const index = Number(indexed);
      if (index === 64) return '#000000';
      if (index === 65) return '#ffffff';
      const base = customIndexed[index] ?? (INDEXED[index] ? `#${INDEXED[index]}` : undefined);
      return base ? applySheetTint(base, tint) : undefined;
    }
    return undefined;
  };

  const parseFont = (font: XmlElement | null): FontDef => {
    const result: FontDef = {};
    if (!font) return result;
    const bold = flagOn(child(font, 'b'));
    const italic = flagOn(child(font, 'i'));
    const strike = flagOn(child(font, 'strike'));
    if (bold !== undefined) result.bold = bold;
    if (italic !== undefined) result.italic = italic;
    if (strike !== undefined) result.strike = strike;
    const underline = child(font, 'u');
    if (underline) result.underline = attr(underline, 'val') !== 'none';
    const fontColor = color(child(font, 'color'));
    if (fontColor) result.color = fontColor;
    // Run properties call the font `rFont`; cell fonts call it `name`.
    const scheme = attr(child(font, 'scheme'), 'val');
    const name =
      (scheme === 'major' ? theme.majorFont : scheme === 'minor' ? theme.minorFont : undefined) ??
      safeFontName(attr(child(font, 'name') ?? child(font, 'rFont'), 'val'));
    if (name) result.name = name;
    const size = Number(attr(child(font, 'sz'), 'val'));
    if (size > 0 && size < 500) result.size = size;
    return result;
  };

  /** `differential` fills use `bgColor` as the solid colour and may clear the fill. */
  const parseFill = (fill: XmlElement | null, differential: boolean): string | null | undefined => {
    if (!fill) return undefined;
    const gradient = child(fill, 'gradientFill');
    if (gradient) {
      const stops = childElements(gradient, 'stop')
        .map((stop) => color(child(stop, 'color')))
        .filter((stop): stop is string => stop !== undefined);
      if (stops.length === 0) return undefined;
      if (stops.length === 1) return stops[0];
      const degree = Number(attr(gradient, 'degree') ?? 0) || 0;
      return `linear-gradient(${degree + 90}deg, ${stops.join(', ')})`;
    }
    const pattern = child(fill, 'patternFill');
    if (!pattern) return undefined;
    const type = attr(pattern, 'patternType');
    const foreground = color(child(pattern, 'fgColor'));
    const background = color(child(pattern, 'bgColor'));
    if (type === 'none') return differential ? null : undefined;
    if (type === 'solid' || (differential && type === undefined)) {
      return differential ? (background ?? foreground) : (foreground ?? background);
    }
    if (type === undefined) return undefined;
    return mix(foreground ?? '#000000', background ?? '#ffffff', PATTERN_DENSITY[type] ?? 0.3);
  };

  const parseSide = (
    side: XmlElement | null,
    differential: boolean,
  ): GridBorder | null | undefined => {
    if (!side) return undefined;
    const style = BORDER_STYLES[attr(side, 'style') ?? ''];
    if (!style) return differential ? null : undefined;
    return { width: style[0], style: style[1], color: color(child(side, 'color')) ?? '#000000' };
  };

  const parseBorder = (border: XmlElement | null, differential: boolean): BorderDef => {
    const result: BorderDef = {};
    if (!border) return result;
    const sides = [
      ['left', child(border, 'left') ?? child(border, 'start')],
      ['right', child(border, 'right') ?? child(border, 'end')],
      ['top', child(border, 'top')],
      ['bottom', child(border, 'bottom')],
      ['horizontal', child(border, 'horizontal')],
      ['vertical', child(border, 'vertical')],
    ] as const;
    for (const [name, element] of sides) {
      const side = parseSide(element, differential);
      if (side !== undefined) result[name] = side;
    }
    return result;
  };

  const parseAlignment = (alignment: XmlElement | null): AlignDef => {
    const result: AlignDef = {};
    if (!alignment) return result;
    const horizontal = attr(alignment, 'horizontal');
    const vertical = attr(alignment, 'vertical');
    if (horizontal) result.horizontal = horizontal;
    if (vertical) result.vertical = vertical;
    if (['1', 'true'].includes(attr(alignment, 'wrapText') ?? '')) result.wrap = true;
    const indent = Number(attr(alignment, 'indent'));
    if (indent > 0) result.indent = Math.min(indent, 50);
    return result;
  };

  const numberFormats = new Map<number, string>();
  for (const element of childElements(child(root, 'numFmts'), 'numFmt')) {
    const id = Number(attr(element, 'numFmtId'));
    const code = attr(element, 'formatCode');
    if (Number.isFinite(id) && code !== undefined) numberFormats.set(id, code);
  }
  const formatCode = (id: number): string =>
    numberFormats.get(id) ?? BUILTIN_FORMATS[id] ?? 'General';

  const fonts = childElements(child(root, 'fonts'), 'font').map(parseFont);
  const fills = childElements(child(root, 'fills'), 'fill').map(
    (fill) => parseFill(fill, false) ?? undefined,
  );
  const borders = childElements(child(root, 'borders'), 'border').map((border) =>
    parseBorder(border, false),
  );

  const xfElements = childElements(child(root, 'cellXfs'), 'xf');
  const baseXf = xfElements[0] ?? null;
  const id = (xf: XmlElement | null, name: string): number => Number(attr(xf, name) ?? 0) || 0;
  const applied = (xf: XmlElement, name: string): boolean =>
    ['1', 'true'].includes(attr(xf, name) ?? '');
  const cellXfs = xfElements.map((xf, index): CellXf => {
    const fillId = id(xf, 'fillId');
    const borderId = id(xf, 'borderId');
    return {
      numberFormat: formatCode(id(xf, 'numFmtId')),
      font: fonts[id(xf, 'fontId')] ?? {},
      fill: fills[fillId],
      border: borders[borderId] ?? {},
      align: parseAlignment(child(xf, 'alignment')),
      // Formatting equal to the default cell format does not count as set by the
      // author, so table styles and banding still show through it.
      explicitFont: index > 0 && applied(xf, 'applyFont'),
      explicitFill: index > 0 && fillId !== id(baseXf, 'fillId') && fills[fillId] !== undefined,
      explicitBorder: index > 0 && borderId !== id(baseXf, 'borderId'),
    };
  });

  const parseDxf = (dxf: XmlElement): Dxf => {
    const result: Dxf = {};
    const font = child(dxf, 'font');
    if (font) result.font = parseFont(font);
    const fill = parseFill(child(dxf, 'fill'), true);
    if (fill !== undefined) result.fill = fill;
    const border = child(dxf, 'border');
    if (border) result.border = parseBorder(border, true);
    const code = attr(child(dxf, 'numFmt'), 'formatCode');
    if (code !== undefined) result.numberFormat = code;
    return result;
  };
  const dxfs = childElements(child(root, 'dxfs'), 'dxf').map(parseDxf);

  const tableStylesRoot = child(root, 'tableStyles');
  const tableStyles = new Map<string, TableStyleDef>();
  for (const style of childElements(tableStylesRoot, 'tableStyle')) {
    const name = attr(style, 'name');
    if (!name) continue;
    const elements = new Map<string, { dxf: Dxf; size: number }>();
    for (const element of childElements(style, 'tableStyleElement')) {
      const dxf = dxfs[Number(attr(element, 'dxfId'))];
      const type = attr(element, 'type');
      if (dxf && type) {
        elements.set(type, { dxf, size: Math.max(1, Number(attr(element, 'size') ?? 1) || 1) });
      }
    }
    tableStyles.set(name, { elements });
  }

  return {
    cellXfs,
    dxfs,
    tableStyles,
    defaultTableStyle: attr(tableStylesRoot, 'defaultTableStyle'),
    defaultFont: fonts[0] ?? {},
    theme,
    color,
    font: parseFont,
  };
}

/** Parses the `x14:dxf` elements used by extended conditional formatting rules. */
export function parseInlineDxf(dxf: XmlElement, styles: StyleSheet): Dxf {
  const result: Dxf = {};
  const font = child(dxf, 'font');
  if (font) {
    const parsed: FontDef = {};
    const bold = flagOn(child(font, 'b'));
    const italic = flagOn(child(font, 'i'));
    const strike = flagOn(child(font, 'strike'));
    if (bold !== undefined) parsed.bold = bold;
    if (italic !== undefined) parsed.italic = italic;
    if (strike !== undefined) parsed.strike = strike;
    const fontColor = styles.color(child(font, 'color'));
    if (fontColor) parsed.color = fontColor;
    result.font = parsed;
  }
  const pattern = descendants(dxf, 'patternFill')[0];
  if (pattern) {
    const fill = styles.color(child(pattern, 'bgColor')) ?? styles.color(child(pattern, 'fgColor'));
    if (fill) result.fill = fill;
  }
  return result;
}

const thin = (color: string): GridBorder => ({ width: 1, style: 'solid', color });
const medium = (color: string): GridBorder => ({ width: 2, style: 'solid', color });

/**
 * Approximates Excel's built-in table styles (`TableStyleLight1`…`TableStyleDark11`),
 * which are not stored in workbooks. Each family is generated from a theme accent.
 */
export function builtinTableStyle(name: string, theme: OfficeTheme): TableStyleDef | undefined {
  const match = /^TableStyle(Light|Medium|Dark)(\d+)$/.exec(name);
  if (!match) return undefined;
  const family = match[1]!;
  const number = Number(match[2]);
  const accentIndex = (number - 1) % 7;
  const neutral = accentIndex === 0;
  const accent = neutral ? '#000000' : (themeColorAt(theme, accentIndex + 3) ?? '#4472c4');
  const light = (amount: number): string =>
    neutral
      ? applySheetTint('#ffffff', amount >= 0.75 ? -0.15 : amount >= 0.55 ? -0.25 : -0.35)
      : applySheetTint(accent, amount);
  const white = '#ffffff';
  const elements = new Map<string, { dxf: Dxf; size: number }>();
  const set = (type: string, dxf: Dxf): void => void elements.set(type, { dxf, size: 1 });
  const bold: FontDef = { bold: true };
  const group = Math.floor((number - 1) / 7);

  if (family === 'Light' && group === 0) {
    const text = neutral ? '#000000' : applySheetTint(accent, -0.25);
    set('wholeTable', {
      font: { color: text },
      border: { top: thin(accent), bottom: thin(accent) },
    });
    set('headerRow', { font: { ...bold, color: text }, border: { bottom: thin(accent) } });
    set('totalRow', { font: { ...bold, color: text }, border: { top: thin(accent) } });
    set('firstRowStripe', { fill: light(0.8) });
  } else if (family === 'Light' && group === 1) {
    const line = thin(accent);
    set('wholeTable', {
      border: { top: line, bottom: line, left: line, right: line, horizontal: line },
    });
    set('headerRow', { font: { ...bold, color: white }, fill: accent });
    set('totalRow', {
      font: bold,
      border: { top: { width: 3, style: 'double', color: accent } },
    });
  } else if (family === 'Light') {
    const line = thin(accent);
    set('wholeTable', {
      border: {
        top: line,
        bottom: line,
        left: line,
        right: line,
        horizontal: line,
        vertical: line,
      },
    });
    set('headerRow', { font: bold, border: { bottom: medium(accent) } });
    set('totalRow', {
      font: bold,
      border: { top: { width: 3, style: 'double', color: accent } },
    });
    set('firstRowStripe', { fill: light(0.8) });
  } else if (family === 'Medium' && group === 0) {
    const line = thin(light(0.4));
    set('wholeTable', {
      border: { top: line, bottom: line, left: line, right: line, horizontal: line },
    });
    set('headerRow', { font: { ...bold, color: white }, fill: accent });
    set('totalRow', {
      font: bold,
      border: { top: { width: 3, style: 'double', color: accent } },
    });
    set('firstRowStripe', { fill: light(0.8) });
  } else if (family === 'Medium' && group === 1) {
    const line = thin(white);
    set('wholeTable', { fill: light(0.8), border: { horizontal: line, vertical: line } });
    set('headerRow', {
      font: { ...bold, color: white },
      fill: accent,
      border: { bottom: { width: 3, style: 'solid', color: white } },
    });
    set('totalRow', {
      font: { ...bold, color: white },
      fill: accent,
      border: { top: { width: 3, style: 'solid', color: white } },
    });
    set('firstRowStripe', { fill: light(0.6) });
  } else if (family === 'Medium' && group === 2) {
    const line = medium('#000000');
    set('wholeTable', { border: { top: line, bottom: line } });
    set('headerRow', { font: { ...bold, color: white }, fill: accent, border: { bottom: line } });
    set('totalRow', {
      font: bold,
      border: { top: { width: 3, style: 'double', color: '#000000' } },
    });
    set('firstRowStripe', { fill: '#d9d9d9' });
  } else if (family === 'Medium') {
    const line = thin(light(0.4));
    set('wholeTable', {
      fill: light(0.8),
      border: {
        top: line,
        bottom: line,
        left: line,
        right: line,
        horizontal: line,
        vertical: line,
      },
    });
    set('headerRow', { font: bold });
    set('totalRow', { font: bold, border: { top: medium(accent) } });
    set('firstRowStripe', { fill: light(0.6) });
  } else {
    const body = neutral ? '#737373' : accent;
    const head = neutral ? '#000000' : applySheetTint(accent, -0.5);
    set('wholeTable', { font: { color: white }, fill: body });
    set('headerRow', {
      font: { ...bold, color: white },
      fill: head,
      border: { bottom: medium(white) },
    });
    set('totalRow', { font: { ...bold, color: white }, fill: head });
    set('firstRowStripe', { fill: neutral ? '#404040' : applySheetTint(accent, -0.25) });
  }
  set('firstColumn', { font: bold });
  set('lastColumn', { font: bold });
  return { elements };
}

/** Pixel size of a font given in points. */
export const pointsToPx = (points: number): number => (points * 96) / 72;

/** Merges font properties into a grid style. */
export function applyFont(style: GridCellStyle, font: FontDef | undefined): void {
  if (!font) return;
  if (font.bold !== undefined) style.bold = font.bold;
  if (font.italic !== undefined) style.italic = font.italic;
  if (font.underline !== undefined) style.underline = font.underline;
  if (font.strike !== undefined) style.strike = font.strike;
  if (font.color) style.color = font.color;
  if (font.name) style.fontFamily = font.name;
  if (font.size) style.fontSize = pointsToPx(font.size);
}
