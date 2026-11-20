import { attr, child, childElements, type XmlElement } from '../util/xml';
import { hexColor, isOn } from './ooxml';
import { safeFontName, shadeColor, tintColor, type OfficeTheme } from './theme';

/**
 * WordprocessingML property parsing: run, paragraph, table and cell properties
 * turned into plain objects that the renderer merges along the style chain.
 */

export interface Border {
  /** CSS border style; `none` for an explicit "no border". */
  style: string;
  /** Width in CSS pixels. */
  width: number;
  color: string;
  /** Distance from the text in CSS pixels (paragraph borders). */
  space: number;
}

export type Side = 'top' | 'left' | 'bottom' | 'right';
export type BorderSet = Partial<Record<Side | 'insideH' | 'insideV' | 'between', Border>>;

export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  /** CSS `text-decoration-style` for the underline, or `none`. */
  underline?: string;
  underlineColor?: string | undefined;
  strike?: boolean;
  doubleStrike?: boolean;
  caps?: boolean;
  smallCaps?: boolean;
  hidden?: boolean;
  color?: string | undefined;
  highlight?: string | undefined;
  /** Font size in points. */
  size?: number;
  font?: string;
  vertAlign?: string;
  /** Extra letter spacing in CSS pixels. */
  spacing?: number;
  /** Raised (positive) or lowered text, in points. */
  position?: number;
  border?: Border;
  shadow?: boolean;
}

export interface TabStop {
  /** Position from the left text margin, in CSS pixels. */
  pos: number;
  align: string;
  leader: string;
}

export interface Frame {
  dropCap: string | undefined;
  lines: number;
  hSpace: number;
  width: number | undefined;
  xAlign: string | undefined;
}

export interface ParaProps {
  align?: string;
  /** Indents and spacing in twips (1/20 pt). */
  indentLeft?: number;
  indentRight?: number;
  firstLine?: number;
  spaceBefore?: number;
  spaceAfter?: number;
  line?: number;
  lineRule?: string;
  numId?: string;
  level?: number;
  outlineLevel?: number;
  background?: string | undefined;
  pageBreakBefore?: boolean;
  keepNext?: boolean;
  keepLines?: boolean;
  widowControl?: boolean;
  contextualSpacing?: boolean;
  rtl?: boolean;
  tabs?: TabStop[];
  borders?: BorderSet;
  frame?: Frame;
}

export interface CellProps {
  background?: string | undefined;
  borders?: BorderSet;
  vAlign?: string;
  /** Cell padding in CSS pixels. */
  margins?: Partial<Record<Side, number>>;
  noWrap?: boolean;
  textDirection?: string;
}

export interface TableProps {
  borders?: BorderSet;
  margins?: Partial<Record<Side, number>>;
  /** Left indent in twips. */
  indent?: number;
  align?: string;
  rowBandSize?: number;
  colBandSize?: number;
}

export interface ConditionalFormat {
  para: ParaProps;
  run: RunProps;
  cell: CellProps;
}

export interface StyleDef {
  id: string;
  name: string;
  type: string;
  basedOn: string | undefined;
  para: ParaProps;
  run: RunProps;
  table: TableProps;
  cell: CellProps;
  conditionals: Map<string, ConditionalFormat>;
}

const HIGHLIGHTS: Readonly<Record<string, string>> = {
  yellow: '#ffff00',
  green: '#00ff00',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  blue: '#0000ff',
  red: '#ff0000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkYellow: '#808000',
  darkGray: '#808080',
  lightGray: '#c0c0c0',
  black: '#000000',
  white: '#ffffff',
};

const THEME_COLOR_NAMES: Readonly<Record<string, string>> = {
  dark1: 'dk1',
  light1: 'lt1',
  dark2: 'dk2',
  light2: 'lt2',
  text1: 'dk1',
  text2: 'dk2',
  background1: 'lt1',
  background2: 'lt2',
  hyperlink: 'hlink',
  followedHyperlink: 'folHlink',
};

const BORDER_STYLES: Readonly<Record<string, string>> = {
  single: 'solid',
  thick: 'solid',
  double: 'double',
  triple: 'double',
  dotted: 'dotted',
  dashed: 'dashed',
  dashSmallGap: 'dashed',
  dotDash: 'dashed',
  dotDotDash: 'dotted',
  wave: 'solid',
  doubleWave: 'double',
  inset: 'inset',
  outset: 'outset',
  threeDEmboss: 'ridge',
  threeDEngrave: 'groove',
};

const UNDERLINE_STYLES: Readonly<Record<string, string>> = {
  double: 'double',
  dotted: 'dotted',
  dottedHeavy: 'dotted',
  dash: 'dashed',
  dashedHeavy: 'dashed',
  dashLong: 'dashed',
  dashLongHeavy: 'dashed',
  dotDash: 'dashed',
  dashDotHeavy: 'dashed',
  dotDotDash: 'dotted',
  dashDotDotHeavy: 'dotted',
  wave: 'wavy',
  wavyHeavy: 'wavy',
  wavyDouble: 'wavy',
};

export const twipsToPx = (twips: number): number => twips / 15;

export function num(value: string | undefined): number | undefined {
  if (value === undefined || value === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
}

/** Later objects override earlier ones; used to resolve style inheritance. */
export function merge<T extends object>(...layers: (T | undefined)[]): T {
  const result = {} as T;
  for (const layer of layers) if (layer) Object.assign(result, layer);
  return result;
}

export class PropertyParser {
  constructor(private readonly theme: OfficeTheme) {}

  /**
   * Resolves a colour given as hex (`w:val`) or as a theme colour with an
   * optional tint or shade. `prefix` selects the attribute family: `""` for
   * `themeColor`/`themeTint`, `"Fill"` for `themeFill`/`themeFillTint`.
   */
  color(element: XmlElement | null, valueName = 'w:val', prefix = 'Color'): string | undefined {
    if (!element) return undefined;
    const themeName = attr(element, `w:theme${prefix}`);
    if (themeName) {
      const base = this.theme.colors.get(THEME_COLOR_NAMES[themeName] ?? themeName);
      if (base) {
        const family = prefix === 'Color' ? '' : prefix;
        const tint = parseInt(attr(element, `w:theme${family}Tint`) ?? '', 16);
        const shade = parseInt(attr(element, `w:theme${family}Shade`) ?? '', 16);
        if (Number.isFinite(tint)) return tintColor(base, tint / 255);
        if (Number.isFinite(shade)) return shadeColor(base, shade / 255);
        return base;
      }
    }
    return hexColor(attr(element, valueName));
  }

  /** Background of a `w:shd` element. */
  shading(element: XmlElement | null): string | undefined {
    if (!element) return undefined;
    const pattern = attr(element, 'w:val');
    if (pattern === 'solid') return this.color(element, 'w:color') ?? '#000000';
    return this.color(element, 'w:fill', 'Fill');
  }

  font(element: XmlElement | null): string | undefined {
    if (!element) return undefined;
    const themed = attr(element, 'w:asciiTheme') ?? attr(element, 'w:hAnsiTheme');
    if (themed) {
      const font = themed.startsWith('major') ? this.theme.majorFont : this.theme.minorFont;
      if (font) return font;
    }
    return safeFontName(attr(element, 'w:ascii') ?? attr(element, 'w:hAnsi'));
  }

  border(element: XmlElement | null): Border | undefined {
    if (!element) return undefined;
    const value = attr(element, 'w:val');
    if (value === undefined) return undefined;
    const space = ((num(attr(element, 'w:space')) ?? 0) * 4) / 3;
    if (value === 'nil' || value === 'none') {
      return { style: 'none', width: 0, color: 'transparent', space };
    }
    const eighths = num(attr(element, 'w:sz')) ?? 4;
    const style = BORDER_STYLES[value] ?? 'solid';
    // A CSS double border needs three pixels before both lines are visible.
    const width =
      style === 'double'
        ? Math.max(3, eighths / 2)
        : Math.max(0.75, Math.round((eighths / 6) * 4) / 4);
    return { style, width, color: this.color(element, 'w:color') ?? '#000000', space };
  }

  borders(element: XmlElement | null): BorderSet | undefined {
    if (!element) return undefined;
    const set: BorderSet = {};
    for (const side of childElements(element)) {
      const name = side.local === 'start' ? 'left' : side.local === 'end' ? 'right' : side.local;
      const border = this.border(side);
      if (
        border &&
        (name === 'top' ||
          name === 'left' ||
          name === 'bottom' ||
          name === 'right' ||
          name === 'insideH' ||
          name === 'insideV' ||
          name === 'between')
      ) {
        set[name] = border;
      }
    }
    return set;
  }

  run(rPr: XmlElement | null): RunProps {
    const props: RunProps = {};
    if (!rPr) return props;
    for (const element of childElements(rPr)) {
      switch (element.local) {
        case 'b':
          props.bold = isOn(element);
          break;
        case 'i':
          props.italic = isOn(element);
          break;
        case 'u': {
          const value = attr(element, 'w:val') ?? 'single';
          props.underline = value === 'none' ? 'none' : (UNDERLINE_STYLES[value] ?? 'solid');
          props.underlineColor = this.color(element, 'w:color');
          break;
        }
        case 'strike':
          props.strike = isOn(element);
          break;
        case 'dstrike':
          props.doubleStrike = isOn(element);
          break;
        case 'caps':
          props.caps = isOn(element);
          break;
        case 'smallCaps':
          props.smallCaps = isOn(element);
          break;
        case 'vanish':
          props.hidden = isOn(element);
          break;
        case 'shadow':
          props.shadow = isOn(element);
          break;
        case 'color':
          props.color = this.color(element);
          break;
        case 'highlight': {
          const value = attr(element, 'w:val') ?? '';
          props.highlight = value === 'none' ? undefined : HIGHLIGHTS[value];
          break;
        }
        case 'shd':
          props.highlight ??= this.shading(element);
          break;
        case 'sz': {
          const halfPoints = num(attr(element, 'w:val'));
          if (halfPoints) props.size = halfPoints / 2;
          break;
        }
        case 'rFonts': {
          const font = this.font(element);
          if (font) props.font = font;
          break;
        }
        case 'vertAlign':
          props.vertAlign = attr(element, 'w:val') ?? '';
          break;
        case 'spacing': {
          const twips = num(attr(element, 'w:val'));
          if (twips !== undefined) props.spacing = twipsToPx(twips);
          break;
        }
        case 'position': {
          const halfPoints = num(attr(element, 'w:val'));
          if (halfPoints !== undefined) props.position = halfPoints / 2;
          break;
        }
        case 'bdr': {
          const border = this.border(element);
          if (border) props.border = border;
          break;
        }
      }
    }
    return props;
  }

  para(pPr: XmlElement | null): ParaProps {
    const props: ParaProps = {};
    if (!pPr) return props;
    for (const element of childElements(pPr)) {
      switch (element.local) {
        case 'jc':
          props.align = attr(element, 'w:val') ?? '';
          break;
        case 'ind': {
          const left = num(attr(element, 'w:left') ?? attr(element, 'w:start'));
          const right = num(attr(element, 'w:right') ?? attr(element, 'w:end'));
          const hanging = num(attr(element, 'w:hanging'));
          const firstLine = num(attr(element, 'w:firstLine'));
          if (left !== undefined) props.indentLeft = left;
          if (right !== undefined) props.indentRight = right;
          if (hanging !== undefined) props.firstLine = -hanging;
          else if (firstLine !== undefined) props.firstLine = firstLine;
          break;
        }
        case 'spacing': {
          const automatic = (name: string): boolean =>
            attr(element, name) !== undefined && isOn(element, name);
          const before = automatic('w:beforeAutospacing') ? 280 : num(attr(element, 'w:before'));
          const after = automatic('w:afterAutospacing') ? 280 : num(attr(element, 'w:after'));
          const line = num(attr(element, 'w:line'));
          if (before !== undefined) props.spaceBefore = before;
          if (after !== undefined) props.spaceAfter = after;
          if (line !== undefined) {
            props.line = line;
            props.lineRule = attr(element, 'w:lineRule') ?? 'auto';
          }
          break;
        }
        case 'numPr': {
          const numId = attr(child(element, 'numId'), 'w:val');
          const level = num(attr(child(element, 'ilvl'), 'w:val'));
          if (numId !== undefined) props.numId = numId;
          if (level !== undefined) props.level = level;
          break;
        }
        case 'outlineLvl': {
          const level = num(attr(element, 'w:val'));
          if (level !== undefined) props.outlineLevel = level;
          break;
        }
        case 'shd':
          props.background = this.shading(element);
          break;
        case 'pageBreakBefore':
          props.pageBreakBefore = isOn(element);
          break;
        case 'keepNext':
          props.keepNext = isOn(element);
          break;
        case 'keepLines':
          props.keepLines = isOn(element);
          break;
        case 'widowControl':
          props.widowControl = isOn(element);
          break;
        case 'contextualSpacing':
          props.contextualSpacing = isOn(element);
          break;
        case 'bidi':
          props.rtl = isOn(element);
          break;
        case 'pBdr': {
          const borders = this.borders(element);
          if (borders) props.borders = borders;
          break;
        }
        case 'tabs':
          props.tabs = childElements(element, 'tab').flatMap((tab): TabStop[] => {
            const pos = num(attr(tab, 'w:pos'));
            if (pos === undefined) return [];
            return [
              {
                pos: twipsToPx(pos),
                align: attr(tab, 'w:val') ?? 'left',
                leader: attr(tab, 'w:leader') ?? 'none',
              },
            ];
          });
          break;
        case 'framePr':
          props.frame = {
            dropCap: attr(element, 'w:dropCap'),
            lines: num(attr(element, 'w:lines')) ?? 1,
            hSpace: twipsToPx(num(attr(element, 'w:hSpace')) ?? 0),
            width: num(attr(element, 'w:w')),
            xAlign: attr(element, 'w:xAlign'),
          };
          break;
      }
    }
    return props;
  }

  private margins(element: XmlElement | null): Partial<Record<Side, number>> | undefined {
    if (!element) return undefined;
    const margins: Partial<Record<Side, number>> = {};
    for (const side of childElements(element)) {
      const name = side.local === 'start' ? 'left' : side.local === 'end' ? 'right' : side.local;
      const width = num(attr(side, 'w:w'));
      if (
        width !== undefined &&
        (name === 'top' || name === 'left' || name === 'bottom' || name === 'right')
      ) {
        margins[name] = twipsToPx(width);
      }
    }
    return margins;
  }

  cell(tcPr: XmlElement | null): CellProps {
    const props: CellProps = {};
    if (!tcPr) return props;
    const shading = child(tcPr, 'shd');
    if (shading) props.background = this.shading(shading);
    const borders = this.borders(child(tcPr, 'tcBorders'));
    if (borders) props.borders = borders;
    const vAlign = attr(child(tcPr, 'vAlign'), 'w:val');
    if (vAlign) props.vAlign = vAlign;
    const margins = this.margins(child(tcPr, 'tcMar'));
    if (margins) props.margins = margins;
    const noWrap = child(tcPr, 'noWrap');
    if (noWrap) props.noWrap = isOn(noWrap);
    const direction = attr(child(tcPr, 'textDirection'), 'w:val');
    if (direction) props.textDirection = direction;
    return props;
  }

  table(tblPr: XmlElement | null): TableProps {
    const props: TableProps = {};
    if (!tblPr) return props;
    const borders = this.borders(child(tblPr, 'tblBorders'));
    if (borders) props.borders = borders;
    const margins = this.margins(child(tblPr, 'tblCellMar'));
    if (margins) props.margins = margins;
    const indent = child(tblPr, 'tblInd');
    if (indent && attr(indent, 'w:type') !== 'pct') {
      const width = num(attr(indent, 'w:w'));
      if (width !== undefined) props.indent = width;
    }
    const align = attr(child(tblPr, 'jc'), 'w:val');
    if (align) props.align = align;
    const rowBand = num(attr(child(tblPr, 'tblStyleRowBandSize'), 'w:val'));
    if (rowBand) props.rowBandSize = rowBand;
    const colBand = num(attr(child(tblPr, 'tblStyleColBandSize'), 'w:val'));
    if (colBand) props.colBandSize = colBand;
    return props;
  }

  style(element: XmlElement): StyleDef | null {
    const id = attr(element, 'w:styleId');
    if (!id) return null;
    const conditionals = new Map<string, ConditionalFormat>();
    for (const conditional of childElements(element, 'tblStylePr')) {
      const type = attr(conditional, 'w:type');
      if (!type) continue;
      conditionals.set(type, {
        para: this.para(child(conditional, 'pPr')),
        run: this.run(child(conditional, 'rPr')),
        cell: this.cell(child(conditional, 'tcPr')),
      });
    }
    return {
      id,
      name: attr(child(element, 'name'), 'w:val') ?? id,
      type: attr(element, 'w:type') ?? 'paragraph',
      basedOn: attr(child(element, 'basedOn'), 'w:val'),
      para: this.para(child(element, 'pPr')),
      run: this.run(child(element, 'rPr')),
      table: this.table(child(element, 'tblPr')),
      cell: this.cell(child(element, 'tcPr')),
      conditionals,
    };
  }
}

/** Tab stops along a style chain: later layers add stops or clear earlier ones. */
export function mergeTabs(layers: (TabStop[] | undefined)[]): TabStop[] {
  const stops = new Map<number, TabStop>();
  for (const layer of layers) {
    for (const stop of layer ?? []) {
      const key = Math.round(stop.pos * 10);
      if (stop.align === 'clear') stops.delete(key);
      else stops.set(key, stop);
    }
  }
  return [...stops.values()].sort((a, b) => a.pos - b.pos);
}

/**
 * Height of a single-spaced line as a multiple of the font size, per font
 * (ascent + descent + line gap). Word's "multiple" line spacing scales this.
 */
const LINE_FACTORS: Readonly<Record<string, number>> = {
  calibri: 1.2207,
  'calibri light': 1.2207,
  candara: 1.2207,
  constantia: 1.2207,
  corbel: 1.2207,
  cambria: 1.1724,
  'times new roman': 1.1499,
  arial: 1.1499,
  helvetica: 1.1499,
  'arial narrow': 1.1475,
  'arial black': 1.4102,
  'courier new': 1.1328,
  verdana: 1.2153,
  tahoma: 1.207,
  georgia: 1.1362,
  'trebuchet ms': 1.1611,
  'segoe ui': 1.3301,
  consolas: 1.1709,
  garamond: 1.125,
  'comic sans ms': 1.3936,
  impact: 1.2197,
  'century gothic': 1.2256,
  'book antiqua': 1.2598,
  'palatino linotype': 1.3496,
  aptos: 1.2207,
  ubuntu: 1.149,
  'ubuntu mono': 1,
};

export function lineFactor(font: string | undefined): number {
  return LINE_FACTORS[(font ?? '').toLowerCase()] ?? 1.2;
}
