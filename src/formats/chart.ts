import {
  attr,
  child,
  childElements,
  descendants,
  find,
  textOf,
  type XmlElement,
} from '../util/xml';
import { hexColor } from './ooxml';
import { safeFontName, type OfficeTheme } from './theme';

/**
 * DrawingML charts (`c:chartSpace`) → a small model that `ui/chart.ts` draws.
 * Values come from the caches stored in the file; nothing is recalculated.
 */

export type ChartKind = 'bar' | 'line' | 'area' | 'pie' | 'doughnut' | 'scatter';

export interface ChartSeries {
  name: string;
  categories: string[];
  values: (number | null)[];
  /** Numeric x values (scatter charts). */
  xValues?: (number | null)[];
  /** Fill colour, or the line colour for line and scatter charts. */
  color: string;
  /** Per-point colour overrides (data points and pie slices). */
  pointColors: (string | undefined)[];
  /** Outline of bars and slices. */
  stroke?: { color: string; width: number };
  /** Line width in CSS pixels (line and scatter charts). */
  lineWidth: number;
  markers: boolean;
  /** `false` for scatter series that only show markers. */
  connected: boolean;
  formatCode?: string;
}

export interface ChartGroup {
  kind: ChartKind;
  /** Bar charts only: bars run left to right. */
  horizontal: boolean;
  /** `clustered`, `stacked`, `percentStacked` or `standard`. */
  grouping: string;
  series: ChartSeries[];
  varyColors: boolean;
  /** Space between categories as a percentage of the bar width. */
  gapWidth: number;
  /** Doughnut hole as a percentage of the radius. */
  holeSize: number;
  showValues: boolean;
  /** Pie and doughnut charts: label each slice with its share. */
  showPercent: boolean;
}

export interface ChartText {
  /** Font size in CSS pixels. */
  size: number;
  color: string;
  bold: boolean;
  italic: boolean;
}

export interface ChartAxis {
  deleted: boolean;
  min?: number;
  max?: number;
  majorUnit?: number;
  formatCode?: string;
  /** Colour of major gridlines, when shown. */
  gridlines?: string;
  /** Colour of the axis line, when shown. */
  line?: string;
  ticks: boolean;
  text: ChartText;
}

export interface ChartModel {
  title?: string;
  titleText: ChartText;
  groups: ChartGroup[];
  legend?: 'b' | 't' | 'l' | 'r';
  legendText: ChartText;
  valueAxis: ChartAxis;
  categoryAxis: ChartAxis;
  background?: string;
  border?: { color: string; width: number };
  plotBackground?: string;
  plotBorder?: { color: string; width: number };
  fontFamily?: string;
}

const SCHEME_ALIASES: Readonly<Record<string, string>> = {
  tx1: 'dk1',
  tx2: 'dk2',
  bg1: 'lt1',
  bg2: 'lt2',
};
const SCHEME_FALLBACK: Readonly<Record<string, string>> = {
  dk1: '#000000',
  lt1: '#ffffff',
  dk2: '#44546a',
  lt2: '#e7e6e6',
  accent1: '#4472c4',
  accent2: '#ed7d31',
  accent3: '#a5a5a5',
  accent4: '#ffc000',
  accent5: '#5b9bd5',
  accent6: '#70ad47',
};
const PRESET_COLORS: Readonly<Record<string, string>> = {
  black: '#000000',
  white: '#ffffff',
  red: '#ff0000',
  green: '#008000',
  blue: '#0000ff',
  yellow: '#ffff00',
  gray: '#808080',
  grey: '#808080',
  orange: '#ffa500',
};

function toHsl(hex: string): [number, number, number] {
  const r = parseInt(hex.slice(1, 3), 16) / 255;
  const g = parseInt(hex.slice(3, 5), 16) / 255;
  const b = parseInt(hex.slice(5, 7), 16) / 255;
  const max = Math.max(r, g, b);
  const min = Math.min(r, g, b);
  const lightness = (max + min) / 2;
  if (max === min) return [0, 0, lightness];
  const delta = max - min;
  const saturation = lightness > 0.5 ? delta / (2 - max - min) : delta / (max + min);
  let hue: number;
  if (max === r) hue = (g - b) / delta + (g < b ? 6 : 0);
  else if (max === g) hue = (b - r) / delta + 2;
  else hue = (r - g) / delta + 4;
  return [hue / 6, saturation, lightness];
}

function fromHsl(hue: number, saturation: number, lightness: number): string {
  const q =
    lightness < 0.5
      ? lightness * (1 + saturation)
      : lightness + saturation - lightness * saturation;
  const p = 2 * lightness - q;
  const channel = (t: number): number => {
    const wrapped = t < 0 ? t + 1 : t > 1 ? t - 1 : t;
    if (wrapped < 1 / 6) return p + (q - p) * 6 * wrapped;
    if (wrapped < 1 / 2) return q;
    if (wrapped < 2 / 3) return p + (q - p) * (2 / 3 - wrapped) * 6;
    return p;
  };
  const values =
    saturation === 0
      ? [lightness, lightness, lightness]
      : [channel(hue + 1 / 3), channel(hue), channel(hue - 1 / 3)];
  return `#${values
    .map((value) =>
      Math.max(0, Math.min(255, Math.round(value * 255)))
        .toString(16)
        .padStart(2, '0'),
    )
    .join('')}`;
}

/** Resolves the colour element inside `parent` (`a:srgbClr`, `a:schemeClr`…) with its modifiers. */
export function drawingColor(parent: XmlElement | null, theme: OfficeTheme): string | undefined {
  for (const element of childElements(parent)) {
    let base: string | undefined;
    if (element.local === 'srgbClr') base = hexColor(attr(element, 'val'));
    else if (element.local === 'sysClr') {
      base = hexColor(attr(element, 'lastClr')) ?? '#000000';
    } else if (element.local === 'prstClr') base = PRESET_COLORS[attr(element, 'val') ?? ''];
    else if (element.local === 'schemeClr') {
      const raw = attr(element, 'val') ?? '';
      const name = SCHEME_ALIASES[raw] ?? raw;
      base = theme.colors.get(name) ?? SCHEME_FALLBACK[name];
    } else continue;
    if (!base) return undefined;
    const [hue, baseSaturation, baseLightness] = toHsl(base.toLowerCase());
    let saturation = baseSaturation;
    let lightness = baseLightness;
    for (const modifier of childElements(element)) {
      const amount = Number(attr(modifier, 'val') ?? 100000) / 100000;
      if (!Number.isFinite(amount)) continue;
      if (modifier.local === 'lumMod') lightness *= amount;
      else if (modifier.local === 'lumOff') lightness += amount;
      else if (modifier.local === 'tint') lightness = lightness * amount + (1 - amount);
      else if (modifier.local === 'shade') lightness *= amount;
      else if (modifier.local === 'satMod') saturation *= amount;
    }
    return fromHsl(hue, Math.max(0, Math.min(1, saturation)), Math.max(0, Math.min(1, lightness)));
  }
  return undefined;
}

const number = (value: string | undefined): number | undefined => {
  if (value === undefined || value === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
};
const flag = (element: XmlElement | null, fallback: boolean): boolean => {
  if (!element) return fallback;
  const value = attr(element, 'val');
  return value === undefined ? true : value !== '0' && value !== 'false';
};
/** EMU line width → CSS pixels. */
const lineWidth = (emu: number | undefined, fallback: number): number =>
  emu === undefined ? fallback : Math.max(0.5, Math.min((emu / 12700) * (96 / 72), 12));

function fillOf(shape: XmlElement | null, theme: OfficeTheme): string | undefined {
  if (!shape) return undefined;
  if (child(shape, 'noFill')) return 'none';
  const solid = child(shape, 'solidFill');
  if (solid) return drawingColor(solid, theme);
  // Gradients are approximated by their first stop.
  const stop = descendants(child(shape, 'gradFill'), 'gs')[0];
  return stop ? drawingColor(stop, theme) : undefined;
}

function lineOf(
  shape: XmlElement | null,
  theme: OfficeTheme,
): { color: string; width: number } | 'none' | undefined {
  const line = child(shape, 'ln');
  if (!line) return undefined;
  if (child(line, 'noFill')) return 'none';
  const color = drawingColor(child(line, 'solidFill'), theme);
  if (!color) return undefined;
  return { color, width: lineWidth(number(attr(line, 'w')), 1) };
}

function textOfProps(props: XmlElement | null, theme: OfficeTheme, fallback: ChartText): ChartText {
  const run = descendants(props, 'defRPr')[0] ?? descendants(props, 'rPr')[0];
  if (!run) return fallback;
  const size = number(attr(run, 'sz'));
  const bold = attr(run, 'b');
  const italic = attr(run, 'i');
  return {
    size: size ? Math.max(6, Math.min((size / 100) * (96 / 72), 96)) : fallback.size,
    color: drawingColor(child(run, 'solidFill'), theme) ?? fallback.color,
    bold: bold === undefined ? fallback.bold : bold === '1' || bold === 'true',
    italic: italic === undefined ? fallback.italic : italic === '1' || italic === 'true',
  };
}

/** Reads a `c:strCache`/`c:numCache` (or literal) into a dense list. */
function cachedPoints(reference: XmlElement | null): { values: string[]; formatCode?: string } {
  const cache =
    descendants(reference, 'numCache')[0] ??
    descendants(reference, 'strCache')[0] ??
    descendants(reference, 'numLit')[0] ??
    descendants(reference, 'strLit')[0] ??
    // Multi-level categories: the innermost level carries the labels.
    descendants(reference, 'lvl')[0];
  if (!cache) return { values: [] };
  const points = childElements(cache, 'pt');
  const count = Math.min(
    number(attr(child(cache, 'ptCount'), 'val')) ?? points.length,
    Math.max(points.length, 0) + 10_000,
  );
  const values: string[] = new Array<string>(count).fill('');
  points.forEach((point, order) => {
    const index = number(attr(point, 'idx')) ?? order;
    if (index >= 0 && index < 50_000) values[index] = textOf(child(point, 'v'));
  });
  const formatCode = textOf(child(cache, 'formatCode'));
  return formatCode ? { values, formatCode } : { values };
}

const toNumbers = (values: string[]): (number | null)[] =>
  values.map((value) => {
    if (value === '') return null;
    const parsed = Number(value);
    return Number.isFinite(parsed) ? parsed : null;
  });

const GROUP_KINDS: Readonly<Record<string, ChartKind>> = {
  barChart: 'bar',
  bar3DChart: 'bar',
  lineChart: 'line',
  line3DChart: 'line',
  areaChart: 'area',
  area3DChart: 'area',
  pieChart: 'pie',
  pie3DChart: 'pie',
  ofPieChart: 'pie',
  doughnutChart: 'doughnut',
  scatterChart: 'scatter',
  bubbleChart: 'scatter',
  radarChart: 'line',
  stockChart: 'line',
};

function parseAxis(axis: XmlElement | null, theme: OfficeTheme, base: ChartText): ChartAxis {
  const scaling = child(axis, 'scaling');
  const shape = child(axis, 'spPr');
  const line = lineOf(shape, theme);
  const gridShape = child(child(axis, 'majorGridlines'), 'spPr');
  const gridLine = lineOf(gridShape, theme);
  const result: ChartAxis = {
    deleted: flag(child(axis, 'delete'), false),
    ticks: !['none', undefined].includes(attr(child(axis, 'majorTickMark'), 'val')),
    text: textOfProps(child(axis, 'txPr'), theme, base),
  };
  const min = number(attr(child(scaling, 'min'), 'val'));
  const max = number(attr(child(scaling, 'max'), 'val'));
  const majorUnit = number(attr(child(axis, 'majorUnit'), 'val'));
  if (min !== undefined) result.min = min;
  if (max !== undefined) result.max = max;
  if (majorUnit !== undefined && majorUnit > 0) result.majorUnit = majorUnit;
  const format = child(axis, 'numFmt');
  const code = attr(format, 'formatCode');
  if (code && attr(format, 'sourceLinked') !== '1') result.formatCode = code;
  if (child(axis, 'majorGridlines')) {
    if (gridLine !== 'none') result.gridlines = gridLine?.color ?? '#d9d9d9';
  }
  if (line !== 'none') result.line = line?.color ?? '#bfbfbf';
  return result;
}

/** Parses a chart part. Returns `null` when it holds no chart type that can be drawn. */
export function parseChart(root: XmlElement | null, theme: OfficeTheme): ChartModel | null {
  const chart = child(root, 'chart');
  const plotArea = child(chart, 'plotArea');
  if (!plotArea) return null;

  const base: ChartText = { size: 13.3, color: '#595959', bold: false, italic: false };
  const spaceText = textOfProps(child(root, 'txPr'), theme, base);
  const palette = [1, 2, 3, 4, 5, 6].map((index) =>
    (theme.colors.get(`accent${index}`) ?? SCHEME_FALLBACK[`accent${index}`]!).toLowerCase(),
  );

  const groups: ChartGroup[] = [];
  let seriesCount = 0;
  for (const element of childElements(plotArea)) {
    const kind = GROUP_KINDS[element.local];
    if (!kind) continue;
    const varyColors = flag(child(element, 'varyColors'), false);
    const groupMarkers = flag(child(element, 'marker'), false);
    const scatterStyle = attr(child(element, 'scatterStyle'), 'val') ?? 'lineMarker';
    const series: ChartSeries[] = [];
    for (const ser of childElements(element, 'ser')) {
      const shape = child(ser, 'spPr');
      const fill = fillOf(shape, theme);
      const line = lineOf(shape, theme);
      const automatic = palette[seriesCount % palette.length]!;
      seriesCount++;
      const linear = kind === 'line' || kind === 'scatter';
      const categories = cachedPoints(child(ser, 'cat') ?? child(ser, 'xVal'));
      const values = cachedPoints(child(ser, 'val') ?? child(ser, 'yVal'));
      const pointColors: (string | undefined)[] = [];
      for (const point of childElements(ser, 'dPt')) {
        const index = number(attr(child(point, 'idx'), 'val'));
        const color = fillOf(child(point, 'spPr'), theme);
        if (index !== undefined && index >= 0 && index < 50_000 && color && color !== 'none') {
          pointColors[index] = color;
        }
      }
      const markerSymbol = attr(find(ser, 'marker', 'symbol'), 'val');
      const item: ChartSeries = {
        name: cachedPoints(child(ser, 'tx')).values[0] ?? textOf(find(ser, 'tx', 'v')),
        categories: categories.values,
        values: toNumbers(values.values),
        color: linear
          ? line && line !== 'none'
            ? line.color
            : automatic
          : fill && fill !== 'none'
            ? fill
            : automatic,
        pointColors,
        lineWidth: line && line !== 'none' ? line.width : 3,
        markers:
          kind === 'scatter' ? markerSymbol !== 'none' : groupMarkers && markerSymbol !== 'none',
        connected: !(
          kind === 'scatter' &&
          (line === 'none' || (scatterStyle === 'marker' && !line))
        ),
      };
      if (!item.name) item.name = `Series ${series.length + 1}`;
      if (kind === 'scatter') item.xValues = toNumbers(categories.values);
      if (!linear && line && line !== 'none') item.stroke = line;
      if (values.formatCode) item.formatCode = values.formatCode;
      series.push(item);
    }
    if (series.length === 0) continue;
    groups.push({
      kind,
      horizontal: kind === 'bar' && attr(child(element, 'barDir'), 'val') === 'bar',
      grouping: attr(child(element, 'grouping'), 'val') ?? 'clustered',
      series,
      varyColors,
      gapWidth: Math.max(0, Math.min(number(attr(child(element, 'gapWidth'), 'val')) ?? 150, 500)),
      holeSize: Math.max(10, Math.min(number(attr(child(element, 'holeSize'), 'val')) ?? 50, 90)),
      showValues: flag(find(element, 'dLbls', 'showVal'), false),
      showPercent: flag(find(element, 'dLbls', 'showPercent'), false),
    });
  }
  if (groups.length === 0) return null;

  const valueAxes = childElements(plotArea, 'valAx');
  const categoryElement =
    child(plotArea, 'catAx') ?? child(plotArea, 'dateAx') ?? child(plotArea, 'serAx');
  // Scatter charts have two value axes; the bottom one plays the category role.
  const scatter = groups[0]!.kind === 'scatter';
  const horizontalValue = (axis: XmlElement): boolean =>
    ['b', 't'].includes(attr(child(axis, 'axPos'), 'val') ?? '');
  const valueElement = scatter
    ? (valueAxes.find((axis) => !horizontalValue(axis)) ?? valueAxes[1] ?? null)
    : (valueAxes[0] ?? null);
  const xElement = scatter
    ? (valueAxes.find(horizontalValue) ?? valueAxes[0] ?? null)
    : categoryElement;

  const model: ChartModel = {
    titleText: { ...spaceText, size: Math.max(spaceText.size * 1.25, 16) },
    groups,
    legendText: textOfProps(find(chart, 'legend', 'txPr'), theme, spaceText),
    valueAxis: parseAxis(valueElement, theme, spaceText),
    categoryAxis: parseAxis(xElement, theme, spaceText),
  };

  const title = child(chart, 'title');
  if (title && !flag(child(chart, 'autoTitleDeleted'), false)) {
    const text = descendants(title, 't').map(textOf).join('');
    const only = groups.length === 1 && groups[0]!.series.length === 1;
    model.title = text || (only ? groups[0]!.series[0]!.name : '');
    model.titleText = textOfProps(
      child(title, 'txPr') ?? child(title, 'tx'),
      theme,
      model.titleText,
    );
    if (!model.title) delete model.title;
  }
  const legendPosition = attr(find(chart, 'legend', 'legendPos'), 'val');
  if (child(chart, 'legend')) {
    model.legend =
      legendPosition === 'b' || legendPosition === 't' || legendPosition === 'l'
        ? legendPosition
        : 'r';
  }

  const spaceShape = child(root, 'spPr');
  const background = fillOf(spaceShape, theme);
  if (background && background !== 'none') model.background = background;
  const border = lineOf(spaceShape, theme);
  if (border && border !== 'none') model.border = border;
  const plotShape = child(plotArea, 'spPr');
  const plotBackground = fillOf(plotShape, theme);
  if (plotBackground && plotBackground !== 'none') model.plotBackground = plotBackground;
  const plotBorder = lineOf(plotShape, theme);
  if (plotBorder && plotBorder !== 'none') model.plotBorder = plotBorder;
  const family = safeFontName(theme.minorFont);
  if (family) model.fontFamily = family;
  return model;
}
