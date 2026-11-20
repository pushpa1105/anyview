import { attr, child, childElements, type XmlElement } from '../util/xml';
import { hexColor } from './ooxml';

/**
 * DrawingML shared by Word, Excel and PowerPoint files: colours with their
 * modifiers, shape outlines (preset and custom geometry) and the SVG that
 * draws a shape's fill and line.
 */

const SVG_NS = 'http://www.w3.org/2000/svg';

const num = (value: string | undefined): number | undefined => {
  if (value === undefined || value === '') return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
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
  purple: '#800080',
  cyan: '#00ffff',
  magenta: '#ff00ff',
  lime: '#00ff00',
  navy: '#000080',
  teal: '#008080',
  maroon: '#800000',
  olive: '#808000',
  silver: '#c0c0c0',
  darkGray: '#a9a9a9',
  lightGray: '#d3d3d3',
  brown: '#a52a2a',
  pink: '#ffc0cb',
};

function clampByte(value: number): number {
  return Math.max(0, Math.min(255, Math.round(value)));
}

/** Applies DrawingML colour modifiers (luminance, tint, shade, alpha) to a hex colour. */
export function applyColorModifiers(hex: string, modifiers: readonly XmlElement[]): string {
  let red = parseInt(hex.slice(1, 3), 16);
  let green = parseInt(hex.slice(3, 5), 16);
  let blue = parseInt(hex.slice(5, 7), 16);
  let alpha = 1;
  for (const modifier of modifiers) {
    const amount = (num(attr(modifier, 'val')) ?? 100000) / 100000;
    switch (modifier.local) {
      case 'lumMod':
      case 'shade':
        red *= amount;
        green *= amount;
        blue *= amount;
        break;
      case 'lumOff':
        red += 255 * amount;
        green += 255 * amount;
        blue += 255 * amount;
        break;
      case 'tint':
        red += (255 - red) * (1 - amount);
        green += (255 - green) * (1 - amount);
        blue += (255 - blue) * (1 - amount);
        break;
      case 'alpha':
        alpha = amount;
        break;
    }
  }
  const [r, g, b] = [clampByte(red), clampByte(green), clampByte(blue)];
  if (alpha < 1) return `rgba(${r}, ${g}, ${b}, ${Math.max(0, alpha).toFixed(2)})`;
  return `#${[r, g, b].map((channel) => channel.toString(16).padStart(2, '0')).join('')}`;
}

/** Maps a scheme colour name (`accent1`, `tx1`, …) to a hex colour. */
export type SchemeColors = (name: string) => string | undefined;

/** Resolves the colour element inside `parent` (`srgbClr`, `schemeClr`, …) to a CSS colour. */
export function resolveColor(parent: XmlElement | null, scheme: SchemeColors): string | undefined {
  for (const element of childElements(parent)) {
    let hex: string | undefined;
    if (element.local === 'srgbClr') hex = hexColor(attr(element, 'val'));
    else if (element.local === 'sysClr') hex = hexColor(attr(element, 'lastClr'));
    else if (element.local === 'prstClr') hex = PRESET_COLORS[attr(element, 'val') ?? ''];
    else if (element.local === 'schemeClr') hex = scheme(attr(element, 'val') ?? '');
    else if (element.local === 'scrgbClr') {
      const channel = (name: string): string =>
        clampByte(((num(attr(element, name)) ?? 0) / 100000) * 255)
          .toString(16)
          .padStart(2, '0');
      hex = `#${channel('r')}${channel('g')}${channel('b')}`;
    } else continue;
    return hex ? applyColorModifiers(hex, childElements(element)) : undefined;
  }
  return undefined;
}

export interface GeometryPath {
  d: string;
  filled: boolean;
  stroked: boolean;
}

const fmt = (value: number): string => String(Math.round(value * 100) / 100);

// Evaluates the guide formulas of a custom geometry (multiply-divide, add-subtract, …).
function evaluateGuides(geometry: XmlElement, width: number, height: number): Map<string, number> {
  const smaller = Math.min(width, height);
  const values = new Map<string, number>([
    ['w', width],
    ['h', height],
    ['l', 0],
    ['t', 0],
    ['r', width],
    ['b', height],
    ['wd2', width / 2],
    ['hd2', height / 2],
    ['wd4', width / 4],
    ['hd4', height / 4],
    ['wd8', width / 8],
    ['hd8', height / 8],
    ['hc', width / 2],
    ['vc', height / 2],
    ['ss', smaller],
    ['ssd2', smaller / 2],
    ['ssd4', smaller / 4],
    ['ssd8', smaller / 8],
    ['ls', Math.max(width, height)],
    ['cd2', 10800000],
    ['cd4', 5400000],
    ['cd8', 2700000],
    ['3cd4', 16200000],
    ['3cd8', 8100000],
    ['5cd8', 13500000],
    ['7cd8', 18900000],
  ]);
  const value = (token: string | undefined): number => {
    if (token === undefined) return 0;
    const known = values.get(token);
    if (known !== undefined) return known;
    const parsed = Number(token);
    return Number.isFinite(parsed) ? parsed : 0;
  };
  const radians = (angle: number): number => (angle / 60000) * (Math.PI / 180);
  for (const list of [child(geometry, 'avLst'), child(geometry, 'gdLst')]) {
    for (const guide of childElements(list, 'gd')) {
      const name = attr(guide, 'name');
      const [operator, ...args] = (attr(guide, 'fmla') ?? '').trim().split(/\s+/);
      if (!name || !operator) continue;
      const [a, b, c] = args.map(value) as [number, number, number];
      let result = 0;
      switch (operator) {
        case '*/':
          result = c ? (a * b) / c : 0;
          break;
        case '+-':
          result = a + b - c;
          break;
        case '+/':
          result = c ? (a + b) / c : 0;
          break;
        case '?:':
          result = a > 0 ? b : c;
          break;
        case 'abs':
          result = Math.abs(a);
          break;
        case 'min':
          result = Math.min(a, b);
          break;
        case 'max':
          result = Math.max(a, b);
          break;
        case 'pin':
          result = Math.max(a, Math.min(b, c));
          break;
        case 'sqrt':
          result = Math.sqrt(Math.max(0, a));
          break;
        case 'mod':
          result = Math.sqrt(a * a + b * b + c * c);
          break;
        case 'sin':
          result = a * Math.sin(radians(b));
          break;
        case 'cos':
          result = a * Math.cos(radians(b));
          break;
        case 'tan':
          result = a * Math.tan(radians(b));
          break;
        case 'at2':
          result = (Math.atan2(b, a) * 180 * 60000) / Math.PI;
          break;
        case 'cat2':
          result = a * Math.cos(Math.atan2(c, b));
          break;
        case 'sat2':
          result = a * Math.sin(Math.atan2(c, b));
          break;
        case 'val':
          result = a;
          break;
      }
      values.set(name, Number.isFinite(result) ? result : 0);
    }
  }
  return values;
}

/** Outline of an `a:custGeom`, scaled to `width` × `height` pixels. */
function customGeometry(geometry: XmlElement, width: number, height: number): GeometryPath[] {
  const paths: GeometryPath[] = [];
  for (const path of childElements(child(geometry, 'pathLst'), 'path')) {
    // Each path has its own coordinate space; without one, coordinates are in shape units.
    const spaceWidth = num(attr(path, 'w')) ?? 0;
    const spaceHeight = num(attr(path, 'h')) ?? 0;
    const guides = evaluateGuides(geometry, spaceWidth || width, spaceHeight || height);
    const scaleX = spaceWidth ? width / spaceWidth : 1;
    const scaleY = spaceHeight ? height / spaceHeight : 1;
    const coordinate = (token: string | undefined): number => {
      if (token === undefined) return 0;
      const known = guides.get(token);
      if (known !== undefined) return known;
      const parsed = Number(token);
      return Number.isFinite(parsed) ? parsed : 0;
    };
    const point = (element: XmlElement): [number, number] => [
      coordinate(attr(element, 'x')) * scaleX,
      coordinate(attr(element, 'y')) * scaleY,
    ];
    let d = '';
    let current: [number, number] = [0, 0];
    for (const command of childElements(path)) {
      const points = childElements(command, 'pt').map(point);
      switch (command.local) {
        case 'moveTo':
          if (points[0]) {
            d += `M${fmt(points[0][0])} ${fmt(points[0][1])}`;
            current = points[0];
          }
          break;
        case 'lnTo':
          if (points[0]) {
            d += `L${fmt(points[0][0])} ${fmt(points[0][1])}`;
            current = points[0];
          }
          break;
        case 'cubicBezTo':
          if (points.length === 3) {
            d += `C${points.map(([x, y]) => `${fmt(x)} ${fmt(y)}`).join(' ')}`;
            current = points[2]!;
          }
          break;
        case 'quadBezTo':
          if (points.length === 2) {
            d += `Q${points.map(([x, y]) => `${fmt(x)} ${fmt(y)}`).join(' ')}`;
            current = points[1]!;
          }
          break;
        case 'arcTo': {
          const radiusX = coordinate(attr(command, 'wR')) * scaleX;
          const radiusY = coordinate(attr(command, 'hR')) * scaleY;
          const start = (coordinate(attr(command, 'stAng')) / 60000) * (Math.PI / 180);
          const sweep = (coordinate(attr(command, 'swAng')) / 60000) * (Math.PI / 180);
          const centreX = current[0] - radiusX * Math.cos(start);
          const centreY = current[1] - radiusY * Math.sin(start);
          // A full turn cannot be expressed as one SVG arc, so sweeps are drawn in halves.
          const steps = Math.abs(sweep) > Math.PI ? 2 : 1;
          for (let step = 1; step <= steps; step++) {
            const angle = start + (sweep * step) / steps;
            const x = centreX + radiusX * Math.cos(angle);
            const y = centreY + radiusY * Math.sin(angle);
            d += `A${fmt(radiusX)} ${fmt(radiusY)} 0 0 ${sweep > 0 ? 1 : 0} ${fmt(x)} ${fmt(y)}`;
            current = [x, y];
          }
          break;
        }
        case 'close':
          d += 'Z';
          break;
      }
    }
    if (d) {
      paths.push({
        d,
        filled: attr(path, 'fill') !== 'none',
        stroked: attr(path, 'stroke') !== '0',
      });
    }
  }
  return paths;
}

/** Stretches a set of points so that it exactly fills the shape's box, as Office does. */
function fitted(points: [number, number][], width: number, height: number): [number, number][] {
  const xs = points.map(([x]) => x);
  const ys = points.map(([, y]) => y);
  const [left, top] = [Math.min(...xs), Math.min(...ys)];
  const spanX = Math.max(...xs) - left || 1;
  const spanY = Math.max(...ys) - top || 1;
  return points.map(([x, y]) => [((x - left) / spanX) * width, ((y - top) / spanY) * height]);
}

function polygon(points: [number, number][]): string {
  return `M${points.map(([x, y]) => `${fmt(x)} ${fmt(y)}`).join('L')}Z`;
}

function star(points: number, inner: number, width: number, height: number): string {
  const vertices: [number, number][] = [];
  for (let i = 0; i < points * 2; i++) {
    const radius = i % 2 === 0 ? 1 : inner;
    const angle = (Math.PI * i) / points - Math.PI / 2;
    vertices.push([
      (width / 2) * (1 + radius * Math.cos(angle)),
      (height / 2) * (1 + radius * Math.sin(angle)),
    ]);
  }
  return polygon(fitted(vertices, width, height));
}

function regular(sides: number, width: number, height: number): string {
  const vertices: [number, number][] = [];
  for (let i = 0; i < sides; i++) {
    const angle = (2 * Math.PI * i) / sides - Math.PI / 2;
    vertices.push([(width / 2) * (1 + Math.cos(angle)), (height / 2) * (1 + Math.sin(angle))]);
  }
  return polygon(fitted(vertices, width, height));
}

function roundedRect(
  width: number,
  height: number,
  radii: [number, number, number, number],
): string {
  const [tl, tr, br, bl] = radii;
  return (
    `M${fmt(tl)} 0L${fmt(width - tr)} 0` +
    (tr ? `A${fmt(tr)} ${fmt(tr)} 0 0 1 ${fmt(width)} ${fmt(tr)}` : '') +
    `L${fmt(width)} ${fmt(height - br)}` +
    (br ? `A${fmt(br)} ${fmt(br)} 0 0 1 ${fmt(width - br)} ${fmt(height)}` : '') +
    `L${fmt(bl)} ${fmt(height)}` +
    (bl ? `A${fmt(bl)} ${fmt(bl)} 0 0 1 0 ${fmt(height - bl)}` : '') +
    `L0 ${fmt(tl)}` +
    (tl ? `A${fmt(tl)} ${fmt(tl)} 0 0 1 ${fmt(tl)} 0` : '') +
    'Z'
  );
}

function ellipse(width: number, height: number, inset = 0): string {
  const rx = width / 2 - inset;
  const ry = height / 2 - inset;
  return (
    `M${fmt(inset)} ${fmt(height / 2)}` +
    `A${fmt(rx)} ${fmt(ry)} 0 1 1 ${fmt(width - inset)} ${fmt(height / 2)}` +
    `A${fmt(rx)} ${fmt(ry)} 0 1 1 ${fmt(inset)} ${fmt(height / 2)}Z`
  );
}

/** Preset shapes that are drawn as an open line rather than a filled area. */
export function isLineGeometry(preset: string): boolean {
  return (
    preset === 'line' ||
    /Connector/.test(preset) ||
    /^(arc|bracePair|bracketPair|leftBrace|rightBrace|leftBracket|rightBracket)$/.test(preset)
  );
}

/**
 * Outline of a preset shape. `adjust` returns the shape's adjustment values
 * (`adj`, `adj1`, …) when the file overrides the defaults.
 */
function presetGeometry(
  preset: string,
  width: number,
  height: number,
  adjust: (name: string, fallback: number) => number,
): GeometryPath[] {
  const w = width;
  const h = height;
  const ss = Math.min(w, h);
  const filled = (d: string): GeometryPath[] => [{ d, filled: true, stroked: true }];
  const open = (d: string): GeometryPath[] => [{ d, filled: false, stroked: true }];
  const part = (name: string, fallback: number, of = ss): number =>
    (of * adjust(name, fallback)) / 100000;

  switch (preset) {
    case 'line':
    case 'straightConnector1':
      return open(`M0 0L${fmt(w)} ${fmt(h)}`);
    case 'bentConnector2':
      return open(`M0 0L${fmt(w)} 0L${fmt(w)} ${fmt(h)}`);
    case 'bentConnector3':
    case 'bentConnector4':
    case 'bentConnector5': {
      const x = part('adj1', 50000, w);
      return open(`M0 0L${fmt(x)} 0L${fmt(x)} ${fmt(h)}L${fmt(w)} ${fmt(h)}`);
    }
    case 'curvedConnector2':
      return open(`M0 0Q${fmt(w)} 0 ${fmt(w)} ${fmt(h)}`);
    case 'curvedConnector3':
    case 'curvedConnector4':
    case 'curvedConnector5': {
      const x = part('adj1', 50000, w);
      return open(`M0 0C${fmt(x)} 0 ${fmt(x)} ${fmt(h)} ${fmt(w)} ${fmt(h)}`);
    }
    case 'ellipse':
    case 'flowChartConnector':
    case 'smileyFace':
    case 'cloud':
    case 'cloudCallout':
    case 'wedgeEllipseCallout':
    case 'pie':
    case 'chord':
      return filled(ellipse(w, h));
    case 'arc':
      return open(`M${fmt(w / 2)} 0A${fmt(w / 2)} ${fmt(h / 2)} 0 0 1 ${fmt(w)} ${fmt(h / 2)}`);
    case 'donut': {
      const inset = part('adj', 25000);
      return [{ d: ellipse(w, h) + ellipse(w, h, inset), filled: true, stroked: true }];
    }
    case 'roundRect':
    case 'flowChartAlternateProcess':
    case 'wedgeRoundRectCallout': {
      const r = Math.min(part('adj', 16667), ss / 2);
      return filled(roundedRect(w, h, [r, r, r, r]));
    }
    case 'flowChartTerminator': {
      const r = ss / 2;
      return filled(roundedRect(w, h, [r, r, r, r]));
    }
    case 'round1Rect': {
      const r = Math.min(part('adj', 16667), ss / 2);
      return filled(roundedRect(w, h, [0, r, 0, 0]));
    }
    case 'round2SameRect': {
      const top = Math.min(part('adj1', 16667), ss / 2);
      const bottom = Math.min(part('adj2', 0), ss / 2);
      return filled(roundedRect(w, h, [top, top, bottom, bottom]));
    }
    case 'round2DiagRect': {
      const a = Math.min(part('adj1', 16667), ss / 2);
      const b = Math.min(part('adj2', 0), ss / 2);
      return filled(roundedRect(w, h, [a, b, a, b]));
    }
    case 'snip1Rect': {
      const x = Math.min(part('adj', 16667), ss / 2);
      return filled(
        polygon([
          [0, 0],
          [w - x, 0],
          [w, x],
          [w, h],
          [0, h],
        ]),
      );
    }
    case 'snip2SameRect': {
      const x = Math.min(part('adj1', 16667), ss / 2);
      return filled(
        polygon([
          [x, 0],
          [w - x, 0],
          [w, x],
          [w, h],
          [0, h],
          [0, x],
        ]),
      );
    }
    case 'snip2DiagRect': {
      const x = Math.min(part('adj2', 16667), ss / 2);
      return filled(
        polygon([
          [0, 0],
          [w - x, 0],
          [w, x],
          [w, h],
          [x, h],
          [0, h - x],
        ]),
      );
    }
    case 'triangle':
    case 'flowChartExtract': {
      const x = part('adj', 50000, w);
      return filled(
        polygon([
          [0, h],
          [x, 0],
          [w, h],
        ]),
      );
    }
    case 'flowChartMerge':
      return filled(
        polygon([
          [0, 0],
          [w, 0],
          [w / 2, h],
        ]),
      );
    case 'rtTriangle':
      return filled(
        polygon([
          [0, h],
          [0, 0],
          [w, h],
        ]),
      );
    case 'diamond':
    case 'flowChartDecision':
      return filled(
        polygon([
          [w / 2, 0],
          [w, h / 2],
          [w / 2, h],
          [0, h / 2],
        ]),
      );
    case 'parallelogram':
    case 'flowChartInputOutput': {
      const x = Math.min(part('adj', 25000), w);
      return filled(
        polygon([
          [0, h],
          [x, 0],
          [w, 0],
          [w - x, h],
        ]),
      );
    }
    case 'trapezoid': {
      const x = Math.min(part('adj', 25000), w / 2);
      return filled(
        polygon([
          [0, h],
          [x, 0],
          [w - x, 0],
          [w, h],
        ]),
      );
    }
    case 'flowChartManualOperation':
      return filled(
        polygon([
          [0, 0],
          [w, 0],
          [w * 0.8, h],
          [w * 0.2, h],
        ]),
      );
    case 'pentagon':
      return filled(regular(5, w, h));
    case 'hexagon':
    case 'flowChartPreparation': {
      const x = Math.min(part('adj', 25000), w / 2);
      return filled(
        polygon([
          [0, h / 2],
          [x, 0],
          [w - x, 0],
          [w, h / 2],
          [w - x, h],
          [x, h],
        ]),
      );
    }
    case 'heptagon':
      return filled(regular(7, w, h));
    case 'octagon': {
      const x = Math.min(part('adj', 29289), ss / 2);
      return filled(
        polygon([
          [x, 0],
          [w - x, 0],
          [w, x],
          [w, h - x],
          [w - x, h],
          [x, h],
          [0, h - x],
          [0, x],
        ]),
      );
    }
    case 'decagon':
      return filled(regular(10, w, h));
    case 'dodecagon':
      return filled(regular(12, w, h));
    case 'plus':
    case 'mathPlus': {
      const x = Math.min(part('adj', 25000), ss / 2);
      return filled(
        polygon([
          [x, 0],
          [w - x, 0],
          [w - x, x],
          [w, x],
          [w, h - x],
          [w - x, h - x],
          [w - x, h],
          [x, h],
          [x, h - x],
          [0, h - x],
          [0, x],
          [x, x],
        ]),
      );
    }
    case 'rightArrow':
    case 'leftArrow':
    case 'upArrow':
    case 'downArrow': {
      const horizontal = preset === 'rightArrow' || preset === 'leftArrow';
      const across = horizontal ? h : w;
      const along = horizontal ? w : h;
      const shaft = (across * adjust('adj1', 50000)) / 200000;
      const head = Math.min(part('adj2', 50000), along);
      // Drawn pointing right, then mapped to the requested direction.
      const points: [number, number][] = [
        [0, across / 2 - shaft],
        [along - head, across / 2 - shaft],
        [along - head, 0],
        [along, across / 2],
        [along - head, across],
        [along - head, across / 2 + shaft],
        [0, across / 2 + shaft],
      ];
      const mapped = points.map(([x, y]): [number, number] => {
        if (preset === 'rightArrow') return [x, y];
        if (preset === 'leftArrow') return [w - x, y];
        if (preset === 'downArrow') return [y, x];
        return [y, h - x];
      });
      return filled(polygon(mapped));
    }
    case 'leftRightArrow': {
      const shaft = (h * adjust('adj1', 50000)) / 200000;
      const head = Math.min(part('adj2', 50000), w / 2);
      return filled(
        polygon([
          [0, h / 2],
          [head, 0],
          [head, h / 2 - shaft],
          [w - head, h / 2 - shaft],
          [w - head, 0],
          [w, h / 2],
          [w - head, h],
          [w - head, h / 2 + shaft],
          [head, h / 2 + shaft],
          [head, h],
        ]),
      );
    }
    case 'upDownArrow': {
      const shaft = (w * adjust('adj1', 50000)) / 200000;
      const head = Math.min(part('adj2', 50000), h / 2);
      return filled(
        polygon([
          [w / 2, 0],
          [w, head],
          [w / 2 + shaft, head],
          [w / 2 + shaft, h - head],
          [w, h - head],
          [w / 2, h],
          [0, h - head],
          [w / 2 - shaft, h - head],
          [w / 2 - shaft, head],
          [0, head],
        ]),
      );
    }
    case 'chevron': {
      const x = Math.min(part('adj', 50000), w);
      return filled(
        polygon([
          [0, 0],
          [w - x, 0],
          [w, h / 2],
          [w - x, h],
          [0, h],
          [x, h / 2],
        ]),
      );
    }
    case 'homePlate': {
      const x = Math.min(part('adj', 50000), w);
      return filled(
        polygon([
          [0, 0],
          [w - x, 0],
          [w, h / 2],
          [w - x, h],
          [0, h],
        ]),
      );
    }
    case 'star4':
      return filled(star(4, adjust('adj', 12500) / 50000, w, h));
    case 'star5':
      return filled(star(5, 0.382, w, h));
    case 'star6':
      return filled(star(6, 0.577, w, h));
    case 'star7':
      return filled(star(7, 0.69, w, h));
    case 'star8':
      return filled(star(8, 0.75, w, h));
    case 'star10':
      return filled(star(10, 0.85, w, h));
    case 'star12':
      return filled(star(12, 0.75, w, h));
    case 'star16':
    case 'star24':
    case 'star32':
      return filled(star(Number(preset.slice(4)), 0.75, w, h));
    case 'frame': {
      const x = Math.min(part('adj1', 12500), ss / 2);
      return [
        {
          d:
            polygon([
              [0, 0],
              [w, 0],
              [w, h],
              [0, h],
            ]) +
            polygon([
              [x, x],
              [x, h - x],
              [w - x, h - x],
              [w - x, x],
            ]),
          filled: true,
          stroked: true,
        },
      ];
    }
    case 'heart': {
      const outer = (w * 49) / 48;
      const inner = (w * 10) / 48;
      const top = -h / 3;
      return filled(
        `M${fmt(w / 2)} ${fmt(h / 4)}` +
          `C${fmt(w / 2 + inner)} ${fmt(top)} ${fmt(w / 2 + outer)} ${fmt(h / 4)} ${fmt(w / 2)} ${fmt(h)}` +
          `C${fmt(w / 2 - outer)} ${fmt(h / 4)} ${fmt(w / 2 - inner)} ${fmt(top)} ${fmt(w / 2)} ${fmt(h / 4)}Z`,
      );
    }
    case 'can':
    case 'flowChartMagneticDisk': {
      const y = Math.min(part('adj', 25000), h / 2) / 2;
      return [
        {
          d:
            `M0 ${fmt(y)}A${fmt(w / 2)} ${fmt(y)} 0 0 1 ${fmt(w)} ${fmt(y)}L${fmt(w)} ${fmt(h - y)}` +
            `A${fmt(w / 2)} ${fmt(y)} 0 0 1 0 ${fmt(h - y)}Z`,
          filled: true,
          stroked: true,
        },
        {
          d: `M0 ${fmt(y)}A${fmt(w / 2)} ${fmt(y)} 0 0 0 ${fmt(w)} ${fmt(y)}`,
          filled: false,
          stroked: true,
        },
      ];
    }
    case 'bracketPair': {
      const r = Math.min(part('adj', 16667), ss / 2);
      return open(
        `M${fmt(r)} 0A${fmt(r)} ${fmt(r)} 0 0 0 0 ${fmt(r)}L0 ${fmt(h - r)}A${fmt(r)} ${fmt(r)} 0 0 0 ${fmt(r)} ${fmt(h)}` +
          `M${fmt(w - r)} 0A${fmt(r)} ${fmt(r)} 0 0 1 ${fmt(w)} ${fmt(r)}L${fmt(w)} ${fmt(h - r)}A${fmt(r)} ${fmt(r)} 0 0 1 ${fmt(w - r)} ${fmt(h)}`,
      );
    }
    case 'leftBracket':
      return open(`M${fmt(w)} 0L0 0L0 ${fmt(h)}L${fmt(w)} ${fmt(h)}`);
    case 'rightBracket':
      return open(`M0 0L${fmt(w)} 0L${fmt(w)} ${fmt(h)}L0 ${fmt(h)}`);
    case 'bracePair':
    case 'leftBrace':
    case 'rightBrace':
      return open(`M${fmt(w)} 0L0 0L0 ${fmt(h)}L${fmt(w)} ${fmt(h)}`);
    default:
      return filled(
        polygon([
          [0, 0],
          [w, 0],
          [w, h],
          [0, h],
        ]),
      );
  }
}

/** Outline of a shape from its `spPr`, in a `width` × `height` pixel box. */
export function shapeGeometry(
  properties: XmlElement | null,
  width: number,
  height: number,
): GeometryPath[] {
  const custom = child(properties, 'custGeom');
  if (custom) {
    const paths = customGeometry(custom, width, height);
    if (paths.length > 0) return paths;
  }
  const presetElement = child(properties, 'prstGeom');
  const preset = attr(presetElement, 'prst') ?? 'rect';
  const adjustments = new Map<string, number>();
  for (const guide of childElements(child(presetElement, 'avLst'), 'gd')) {
    const match = /^val\s+(-?\d+)/.exec(attr(guide, 'fmla') ?? '');
    const name = attr(guide, 'name');
    if (match && name) adjustments.set(name, Number(match[1]));
  }
  return presetGeometry(
    preset,
    width,
    height,
    (name, fallback) => adjustments.get(name) ?? fallback,
  );
}

export interface GradientFill {
  /** Direction in degrees, clockwise from left-to-right. */
  angle: number;
  stops: { position: number; color: string }[];
}

export interface ShapePaint {
  /** CSS colour, gradient, picture URL promise, or `null` for no fill. */
  fill: string | GradientFill | { image: Promise<string | null> } | null;
  /** CSS colour, or `null` for no line. */
  stroke: string | null;
  strokeWidth: number;
  /** `a:prstDash` value. */
  dash?: string | undefined;
  /** Arrow heads on open paths. */
  headEnd?: boolean;
  tailEnd?: boolean;
}

/** Parses an `a:gradFill` element. */
export function parseGradient(element: XmlElement, scheme: SchemeColors): GradientFill | null {
  const stops = childElements(child(element, 'gsLst'), 'gs')
    .map((stop) => ({
      position: (num(attr(stop, 'pos')) ?? 0) / 1000,
      color: resolveColor(stop, scheme),
    }))
    .filter((stop): stop is { position: number; color: string } => stop.color !== undefined)
    .sort((a, b) => a.position - b.position);
  if (stops.length === 0) return null;
  return { angle: (num(attr(child(element, 'lin'), 'ang')) ?? 5400000) / 60000, stops };
}

let sequence = 0;

const DASHES: Readonly<Record<string, number[]>> = {
  dot: [1, 3],
  sysDot: [1, 1],
  dash: [4, 3],
  sysDash: [3, 1],
  lgDash: [8, 3],
  dashDot: [4, 3, 1, 3],
  sysDashDot: [3, 1, 1, 1],
  lgDashDot: [8, 3, 1, 3],
  lgDashDotDot: [8, 3, 1, 3, 1, 3],
  sysDashDotDot: [3, 1, 1, 1, 1, 1],
};

/** Draws a shape outline as an SVG that fills its (positioned) parent. */
export function buildShapeSvg(
  paths: readonly GeometryPath[],
  width: number,
  height: number,
  paint: ShapePaint,
): SVGSVGElement {
  const create = <K extends keyof SVGElementTagNameMap>(tag: K): SVGElementTagNameMap[K] =>
    document.createElementNS(SVG_NS, tag);
  const svg = create('svg');
  // A horizontal or vertical line has an empty box; a zero-sized SVG would not be drawn at all.
  svg.setAttribute('viewBox', `0 0 ${fmt(Math.max(width, 1))} ${fmt(Math.max(height, 1))}`);
  svg.setAttribute('preserveAspectRatio', 'none');
  svg.setAttribute('aria-hidden', 'true');
  svg.setAttribute('class', 'av-shape-svg');
  svg.style.cssText = `position:absolute;left:0;top:0;width:${width < 1 ? '1px' : '100%'};height:${height < 1 ? '1px' : '100%'};overflow:visible;pointer-events:none`;

  let fillValue = 'none';
  const { fill } = paint;
  if (typeof fill === 'string') {
    fillValue = fill;
  } else if (fill && 'stops' in fill) {
    const id = `av-grad-${++sequence}`;
    const gradient = create('linearGradient');
    gradient.setAttribute('id', id);
    const radians = (fill.angle * Math.PI) / 180;
    const dx = Math.cos(radians) / 2;
    const dy = Math.sin(radians) / 2;
    gradient.setAttribute('x1', fmt(0.5 - dx));
    gradient.setAttribute('y1', fmt(0.5 - dy));
    gradient.setAttribute('x2', fmt(0.5 + dx));
    gradient.setAttribute('y2', fmt(0.5 + dy));
    for (const stop of fill.stops) {
      const node = create('stop');
      node.setAttribute('offset', `${fmt(stop.position)}%`);
      node.setAttribute('stop-color', stop.color);
      gradient.appendChild(node);
    }
    const defs = create('defs');
    defs.appendChild(gradient);
    svg.appendChild(defs);
    fillValue = `url(#${id})`;
  } else if (fill && 'image' in fill) {
    // The picture is clipped to the outline of the shape.
    const id = `av-clip-${++sequence}`;
    const clip = create('clipPath');
    clip.setAttribute('id', id);
    for (const path of paths) {
      if (!path.filled) continue;
      const outline = create('path');
      outline.setAttribute('d', path.d);
      clip.appendChild(outline);
    }
    const defs = create('defs');
    defs.appendChild(clip);
    svg.appendChild(defs);
    const image = create('image');
    image.setAttribute('width', fmt(width));
    image.setAttribute('height', fmt(height));
    image.setAttribute('preserveAspectRatio', 'none');
    image.setAttribute('clip-path', `url(#${id})`);
    svg.appendChild(image);
    void fill.image.then((url) => {
      if (url) image.setAttribute('href', url);
    });
  }

  const dash = paint.dash ? DASHES[paint.dash] : undefined;
  for (const path of paths) {
    const node = create('path');
    node.setAttribute('d', path.d);
    node.setAttribute('fill', path.filled ? fillValue : 'none');
    node.setAttribute('fill-rule', 'evenodd');
    if (paint.stroke && path.stroked) {
      node.setAttribute('stroke', paint.stroke);
      node.setAttribute('stroke-width', fmt(paint.strokeWidth));
      node.setAttribute('vector-effect', 'non-scaling-stroke');
      if (dash) {
        node.setAttribute(
          'stroke-dasharray',
          dash.map((length) => fmt(length * paint.strokeWidth)).join(' '),
        );
      }
      if (!path.filled && (paint.headEnd || paint.tailEnd)) {
        const id = `av-arrow-${++sequence}`;
        const marker = create('marker');
        marker.setAttribute('id', id);
        marker.setAttribute('viewBox', '0 0 10 10');
        marker.setAttribute('refX', '9');
        marker.setAttribute('refY', '5');
        marker.setAttribute('markerWidth', '8');
        marker.setAttribute('markerHeight', '8');
        marker.setAttribute('orient', 'auto-start-reverse');
        const head = create('path');
        head.setAttribute('d', 'M0 0L10 5L0 10Z');
        head.setAttribute('fill', paint.stroke);
        marker.appendChild(head);
        const defs = create('defs');
        defs.appendChild(marker);
        svg.appendChild(defs);
        if (paint.tailEnd) node.setAttribute('marker-end', `url(#${id})`);
        if (paint.headEnd) node.setAttribute('marker-start', `url(#${id})`);
      }
    }
    svg.appendChild(node);
  }
  return svg;
}
