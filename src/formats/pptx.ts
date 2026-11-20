import { AnyViewError } from '../core/errors';
import { externalLink, h, setStyles } from '../ui/dom';
import type { SlideDeck } from '../ui/slides';
import { toBlobPart } from '../util/text';
import { safeLinkUrl } from '../util/url';
import {
  attr,
  child,
  childElements,
  descendants,
  find,
  textOf,
  type XmlElement,
} from '../util/xml';
import type { ZipArchive } from '../util/zip';
import { renderChart } from '../ui/chart';
import { parseChart } from './chart';
import {
  applyColorModifiers,
  buildShapeSvg,
  isLineGeometry,
  parseGradient,
  resolveColor,
  shapeGeometry,
  type ShapePaint,
} from './drawingml';
import { EMPTY_THEME, parseTheme, type OfficeTheme } from './theme';
import {
  emuToPx,
  hexColor,
  imageMimeType,
  readRelationships,
  readXmlPart,
  requireXmlPart,
  type Relationship,
} from './ooxml';

/**
 * PPTX → DOM. Each slide becomes an absolutely positioned layout of shapes,
 * pictures, tables and text, with placeholders, text styles, colours and
 * backgrounds inherited from the slide layout, master and theme. Shapes are
 * drawn from their preset or custom geometry and charts from their cached
 * values. Animations and SmartArt are not reproduced.
 */

export { applyColorModifiers };

export interface PptxHost {
  objectUrl(blob: Blob): string;
}

interface Part {
  path: string;
  xml: XmlElement;
  relationships: Map<string, Relationship>;
}

interface Transform {
  offsetX: number;
  offsetY: number;
  scaleX: number;
  scaleY: number;
}

interface Box {
  x: number;
  y: number;
  width: number;
  height: number;
  rotation: number;
  flipH: boolean;
  flipV: boolean;
}

interface SlideContext {
  slide: Part;
  layout: Part | null;
  master: Part | null;
  /** Theme colour scheme, e.g. `accent1` → `#4472C4`. */
  colors: Map<string, string>;
  /** Master colour map, e.g. `tx1` → `dk1`. */
  colorMap: Map<string, string>;
  theme: OfficeTheme;
  number: number;
  tasks: Promise<void>[];
}

const IDENTITY: Transform = { offsetX: 0, offsetY: 0, scaleX: 1, scaleY: 1 };

const num = (value: string | undefined): number | undefined => {
  if (value === undefined) return undefined;
  const parsed = Number(value);
  return Number.isFinite(parsed) ? parsed : undefined;
};

class Presentation implements SlideDeck {
  width = 960;
  height = 540;
  count = 0;
  private slidePaths: string[] = [];
  private parts = new Map<string, Promise<Part | null>>();
  private defaultTextStyle: XmlElement | null = null;

  constructor(
    private readonly zip: ZipArchive,
    private readonly host: PptxHost,
  ) {}

  async open(): Promise<this> {
    const rootRelationships = await readRelationships(this.zip, '');
    const mainPath =
      [...rootRelationships.values()].find((rel) => rel.type === 'officeDocument')?.target ??
      'ppt/presentation.xml';
    const presentation = await requireXmlPart(this.zip, mainPath, 'PowerPoint presentation');
    const relationships = await readRelationships(this.zip, mainPath);
    const size = child(presentation, 'sldSz');
    this.width = emuToPx(num(attr(size, 'cx')) ?? 9144000);
    this.height = emuToPx(num(attr(size, 'cy')) ?? 5143500);
    this.defaultTextStyle = child(presentation, 'defaultTextStyle');
    for (const slide of childElements(child(presentation, 'sldIdLst'), 'sldId')) {
      const target = relationships.get(attr(slide, 'r:id') ?? '')?.target;
      if (target) this.slidePaths.push(target);
    }
    this.count = this.slidePaths.length;
    if (this.count === 0) {
      throw new AnyViewError('PARSE_FAILED', 'This presentation does not contain any slides.');
    }
    return this;
  }

  private loadPart(path: string | undefined): Promise<Part | null> {
    if (!path) return Promise.resolve(null);
    let pending = this.parts.get(path);
    if (!pending) {
      pending = (async () => {
        const xml = await readXmlPart(this.zip, path);
        if (!xml) return null;
        return { path, xml, relationships: await readRelationships(this.zip, path) };
      })();
      this.parts.set(path, pending);
    }
    return pending;
  }

  private related(part: Part | null, type: string): string | undefined {
    if (!part) return undefined;
    for (const relationship of part.relationships.values()) {
      if (relationship.type === type && !relationship.external) return relationship.target;
    }
    return undefined;
  }

  async renderSlide(index: number): Promise<HTMLElement> {
    const slide = await this.loadPart(this.slidePaths[index]);
    if (!slide)
      throw new AnyViewError('PARSE_FAILED', `Slide ${index + 1} is missing from the file.`);
    const layout = await this.loadPart(this.related(slide, 'slideLayout'));
    const master = await this.loadPart(this.related(layout, 'slideMaster'));
    const theme = await this.loadPart(this.related(master, 'theme'));

    const colors = new Map<string, string>();
    for (const entry of childElements(find(theme?.xml, 'themeElements', 'clrScheme'))) {
      const value = child(entry, 'srgbClr') ?? child(entry, 'sysClr');
      const hex = hexColor(attr(value, 'val')) ?? hexColor(attr(value, 'lastClr'));
      if (hex) colors.set(entry.local, hex);
    }
    const colorMap = new Map<string, string>(
      Object.entries(child(master?.xml, 'clrMap')?.attrs ?? {}),
    );
    const context: SlideContext = {
      slide,
      layout,
      master,
      colors,
      colorMap,
      theme: theme ? parseTheme(theme.xml) : EMPTY_THEME,
      number: index + 1,
      tasks: [],
    };

    const element = h('div');
    this.applyBackground(element, context);
    const showMaster = (part: Part | null): boolean => attr(part?.xml, 'showMasterSp') !== '0';
    // Decorative (non-placeholder) shapes of the master and layout sit behind the slide's own.
    if (master && showMaster(layout) && showMaster(slide)) {
      this.renderTree(find(master.xml, 'cSld', 'spTree'), master, context, element, IDENTITY, true);
    }
    if (layout && showMaster(slide)) {
      this.renderTree(find(layout.xml, 'cSld', 'spTree'), layout, context, element, IDENTITY, true);
    }
    this.renderTree(find(slide.xml, 'cSld', 'spTree'), slide, context, element, IDENTITY, false);
    await Promise.all(context.tasks);
    return element;
  }

  // --- colours and fills -------------------------------------------------

  /** Resolves the colour element inside `parent` (`srgbClr`, `schemeClr`, …) to a CSS colour. */
  private color(parent: XmlElement | null, context: SlideContext): string | undefined {
    return resolveColor(parent, (name) => context.colors.get(context.colorMap.get(name) ?? name));
  }

  /** Fill and line of a shape for SVG drawing. */
  private paint(
    properties: XmlElement | null,
    style: XmlElement | null,
    part: Part,
    context: SlideContext,
    useStyle: boolean,
    scale: number,
  ): ShapePaint {
    const scheme = (name: string): string | undefined =>
      context.colors.get(context.colorMap.get(name) ?? name);
    let fill: ShapePaint['fill'] | undefined;
    for (const element of childElements(properties)) {
      if (element.local === 'noFill') fill = null;
      else if (element.local === 'solidFill') fill = this.color(element, context) ?? null;
      else if (element.local === 'gradFill') fill = parseGradient(element, scheme);
      else if (element.local === 'blipFill') {
        fill = {
          image: new Promise<string | null>((resolve) => {
            const relationship = part.relationships.get(
              attr(child(element, 'blip'), 'r:embed') ?? '',
            );
            const mimeType =
              relationship && !relationship.external && imageMimeType(relationship.target);
            if (!relationship || !mimeType) return resolve(null);
            context.tasks.push(
              this.zip.read(relationship.target).then((data) => {
                resolve(
                  data
                    ? this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType }))
                    : null,
                );
              }),
            );
          }),
        };
      } else continue;
      break;
    }
    if (fill === undefined) {
      fill = useStyle ? (this.styleColor(child(style, 'fillRef'), context) ?? null) : null;
    }
    const line = child(properties, 'ln');
    let stroke: string | null | undefined;
    if (line) {
      if (child(line, 'noFill')) stroke = null;
      else if (child(line, 'solidFill'))
        stroke = this.color(child(line, 'solidFill'), context) ?? null;
    }
    if (stroke === undefined) {
      stroke = useStyle ? (this.styleColor(child(style, 'lnRef'), context) ?? null) : null;
    }
    const ending = (name: string): boolean => {
      const type = attr(child(line, name), 'type');
      return type !== undefined && type !== 'none';
    };
    return {
      fill,
      stroke,
      strokeWidth: Math.max(0.75, emuToPx(num(attr(line, 'w')) ?? 9525) * scale),
      dash: attr(child(line, 'prstDash'), 'val'),
      headEnd: ending('headEnd'),
      tailEnd: ending('tailEnd'),
    };
  }

  /** CSS `background` for the fill described in `properties`, `"none"` for no fill, or `undefined`. */
  private fill(
    properties: XmlElement | null,
    part: Part,
    context: SlideContext,
    target?: HTMLElement,
  ): string | undefined {
    for (const element of childElements(properties)) {
      if (element.local === 'noFill') return 'none';
      if (element.local === 'solidFill') return this.color(element, context);
      if (element.local === 'gradFill') {
        const stops = childElements(child(element, 'gsLst'), 'gs')
          .map((stop) => ({
            position: (num(attr(stop, 'pos')) ?? 0) / 1000,
            color: this.color(stop, context),
          }))
          .filter((stop): stop is { position: number; color: string } => stop.color !== undefined)
          .sort((a, b) => a.position - b.position);
        if (stops.length === 0) return undefined;
        if (stops.length === 1) return stops[0]!.color;
        const angle = (num(attr(child(element, 'lin'), 'ang')) ?? 5400000) / 60000 + 90;
        return `linear-gradient(${angle}deg, ${stops.map((stop) => `${stop.color} ${stop.position}%`).join(', ')})`;
      }
      if (element.local === 'blipFill' && target) {
        this.loadImage(attr(child(element, 'blip'), 'r:embed'), part, context, (url) => {
          setStyles(target, {
            'background-image': `url("${url}")`,
            'background-size': child(element, 'tile') ? 'auto' : '100% 100%',
            'background-repeat': child(element, 'tile') ? 'repeat' : 'no-repeat',
          });
        });
        return undefined;
      }
    }
    return undefined;
  }

  private loadImage(
    relationshipId: string | undefined,
    part: Part,
    context: SlideContext,
    apply: (url: string) => void,
  ): void {
    const relationship = relationshipId ? part.relationships.get(relationshipId) : undefined;
    if (!relationship || relationship.external) return;
    const mimeType = imageMimeType(relationship.target);
    if (!mimeType) return;
    context.tasks.push(
      this.zip.read(relationship.target).then((data) => {
        if (data) apply(this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType })));
      }),
    );
  }

  private applyBackground(element: HTMLElement, context: SlideContext): void {
    for (const part of [context.slide, context.layout, context.master]) {
      const background = find(part?.xml, 'cSld', 'bg');
      if (!background || !part) continue;
      const properties = child(background, 'bgPr');
      const fill = properties
        ? this.fill(properties, part, context, element)
        : this.color(child(background, 'bgRef'), context);
      if (fill && fill !== 'none') element.style.background = fill;
      return;
    }
  }

  // --- geometry ----------------------------------------------------------

  private box(xfrm: XmlElement | null, transform: Transform): Box | null {
    const offset = child(xfrm, 'off');
    const extent = child(xfrm, 'ext');
    if (!xfrm || !offset || !extent) return null;
    return {
      x: emuToPx(transform.offsetX + transform.scaleX * (num(attr(offset, 'x')) ?? 0)),
      y: emuToPx(transform.offsetY + transform.scaleY * (num(attr(offset, 'y')) ?? 0)),
      width: emuToPx(transform.scaleX * (num(attr(extent, 'cx')) ?? 0)),
      height: emuToPx(transform.scaleY * (num(attr(extent, 'cy')) ?? 0)),
      rotation: (num(attr(xfrm, 'rot')) ?? 0) / 60000,
      flipH: attr(xfrm, 'flipH') === '1',
      flipV: attr(xfrm, 'flipV') === '1',
    };
  }

  private place(element: HTMLElement, box: Box): void {
    const transforms = [
      box.rotation ? `rotate(${box.rotation}deg)` : '',
      box.flipH ? 'scaleX(-1)' : '',
      box.flipV ? 'scaleY(-1)' : '',
    ].filter(Boolean);
    setStyles(element, {
      left: `${box.x.toFixed(2)}px`,
      top: `${box.y.toFixed(2)}px`,
      width: `${box.width.toFixed(2)}px`,
      height: `${box.height.toFixed(2)}px`,
      transform: transforms.join(' ') || undefined,
    });
  }

  // --- placeholders ------------------------------------------------------

  private placeholderOf(shape: XmlElement): XmlElement | null {
    for (const wrapper of childElements(shape)) {
      if (!wrapper.local.startsWith('nv')) continue;
      const placeholder = find(wrapper, 'nvPr', 'ph');
      if (placeholder) return placeholder;
    }
    return null;
  }

  /** Finds the shape in `part` that a slide placeholder inherits from. */
  private findPlaceholder(
    part: Part | null,
    placeholder: XmlElement,
    byIndex: boolean,
  ): XmlElement | null {
    if (!part) return null;
    const type = attr(placeholder, 'type') ?? 'body';
    const index = attr(placeholder, 'idx');
    const family = (value: string): string =>
      value === 'ctrTitle' ? 'title' : value === 'subTitle' || value === 'obj' ? 'body' : value;
    let byType: XmlElement | null = null;
    for (const shape of childElements(find(part.xml, 'cSld', 'spTree'))) {
      const candidate = this.placeholderOf(shape);
      if (!candidate) continue;
      if (byIndex && index !== undefined && attr(candidate, 'idx') === index) return shape;
      if (!byType && family(attr(candidate, 'type') ?? 'body') === family(type)) byType = shape;
    }
    return byType;
  }

  // --- shape tree --------------------------------------------------------

  private renderTree(
    tree: XmlElement | null,
    part: Part,
    context: SlideContext,
    parent: HTMLElement,
    transform: Transform,
    decorationOnly: boolean,
  ): void {
    for (const raw of childElements(tree)) {
      // Prefer the fallback of alternate content: it is the widely compatible form.
      const node =
        raw.local === 'AlternateContent'
          ? (childElements(child(raw, 'Fallback'))[0] ?? childElements(child(raw, 'Choice'))[0])
          : raw;
      if (!node) continue;
      if (decorationOnly && this.placeholderOf(node)) continue;
      switch (node.local) {
        case 'sp':
        case 'cxnSp':
          this.renderShape(node, part, context, parent, transform);
          break;
        case 'pic':
          this.renderPicture(node, part, context, parent, transform);
          break;
        case 'graphicFrame':
          this.renderFrame(node, part, context, parent, transform);
          break;
        case 'grpSp': {
          const xfrm = find(node, 'grpSpPr', 'xfrm');
          const offset = child(xfrm, 'off');
          const extent = child(xfrm, 'ext');
          const childOffset = child(xfrm, 'chOff');
          const childExtent = child(xfrm, 'chExt');
          let inner = transform;
          if (offset && extent && childOffset && childExtent) {
            const ratioX =
              (num(attr(extent, 'cx')) ?? 1) / ((num(attr(childExtent, 'cx')) ?? 1) || 1);
            const ratioY =
              (num(attr(extent, 'cy')) ?? 1) / ((num(attr(childExtent, 'cy')) ?? 1) || 1);
            const scaleX = transform.scaleX * ratioX;
            const scaleY = transform.scaleY * ratioY;
            inner = {
              scaleX,
              scaleY,
              offsetX:
                transform.offsetX +
                transform.scaleX * (num(attr(offset, 'x')) ?? 0) -
                scaleX * (num(attr(childOffset, 'x')) ?? 0),
              offsetY:
                transform.offsetY +
                transform.scaleY * (num(attr(offset, 'y')) ?? 0) -
                scaleY * (num(attr(childOffset, 'y')) ?? 0),
            };
          }
          this.renderTree(node, part, context, parent, inner, decorationOnly);
          break;
        }
      }
    }
  }

  private renderShape(
    shape: XmlElement,
    part: Part,
    context: SlideContext,
    parent: HTMLElement,
    transform: Transform,
  ): void {
    const placeholder = this.placeholderOf(shape);
    const layoutShape = placeholder
      ? this.findPlaceholder(context.layout, placeholder, true)
      : null;
    const masterShape = placeholder
      ? this.findPlaceholder(
          context.master,
          layoutShape ? (this.placeholderOf(layoutShape) ?? placeholder) : placeholder,
          false,
        )
      : null;
    const properties = child(shape, 'spPr');
    const box =
      this.box(child(properties, 'xfrm'), transform) ??
      this.box(find(layoutShape, 'spPr', 'xfrm'), IDENTITY) ??
      this.box(find(masterShape, 'spPr', 'xfrm'), IDENTITY);
    if (!box) return;

    const geometry = attr(child(properties, 'prstGeom'), 'prst') ?? 'rect';
    const line = child(properties, 'ln');
    const style = child(shape, 'style');
    const custom = child(properties, 'custGeom') !== null;
    const isLine = shape.local === 'cxnSp' || (!custom && isLineGeometry(geometry));

    const element = h('div', { class: 'av-slide-shape' });
    this.place(element, box);

    const dash = attr(child(line, 'prstDash'), 'val');
    const dashed = dash !== undefined && dash !== 'solid';
    const simple = !custom && !isLine && !dashed && /^(rect|roundRect|ellipse)$/.test(geometry);
    if (!simple) {
      // Drawn from the shape's outline; flips are already applied by `place`.
      const paint = this.paint(properties, style, part, context, !placeholder, transform.scaleX);
      if (isLine && paint.stroke === null && !child(line, 'noFill')) paint.stroke = '#000000';
      if (isLine) paint.fill = null;
      if (paint.fill !== null || paint.stroke !== null) {
        const paths = shapeGeometry(
          shape.local === 'cxnSp' && !custom && !child(properties, 'prstGeom') ? null : properties,
          box.width,
          box.height,
        );
        const outline =
          isLine && !custom && !/Connector[2-5]/.test(geometry)
            ? [
                {
                  d: `M0 0L${box.width.toFixed(2)} ${box.height.toFixed(2)}`,
                  filled: false,
                  stroked: true,
                },
              ]
            : paths;
        element.appendChild(buildShapeSvg(outline, box.width, box.height, paint));
      }
    } else {
      let lineColor = line ? this.fill(line, part, context) : undefined;
      if (lineColor === undefined && !placeholder)
        lineColor = this.styleColor(child(style, 'lnRef'), context);
      const lineWidth = Math.max(1, emuToPx(num(attr(line, 'w')) ?? 9525) * transform.scaleX);
      let fill = this.fill(properties, part, context, element);
      if (fill === undefined && !placeholder)
        fill = this.styleColor(child(style, 'fillRef'), context);
      const smaller = Math.min(box.width, box.height);
      const radius = num(
        /^val\s+(\d+)/.exec(attr(find(properties, 'prstGeom', 'avLst', 'gd'), 'fmla') ?? '')?.[1],
      );
      setStyles(element, {
        background: fill && fill !== 'none' ? fill : undefined,
        border:
          lineColor && lineColor !== 'none'
            ? `${lineWidth.toFixed(2)}px solid ${lineColor}`
            : undefined,
        'border-radius':
          geometry === 'ellipse'
            ? '50%'
            : geometry === 'roundRect'
              ? `${(smaller * Math.min(0.5, (radius ?? 16667) / 100000)).toFixed(1)}px`
              : undefined,
      });
    }

    const body = child(shape, 'txBody');
    if (body) {
      const fontColor = !placeholder
        ? this.styleColor(child(style, 'fontRef'), context)
        : undefined;
      this.renderText(body, element, part, context, {
        placeholderType: placeholder ? (attr(placeholder, 'type') ?? 'body') : null,
        inherited: [layoutShape, masterShape],
        defaultColor: fontColor,
        // Text is mirrored back so flipped shapes stay readable.
        unflip:
          box.flipH || box.flipV
            ? `${box.flipH ? 'scaleX(-1)' : ''} ${box.flipV ? 'scaleY(-1)' : ''}`
            : undefined,
      });
    }
    parent.appendChild(element);
  }

  /** Colour of a `lnRef`/`fillRef`/`fontRef` style reference; index 0 means "none". */
  private styleColor(reference: XmlElement | null, context: SlideContext): string | undefined {
    if (!reference || attr(reference, 'idx') === '0') return undefined;
    return this.color(reference, context);
  }

  private renderPicture(
    picture: XmlElement,
    part: Part,
    context: SlideContext,
    parent: HTMLElement,
    transform: Transform,
  ): void {
    const placeholder = this.placeholderOf(picture);
    const layoutShape = placeholder
      ? this.findPlaceholder(context.layout, placeholder, true)
      : null;
    const box =
      this.box(find(picture, 'spPr', 'xfrm'), transform) ??
      this.box(find(layoutShape, 'spPr', 'xfrm'), IDENTITY);
    if (!box) return;
    const element = h('div', { class: 'av-slide-shape' });
    this.place(element, box);
    const image = h('img', {
      attrs: { alt: attr(find(picture, 'nvPicPr', 'cNvPr'), 'descr') ?? '' },
    });
    // A cropped picture shows only part of the image: the image is enlarged and shifted inside its frame.
    const crop = find(picture, 'blipFill', 'srcRect');
    const side = (name: string): number => (num(attr(crop, name)) ?? 0) / 100000;
    const [left, top, right, bottom] = [side('l'), side('t'), side('r'), side('b')];
    if (crop && (left || top || right || bottom)) {
      const scaleX = 1 / Math.max(0.01, 1 - left - right);
      const scaleY = 1 / Math.max(0.01, 1 - top - bottom);
      element.style.overflow = 'hidden';
      setStyles(image, {
        position: 'absolute',
        width: `${(scaleX * 100).toFixed(3)}%`,
        height: `${(scaleY * 100).toFixed(3)}%`,
        left: `${(-left * scaleX * 100).toFixed(3)}%`,
        top: `${(-top * scaleY * 100).toFixed(3)}%`,
        'max-width': 'none',
      });
    }
    this.loadImage(attr(find(picture, 'blipFill', 'blip'), 'r:embed'), part, context, (url) => {
      image.src = url;
      element.appendChild(image);
    });
    parent.appendChild(element);
  }

  private renderFrame(
    frame: XmlElement,
    part: Part,
    context: SlideContext,
    parent: HTMLElement,
    transform: Transform,
  ): void {
    const box = this.box(child(frame, 'xfrm'), transform);
    if (!box) return;
    const element = h('div', { class: 'av-slide-shape' });
    this.place(element, box);
    const data = find(frame, 'graphic', 'graphicData');
    const table = child(data, 'tbl');
    const chartReference = child(data, 'chart');
    const chartPath = chartReference
      ? part.relationships.get(attr(chartReference, 'r:id') ?? '')?.target
      : undefined;
    if (table) {
      element.appendChild(this.renderTable(table, part, context));
    } else if (chartPath) {
      context.tasks.push(
        readXmlPart(this.zip, chartPath).then((root) => {
          const model = parseChart(root, context.theme);
          if (!model) return;
          const chart = renderChart(model, box.width, box.height);
          chart.setAttribute('width', '100%');
          chart.setAttribute('height', '100%');
          element.appendChild(chart);
        }),
      );
    } else {
      // Embedded objects usually carry a preview picture; charts and SmartArt do not.
      const preview = descendants(frame, 'blip')[0];
      if (preview) {
        const image = h('img', { attrs: { alt: '' } });
        this.loadImage(attr(preview, 'r:embed'), part, context, (url) => {
          image.src = url;
          element.appendChild(image);
        });
      } else {
        const uri = attr(data, 'uri') ?? '';
        setStyles(element, {
          border: '1px dashed #9aa0a6',
          color: '#5f6368',
          'align-items': 'center',
          'justify-content': 'center',
          'font-size': '14pt',
        });
        element.textContent = uri.includes('chart')
          ? 'Chart'
          : uri.includes('diagram')
            ? 'Diagram'
            : 'Object';
      }
    }
    parent.appendChild(element);
  }

  private renderTable(table: XmlElement, part: Part, context: SlideContext): HTMLElement {
    const element = h('table');
    const group = h('colgroup');
    for (const column of childElements(child(table, 'tblGrid'), 'gridCol')) {
      group.appendChild(
        h('col', { style: { width: `${emuToPx(num(attr(column, 'w')) ?? 0).toFixed(1)}px` } }),
      );
    }
    element.appendChild(group);
    const tbody = h('tbody');
    for (const row of childElements(table, 'tr')) {
      const tr = h('tr', {
        style: { height: `${emuToPx(num(attr(row, 'h')) ?? 0).toFixed(1)}px` },
      });
      for (const cell of childElements(row, 'tc')) {
        // Cells swallowed by a merge are present in the file but must not be drawn.
        if (attr(cell, 'hMerge') === '1' || attr(cell, 'vMerge') === '1') continue;
        const td = h('td', {
          attrs: { colspan: num(attr(cell, 'gridSpan')), rowspan: num(attr(cell, 'rowSpan')) },
        });
        const cellProperties = child(cell, 'tcPr');
        const fill = this.fill(cellProperties, part, context);
        if (fill && fill !== 'none') td.style.background = fill;
        const edges: [string, string][] = [
          ['lnL', 'left'],
          ['lnR', 'right'],
          ['lnT', 'top'],
          ['lnB', 'bottom'],
        ];
        for (const [name, edge] of edges) {
          const border = child(cellProperties, name);
          if (!border) continue;
          const color = child(border, 'noFill')
            ? null
            : this.color(child(border, 'solidFill'), context);
          const width = Math.max(1, emuToPx(num(attr(border, 'w')) ?? 12700));
          td.style.setProperty(
            `border-${edge}`,
            color ? `${width.toFixed(2)}px solid ${color}` : 'none',
          );
        }
        const margin = (name: string, fallback: number): string =>
          `${emuToPx(num(attr(cellProperties, name)) ?? fallback).toFixed(1)}px`;
        td.style.padding = [
          margin('marT', 45720),
          margin('marR', 91440),
          margin('marB', 45720),
          margin('marL', 91440),
        ].join(' ');
        const anchor = attr(cellProperties, 'anchor');
        if (anchor === 'ctr') td.style.verticalAlign = 'middle';
        else if (anchor === 'b') td.style.verticalAlign = 'bottom';
        const body = child(cell, 'txBody');
        if (body) {
          this.renderText(body, td, part, context, {
            placeholderType: null,
            inherited: [],
            defaultColor: undefined,
            unflip: undefined,
          });
        }
        tr.appendChild(td);
      }
      tbody.appendChild(tr);
    }
    element.appendChild(tbody);
    return element;
  }

  // --- text --------------------------------------------------------------

  private renderText(
    body: XmlElement,
    parent: HTMLElement,
    part: Part,
    context: SlideContext,
    options: {
      placeholderType: string | null;
      inherited: (XmlElement | null)[];
      defaultColor: string | undefined;
      unflip: string | undefined;
    },
  ): void {
    const { placeholderType } = options;
    const inheritedBodies = options.inherited.map((shape) => child(shape, 'txBody'));
    const bodyProperties = [
      child(body, 'bodyPr'),
      ...inheritedBodies.map((item) => child(item, 'bodyPr')),
    ];
    const bodyAttr = (name: string): string | undefined => {
      for (const properties of bodyProperties) {
        const value = attr(properties, name);
        if (value !== undefined) return value;
      }
      return undefined;
    };

    // Text style sources, from most to least specific.
    const textStyles = child(context.master?.xml, 'txStyles');
    const masterStyle =
      placeholderType === null
        ? child(textStyles, 'otherStyle')
        : placeholderType === 'title' || placeholderType === 'ctrTitle'
          ? child(textStyles, 'titleStyle')
          : ['body', 'subTitle', 'obj', 'half', 'tbl', 'chart', 'pic'].includes(placeholderType)
            ? child(textStyles, 'bodyStyle')
            : child(textStyles, 'otherStyle');
    const listStyles = [
      child(body, 'lstStyle'),
      ...inheritedBodies.map((item) => child(item, 'lstStyle')),
      masterStyle,
      this.defaultTextStyle,
    ];
    const levelStyles = (level: number): XmlElement[] =>
      listStyles
        .map((list) => child(list, `lvl${level + 1}pPr`))
        .filter((item): item is XmlElement => item !== null);

    const autofit = find(body, 'bodyPr', 'normAutofit');
    const fontScale = (num(attr(autofit, 'fontScale')) ?? 100000) / 100000;
    const inset = (name: string, fallback: number): string =>
      `${emuToPx(num(bodyAttr(name)) ?? fallback).toFixed(1)}px`;
    const anchor = bodyAttr('anchor');
    const wrapper = h('div', {
      style: {
        padding: `${inset('tIns', 45720)} ${inset('rIns', 91440)} ${inset('bIns', 45720)} ${inset('lIns', 91440)}`,
        'margin-top': anchor === 'ctr' ? 'auto' : anchor === 'b' ? 'auto' : undefined,
        'margin-bottom': anchor === 'ctr' ? 'auto' : undefined,
        width: '100%',
        // Above the SVG outline of the shape.
        position: 'relative',
        'white-space': bodyAttr('wrap') === 'none' ? 'nowrap' : undefined,
        transform: options.unflip,
      },
    });

    const counters: number[] = [];
    for (const paragraph of childElements(body, 'p')) {
      const paragraphProperties = child(paragraph, 'pPr');
      const level = Math.min(8, Math.max(0, num(attr(paragraphProperties, 'lvl')) ?? 0));
      const levels = levelStyles(level);
      const sources = paragraphProperties ? [paragraphProperties, ...levels] : levels;
      const paragraphAttr = (name: string): string | undefined => {
        for (const source of sources) {
          const value = attr(source, name);
          if (value !== undefined) return value;
        }
        return undefined;
      };
      const firstChild = (...names: string[]): XmlElement | null => {
        for (const source of sources) {
          for (const name of names) {
            const found = child(source, name);
            if (found) return found;
          }
        }
        return null;
      };
      const defaults = sources
        .map((source) => child(source, 'defRPr'))
        .filter((item): item is XmlElement => item !== null);

      /** Resolves run-level CSS from a run's own properties and the inherited defaults. */
      const runCss = (runProperties: XmlElement | null): Record<string, string | undefined> => {
        const chain = runProperties ? [runProperties, ...defaults] : defaults;
        const value = (name: string): string | undefined => {
          for (const item of chain) {
            const found = attr(item, name);
            if (found !== undefined) return found;
          }
          return undefined;
        };
        let color: string | undefined;
        for (const item of chain) {
          const fill = child(item, 'solidFill');
          if (fill) {
            color = this.color(fill, context);
            break;
          }
        }
        let font: string | undefined;
        for (const item of chain) {
          const typeface = attr(child(item, 'latin'), 'typeface');
          if (typeface) {
            font = /^[\w .-]+$/.test(typeface) ? typeface : undefined;
            break;
          }
        }
        const size = (num(value('sz')) ?? 1800) / 100;
        const baseline = num(value('baseline')) ?? 0;
        const underline = value('u');
        const strike = value('strike');
        const decoration = [
          underline && underline !== 'none' ? 'underline' : '',
          strike && strike !== 'noStrike' ? 'line-through' : '',
        ]
          .filter(Boolean)
          .join(' ');
        return {
          'font-size': `${(size * fontScale * (baseline ? 0.7 : 1)).toFixed(1)}pt`,
          'font-weight': value('b') === '1' ? 'bold' : undefined,
          'font-style': value('i') === '1' ? 'italic' : undefined,
          'text-decoration': decoration || undefined,
          color: color ?? options.defaultColor,
          'font-family': font ? `"${font}", Calibri, Arial, sans-serif` : undefined,
          'vertical-align': baseline > 0 ? 'super' : baseline < 0 ? 'sub' : undefined,
          'text-transform': value('cap') === 'all' ? 'uppercase' : undefined,
        };
      };

      const align = paragraphAttr('algn');
      const lineSpacing = num(attr(find(firstChild('lnSpc'), 'spcPct'), 'val'));
      const spacing = (name: string): string | undefined => {
        const points = num(attr(find(firstChild(name), 'spcPts'), 'val'));
        return points ? `${(points / 100).toFixed(1)}pt` : undefined;
      };
      const marginLeft = emuToPx(num(paragraphAttr('marL')) ?? 0);
      const indent = emuToPx(num(paragraphAttr('indent')) ?? 0);
      const element = h('p', {
        style: {
          'text-align':
            align === 'ctr'
              ? 'center'
              : align === 'r'
                ? 'right'
                : align === 'just'
                  ? 'justify'
                  : undefined,
          'margin-left': marginLeft ? `${marginLeft.toFixed(1)}px` : undefined,
          'text-indent': indent ? `${indent.toFixed(1)}px` : undefined,
          'line-height': lineSpacing ? ((lineSpacing / 100000) * 1.2).toFixed(2) : '1.2',
          'margin-top': spacing('spcBef'),
          'margin-bottom': spacing('spcAft'),
        },
      });

      const runs = childElements(paragraph).filter((node) =>
        ['r', 'br', 'fld'].includes(node.local),
      );
      const hasText = runs.some((run) => run.local !== 'br' && textOf(child(run, 't')) !== '');

      // Bullets: the nearest definition wins, and `buNone` switches them off.
      const bullet = firstChild('buNone', 'buChar', 'buAutoNum');
      if (bullet?.local === 'buAutoNum') counters[level] = (counters[level] ?? 0) + 1;
      else counters.length = Math.min(counters.length, level + 1);
      if (hasText && bullet && bullet.local !== 'buNone') {
        const text =
          bullet.local === 'buChar'
            ? (attr(bullet, 'char') ?? '•').replace(/[\ue000-\uf8ff]/g, '•')
            : `${counters[level]}.`;
        const marker = h('span', {
          text,
          style: {
            ...runCss(child(runs[0], 'rPr')),
            display: 'inline-block',
            'min-width': indent < 0 ? `${(-indent).toFixed(1)}px` : '1em',
            'text-indent': '0',
            'text-decoration': undefined,
          },
        });
        element.appendChild(marker);
      }

      for (const run of runs) {
        if (run.local === 'br') {
          element.appendChild(h('br'));
          continue;
        }
        const properties = child(run, 'rPr');
        let text = textOf(child(run, 't'));
        if (run.local === 'fld' && attr(run, 'type') === 'slidenum') text = String(context.number);
        if (text === '') continue;
        const link = part.relationships.get(attr(child(properties, 'hlinkClick'), 'r:id') ?? '');
        const href = link?.external ? safeLinkUrl(link.target) : null;
        const span = href ? externalLink(href, text) : h('span', { text });
        setStyles(span, runCss(properties));
        element.appendChild(span);
      }
      if (!hasText) {
        // Empty paragraphs still take up one line at their font size.
        setStyles(element, { 'font-size': runCss(child(paragraph, 'endParaRPr'))['font-size'] });
        element.appendChild(document.createTextNode('\u200b'));
      }
      wrapper.appendChild(element);
    }
    parent.appendChild(wrapper);
  }
}

/** Opens a PPTX package. Slides are rendered on demand through the returned deck. */
export function openPresentation(zip: ZipArchive, host: PptxHost): Promise<SlideDeck> {
  return new Presentation(zip, host).open();
}
