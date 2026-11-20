import { AnyViewError } from '../core/errors';
import { h, setStyles } from '../ui/dom';
import type { SlideDeck } from '../ui/slides';
import { toBlobPart } from '../util/text';
import { attr, child, childElements, find, type XmlElement, type XmlNode } from '../util/xml';
import type { ZipArchive } from '../util/zip';
import { imageMimeType } from './ooxml';
import {
  collectOdfStyles,
  odfColor,
  odfLength,
  odfParagraphCss,
  odfTextCss,
  readOdfPart,
  requireOdfContent,
  resolveOdfStyle,
  type OdfStyle,
} from './odf';

/**
 * ODP → DOM. Frames, text boxes, pictures and basic shapes are placed at their
 * stored positions. Master-page artwork and animations are not reproduced.
 */

export interface OdpHost {
  objectUrl(blob: Blob): string;
}

const SHAPES = new Set(['frame', 'rect', 'ellipse', 'custom-shape', 'line', 'connector', 'g']);

class OdpPresentation implements SlideDeck {
  width = 1058;
  height = 595;
  count = 0;
  private pages: XmlElement[] = [];
  private styles = new Map<string, OdfStyle>();

  constructor(
    private readonly zip: ZipArchive,
    private readonly host: OdpHost,
  ) {}

  async open(): Promise<this> {
    const content = await requireOdfContent(this.zip, 'OpenDocument presentation');
    const stylesPart = await readOdfPart(this.zip, 'styles.xml');
    this.styles = collectOdfStyles(stylesPart, content);
    this.pages = childElements(find(content, 'body', 'presentation'), 'page');
    this.count = this.pages.length;
    if (this.count === 0) {
      throw new AnyViewError('PARSE_FAILED', 'This presentation does not contain any slides.');
    }
    const layout = find(stylesPart, 'automatic-styles', 'page-layout', 'page-layout-properties');
    this.width = odfLength(attr(layout, 'fo:page-width')) ?? this.width;
    this.height = odfLength(attr(layout, 'fo:page-height')) ?? this.height;
    return this;
  }

  async renderSlide(index: number): Promise<HTMLElement> {
    const page = this.pages[index];
    if (!page)
      throw new AnyViewError('PARSE_FAILED', `Slide ${index + 1} is missing from the file.`);
    const slide = h('div');
    const background = odfColor(
      resolveOdfStyle(this.styles, attr(page, 'draw:style-name'))['fill-color'],
    );
    if (background) slide.style.background = background;
    const tasks: Promise<void>[] = [];
    this.renderShapes(page, slide, tasks);
    await Promise.all(tasks);
    return slide;
  }

  private renderShapes(parent: XmlElement, target: HTMLElement, tasks: Promise<void>[]): void {
    for (const shape of childElements(parent)) {
      if (!SHAPES.has(shape.local)) continue;
      if (shape.local === 'g') {
        this.renderShapes(shape, target, tasks);
        continue;
      }
      // Speaker notes and other non-visual frames are not part of the slide.
      if (attr(shape, 'presentation:class') === 'notes') continue;
      const px = (name: string): number => odfLength(attr(shape, name)) ?? 0;
      const element = h('div', { class: 'av-slide-shape' });
      if (shape.local === 'line' || shape.local === 'connector') {
        const [x1, y1, x2, y2] = [px('svg:x1'), px('svg:y1'), px('svg:x2'), px('svg:y2')];
        setStyles(element, {
          left: `${Math.min(x1, x2)}px`,
          top: `${Math.min(y1, y2)}px`,
          width: `${Math.max(1, Math.abs(x2 - x1))}px`,
          height: `${Math.max(1, Math.abs(y2 - y1))}px`,
          'border-top': Math.abs(y2 - y1) < 2 ? '1px solid #000' : undefined,
          'border-left': Math.abs(x2 - x1) < 2 ? '1px solid #000' : undefined,
        });
        target.appendChild(element);
        continue;
      }
      setStyles(element, {
        left: `${px('svg:x')}px`,
        top: `${px('svg:y')}px`,
        width: `${px('svg:width')}px`,
        height: px('svg:height') ? `${px('svg:height')}px` : undefined,
      });

      const graphic = resolveOdfStyle(this.styles, attr(shape, 'draw:style-name'));
      if (shape.local !== 'frame') {
        const fill = graphic['fill'] === 'none' ? undefined : odfColor(graphic['fill-color']);
        const stroke = graphic['stroke'] === 'none' ? undefined : odfColor(graphic['stroke-color']);
        setStyles(element, {
          background: fill,
          border: stroke ? `1px solid ${stroke}` : undefined,
          'border-radius': shape.local === 'ellipse' ? '50%' : undefined,
          'justify-content': 'center',
        });
      }

      const image = child(shape, 'image');
      const path = attr(image, 'xlink:href');
      if (image && path && !/^[a-z][\w+.-]*:/i.test(path)) {
        const entry = path.replace(/^\.\//, '');
        const mimeType = imageMimeType(entry);
        if (mimeType) {
          const picture = h('img', { attrs: { alt: '' } });
          element.appendChild(picture);
          tasks.push(
            this.zip.read(entry).then((data) => {
              if (data)
                picture.src = this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType }));
            }),
          );
        }
      }

      const kind = attr(shape, 'presentation:class');
      const defaultSize = kind === 'title' ? 40 : kind === 'subtitle' ? 28 : 18;
      const textParent = child(shape, 'text-box') ?? shape;
      const wrapper = h('div', {
        style: { padding: '4px 8px', 'font-size': `${defaultSize}pt`, width: '100%' },
      });
      this.renderText(textParent.children, wrapper);
      if (wrapper.firstChild) element.appendChild(wrapper);
      target.appendChild(element);
    }
  }

  private renderText(nodes: readonly XmlNode[], parent: HTMLElement): void {
    for (const node of nodes) {
      if (typeof node === 'string') continue;
      if (node.local === 'p' || node.local === 'h') {
        const properties = resolveOdfStyle(this.styles, attr(node, 'text:style-name'));
        const paragraph = h('p', {
          style: { ...odfParagraphCss(properties), ...odfTextCss(properties) },
        });
        this.renderInline(node.children, paragraph);
        if (!paragraph.firstChild) paragraph.appendChild(document.createTextNode('\u200b'));
        parent.appendChild(paragraph);
      } else if (node.local === 'list') {
        const list = h('ul', { style: { margin: '0', 'padding-left': '1.2em' } });
        for (const item of childElements(node, 'list-item')) {
          const li = h('li');
          this.renderText(item.children, li);
          list.appendChild(li);
        }
        parent.appendChild(list);
      }
    }
  }

  private renderInline(nodes: readonly XmlNode[], parent: HTMLElement): void {
    for (const node of nodes) {
      if (typeof node === 'string') {
        parent.appendChild(document.createTextNode(node.replace(/[ \t\r\n]+/g, ' ')));
      } else if (node.local === 'span' || node.local === 'a') {
        const span = h('span', {
          style: odfTextCss(resolveOdfStyle(this.styles, attr(node, 'text:style-name'))),
        });
        this.renderInline(node.children, span);
        parent.appendChild(span);
      } else if (node.local === 's') {
        parent.appendChild(
          document.createTextNode(
            ' '.repeat(Math.min(Number(attr(node, 'text:c') ?? 1) || 1, 1000)),
          ),
        );
      } else if (node.local === 'tab') {
        parent.appendChild(document.createTextNode('\t'));
      } else if (node.local === 'line-break') {
        parent.appendChild(h('br'));
      }
    }
  }
}

/** Opens an ODP package. Slides are rendered on demand through the returned deck. */
export function openOdpPresentation(zip: ZipArchive, host: OdpHost): Promise<SlideDeck> {
  return new OdpPresentation(zip, host).open();
}
