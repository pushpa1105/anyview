import { externalLink, h, setStyles } from '../ui/dom';
import { toBlobPart } from '../util/text';
import { safeLinkUrl } from '../util/url';
import {
  attr,
  child,
  childElements,
  find,
  textOf,
  type XmlElement,
  type XmlNode,
} from '../util/xml';
import type { ZipArchive } from '../util/zip';
import { imageMimeType } from './ooxml';
import {
  collectOdfStyles,
  odfLength,
  odfParagraphCss,
  odfTextCss,
  readOdfPart,
  requireOdfContent,
  resolveOdfStyle,
  type OdfStyle,
} from './odf';

/**
 * ODT → DOM. Supports headings, paragraphs, character and paragraph styles
 * with inheritance, nested lists, tables, images and hyperlinks.
 */

export interface OdtHost {
  objectUrl(blob: Blob): string;
}

class OdtRenderer {
  private styles = new Map<string, OdfStyle>();
  private tasks: Promise<void>[] = [];

  constructor(
    private readonly zip: ZipArchive,
    private readonly host: OdtHost,
  ) {}

  async render(): Promise<HTMLElement> {
    const content = await requireOdfContent(this.zip, 'OpenDocument text document');
    const stylesPart = await readOdfPart(this.zip, 'styles.xml');
    this.styles = collectOdfStyles(stylesPart, content);

    const paper = h('div');
    this.applyPageLayout(paper, stylesPart);
    this.renderBlocks(find(content, 'body', 'text')?.children ?? [], paper);
    await Promise.all(this.tasks);
    return paper;
  }

  private applyPageLayout(paper: HTMLElement, stylesPart: XmlElement | null): void {
    const layout = find(stylesPart, 'automatic-styles', 'page-layout', 'page-layout-properties');
    const px = (name: string, fallback: number): string =>
      `${(odfLength(attr(layout, name)) ?? fallback).toFixed(0)}px`;
    setStyles(paper, {
      width: px('fo:page-width', 794),
      'min-height': px('fo:page-height', 1123),
      padding: ['top', 'right', 'bottom', 'left']
        .map((side) => px(`fo:margin-${side}`, 76))
        .join(' '),
      'font-family': '"Liberation Serif", "Times New Roman", serif',
      'font-size': '12pt',
    });
  }

  private renderBlocks(nodes: readonly XmlNode[], parent: HTMLElement): void {
    for (const node of nodes) {
      if (typeof node === 'string') continue;
      switch (node.local) {
        case 'h': {
          const level = Math.min(
            6,
            Math.max(1, Number(attr(node, 'text:outline-level') ?? 1) || 1),
          );
          parent.appendChild(this.renderParagraph(node, `h${level}` as 'h1'));
          break;
        }
        case 'p':
          parent.appendChild(this.renderParagraph(node, 'p'));
          break;
        case 'list':
          parent.appendChild(this.renderList(node, 0, attr(node, 'text:style-name')));
          break;
        case 'table':
          parent.appendChild(this.renderTable(node));
          break;
        case 'section':
        case 'index-body':
        case 'table-of-content':
        case 'tracked-changes':
          if (node.local !== 'tracked-changes') this.renderBlocks(node.children, parent);
          break;
      }
    }
  }

  private renderParagraph(paragraph: XmlElement, tag: 'p' | 'h1'): HTMLElement {
    const properties = resolveOdfStyle(this.styles, attr(paragraph, 'text:style-name'));
    const textCss = odfTextCss(properties);
    if (tag !== 'p') textCss['font-weight'] ??= 'bold';
    const element = h(tag, {
      style: { ...odfParagraphCss(properties), ...textCss, 'white-space': 'pre-wrap' },
    });
    this.renderInline(paragraph.children, element);
    return element;
  }

  private renderInline(nodes: readonly XmlNode[], parent: HTMLElement): void {
    for (const node of nodes) {
      if (typeof node === 'string') {
        // ODF collapses runs of whitespace like HTML does; explicit spaces use <text:s>.
        parent.appendChild(document.createTextNode(node.replace(/[ \t\r\n]+/g, ' ')));
        continue;
      }
      switch (node.local) {
        case 'span': {
          const span = h('span', {
            style: odfTextCss(resolveOdfStyle(this.styles, attr(node, 'text:style-name'))),
          });
          this.renderInline(node.children, span);
          parent.appendChild(span);
          break;
        }
        case 'a': {
          const href = safeLinkUrl(attr(node, 'xlink:href'));
          const link = href ? externalLink(href) : h('span');
          this.renderInline(node.children, link);
          parent.appendChild(link);
          break;
        }
        case 's':
          parent.appendChild(
            document.createTextNode(
              ' '.repeat(Math.min(Number(attr(node, 'text:c') ?? 1) || 1, 1000)),
            ),
          );
          break;
        case 'tab':
          parent.appendChild(document.createTextNode('\t'));
          break;
        case 'line-break':
          parent.appendChild(h('br'));
          break;
        case 'frame':
          this.renderFrame(node, parent);
          break;
        case 'note': {
          const citation = find(node, 'note-citation');
          parent.appendChild(h('sup', { text: citation ? textOf(citation) : '*' }));
          break;
        }
        case 'bookmark-ref':
        case 'page-number':
        case 'date':
        case 'time':
        case 'author-name':
        case 'title':
          this.renderInline(node.children, parent);
          break;
      }
    }
  }

  private renderFrame(frame: XmlElement, parent: HTMLElement): void {
    const width = odfLength(attr(frame, 'svg:width'));
    const height = odfLength(attr(frame, 'svg:height'));
    const size = {
      width: width ? `${width.toFixed(0)}px` : undefined,
      height: height ? `${height.toFixed(0)}px` : undefined,
    };
    const image = child(frame, 'image');
    const path = attr(image, 'xlink:href');
    if (image && path && !/^[a-z][\w+.-]*:/i.test(path)) {
      const target = path.replace(/^\.\//, '');
      const mimeType = imageMimeType(target);
      if (!mimeType) return;
      const element = h('img', { attrs: { alt: '' }, style: size });
      parent.appendChild(element);
      this.tasks.push(
        this.zip.read(target).then((data) => {
          if (data)
            element.src = this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType }));
        }),
      );
      return;
    }
    const textBox = child(frame, 'text-box');
    if (textBox) {
      const box = h('span', {
        style: { display: 'inline-block', 'vertical-align': 'top', width: size.width },
      });
      this.renderBlocks(textBox.children, box);
      parent.appendChild(box);
    }
  }

  private renderList(list: XmlElement, depth: number, styleName: string | undefined): HTMLElement {
    // A list style has one child per level: numbered or bulleted.
    const levelStyles = styleName ? childElements(this.styles.get(styleName)?.element) : [];
    const levelStyle =
      levelStyles.find((level) => attr(level, 'text:level') === String(depth + 1)) ??
      levelStyles[depth];
    const ordered = levelStyle?.local === 'list-level-style-number';
    const element = h(ordered ? 'ol' : 'ul');
    if (ordered) {
      const format = attr(levelStyle, 'style:num-format');
      const type =
        format === 'a' || format === 'A' || format === 'i' || format === 'I' ? format : undefined;
      if (type) element.setAttribute('type', type);
    }
    for (const item of childElements(list)) {
      if (item.local !== 'list-item' && item.local !== 'list-header') continue;
      const li = h('li');
      for (const node of childElements(item)) {
        if (node.local === 'list') {
          li.appendChild(
            this.renderList(node, depth + 1, attr(node, 'text:style-name') ?? styleName),
          );
        } else {
          this.renderBlocks([node], li);
        }
      }
      element.appendChild(li);
    }
    return element;
  }

  private renderTable(table: XmlElement): HTMLElement {
    const tbody = h('tbody');
    const visit = (parent: XmlElement): void => {
      for (const row of childElements(parent)) {
        if (
          row.local === 'table-header-rows' ||
          row.local === 'table-row-group' ||
          row.local === 'table-rows'
        ) {
          visit(row);
          continue;
        }
        if (row.local !== 'table-row') continue;
        const tr = h('tr');
        for (const cell of childElements(row, 'table-cell')) {
          const columnSpan = Number(attr(cell, 'table:number-columns-spanned') ?? 1);
          const rowSpan = Number(attr(cell, 'table:number-rows-spanned') ?? 1);
          const td = h('td', {
            attrs: {
              colspan: columnSpan > 1 ? columnSpan : undefined,
              rowspan: rowSpan > 1 ? rowSpan : undefined,
            },
            style: { border: '1px solid #a6a6a6' },
          });
          this.renderBlocks(cell.children, td);
          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
    };
    visit(table);
    return h('table', { style: { width: '100%' } }, tbody);
  }
}

/** Renders an ODT package into a detached element. */
export function renderOdt(zip: ZipArchive, host: OdtHost): Promise<HTMLElement> {
  return new OdtRenderer(zip, host).render();
}
