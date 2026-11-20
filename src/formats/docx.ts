import { externalLink, h, setStyles } from '../ui/dom';
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
import {
  buildShapeSvg,
  parseGradient,
  resolveColor,
  shapeGeometry,
  type ShapePaint,
} from './drawingml';
import {
  createPage,
  fillFields,
  layoutTabs,
  Paginator,
  type HeaderSet,
  type PageSetup,
} from './docx-layout';
import {
  lineFactor,
  merge,
  mergeTabs,
  num,
  PropertyParser,
  twipsToPx,
  type Border,
  type BorderSet,
  type CellProps,
  type ConditionalFormat,
  type ParaProps,
  type RunProps,
  type Side,
  type StyleDef,
  type TableProps,
} from './docx-props';
import {
  emuToPx,
  imageMimeType,
  isOn,
  readRelationships,
  readXmlPart,
  requireXmlPart,
  type Relationship,
} from './ooxml';
import { EMPTY_THEME, loadTheme, safeFontName, type OfficeTheme } from './theme';

/**
 * DOCX → DOM, laid out like Word: pages with margins, headers and footers,
 * styles with inheritance, theme fonts and colours, numbered and bulleted
 * lists, tab stops with leaders, tables with table styles, pictures (inline and
 * floating), text boxes, drop caps, footnotes, endnotes and embedded fonts.
 */

export interface DocxHost {
  /** Registers a blob and returns a URL that the viewer revokes later. */
  objectUrl(blob: Blob): string;
}

export interface DocxDocument {
  /** Detached element holding the pages. */
  readonly element: HTMLElement;
  /**
   * Positions tab stops and splits the content into pages. Needs layout, so
   * call it after `element` has been added to the document.
   */
  layout(): Promise<void>;
  pageCount(): number;
  /** The element of a page (zero-based), or `null`. */
  page(index: number): HTMLElement | null;
  /** Releases the fonts the document registered. */
  destroy(): void;
}

interface NumberingLevel {
  start: number;
  format: string;
  text: string;
  suffix: string;
  align: string;
  para: ParaProps;
  run: RunProps;
  picture: string | undefined;
}

interface NumberingDef {
  abstractId: string;
  levels: Map<number, NumberingLevel>;
  /** Levels whose counter restarts the first time this list is used. */
  restarts: Map<number, number>;
}

interface Field {
  instruction: string;
  phase: 'instruction' | 'result';
  /** The cached result is replaced by a live value (page numbers). */
  skip: boolean;
  href: string | null;
  anchor: string | null;
  link: HTMLElement | null;
  linkParent: HTMLElement | null;
}

interface ParagraphState {
  sizes: number[];
  hasTabs: boolean;
  font: string | undefined;
  line: number | undefined;
  lineRule: string | undefined;
}

interface CellContext {
  para: ParaProps;
  run: RunProps;
}

interface TableLook {
  firstRow: boolean;
  lastRow: boolean;
  firstCol: boolean;
  lastCol: boolean;
  noHBand: boolean;
  noVBand: boolean;
}

interface RankedBorder {
  border: Border | undefined;
  rank: number;
}

interface CellBox {
  td: HTMLTableCellElement;
  row: number;
  col: number;
  colSpan: number;
  rowSpan: number;
  conditionals: { type: string; format: ConditionalFormat }[];
  direct: CellProps;
  borders: Record<Side, RankedBorder>;
}

const SOFT_HYPHEN = String.fromCharCode(0xad);
const NON_BREAKING_HYPHEN = String.fromCharCode(0x2011);
const SIDES: readonly Side[] = ['top', 'left', 'bottom', 'right'];

const FONT_FALLBACKS: Readonly<Record<string, string>> = {
  calibri: 'Carlito',
  'calibri light': 'Carlito',
  cambria: 'Caladea',
  'times new roman': 'Liberation Serif',
  arial: 'Liberation Sans',
  'courier new': 'Liberation Mono',
};

const SERIF =
  /times|cambria|georgia|garamond|palatino|book antiqua|constantia|serif|century school/i;
const MONO = /courier|consolas|mono|lucida console/i;

const px = (value: number | undefined): string | undefined =>
  value ? `${Math.round(value * 100) / 100}px` : undefined;
const twipsPx = (twips: number | undefined): string | undefined =>
  twips ? px(twipsToPx(twips)) : undefined;

function toRoman(value: number): string {
  if (value <= 0 || value >= 4000) return String(value);
  const numerals: [number, string][] = [
    [1000, 'M'],
    [900, 'CM'],
    [500, 'D'],
    [400, 'CD'],
    [100, 'C'],
    [90, 'XC'],
    [50, 'L'],
    [40, 'XL'],
    [10, 'X'],
    [9, 'IX'],
    [5, 'V'],
    [4, 'IV'],
    [1, 'I'],
  ];
  let remaining = value;
  let result = '';
  for (const [amount, numeral] of numerals) {
    while (remaining >= amount) {
      result += numeral;
      remaining -= amount;
    }
  }
  return result;
}

function toLetters(value: number): string {
  if (value <= 0) return String(value);
  // Word repeats the letter after Z: a…z, aa, bb, cc…
  const letter = String.fromCharCode(97 + ((value - 1) % 26));
  return letter.repeat(Math.floor((value - 1) / 26) + 1);
}

const ORDINALS = ['th', 'st', 'nd', 'rd'];

/** Formats a list counter according to a `w:numFmt` value. */
export function formatCounter(value: number, format: string): string {
  switch (format) {
    case 'lowerLetter':
      return toLetters(value);
    case 'upperLetter':
      return toLetters(value).toUpperCase();
    case 'lowerRoman':
      return toRoman(value).toLowerCase();
    case 'upperRoman':
      return toRoman(value);
    case 'decimalZero':
      return String(value).padStart(2, '0');
    case 'ordinal': {
      const tens = value % 100;
      return `${value}${tens >= 11 && tens <= 13 ? 'th' : (ORDINALS[value % 10] ?? 'th')}`;
    }
    case 'none':
      return '';
    default:
      return String(value);
  }
}

/** Glyphs of the Symbol and Wingdings fonts that lists commonly use, as Unicode. */
const SYMBOL_GLYPHS: Readonly<Record<number, string>> = {
  0xb7: '•',
  0xa7: '▪',
  0xa8: '◆',
  0x6e: '■',
  0x71: '❑',
  0x75: '◆',
  0x76: '❖',
  0xd8: '➢',
  0xfc: '✓',
  0xfb: '✗',
  0xe0: '→',
  0x2d: '–',
  0x6f: '○',
  0x9f: '•',
  0xa1: '○',
  0x6c: '●',
};

/** Bullets are often private-use glyphs of the Symbol/Wingdings fonts; map them to Unicode. */
function bulletGlyph(text: string, font: string | undefined): string {
  const code = text.codePointAt(0) ?? 0;
  const symbolFont = /symbol|wingdings|webdings/i.test(font ?? '');
  if (code >= 0xf000 && code <= 0xf0ff) return SYMBOL_GLYPHS[code - 0xf000] ?? '•';
  if (symbolFont && code < 0x100) return SYMBOL_GLYPHS[code] ?? '•';
  if (code >= 0xe000 && code <= 0xf8ff) return '•';
  return text || '•';
}

function borderCss(border: Border | undefined): string | undefined {
  if (!border) return undefined;
  if (border.style === 'none') return 'none';
  return `${border.width}px ${border.style} ${border.color}`;
}

/** Decides which of two borders meeting on a shared edge is drawn. */
function strongerBorder(a: RankedBorder, b: RankedBorder): RankedBorder {
  if (!a.border) return b;
  if (!b.border) return a;
  if (a.rank !== b.rank) return a.rank > b.rank ? a : b;
  const visibleA = a.border.style !== 'none';
  const visibleB = b.border.style !== 'none';
  if (visibleA !== visibleB) return visibleA ? a : b;
  return a.border.width >= b.border.width ? a : b;
}

/** Reverses the obfuscation Word applies to embedded fonts (ECMA-376 part 4, 2.8.1). */
export function deobfuscateFont(data: Uint8Array, fontKey: string): Uint8Array {
  const guid = fontKey.replace(/[{}-]/g, '');
  if (!/^[0-9a-fA-F]{32}$/.test(guid)) return data;
  const key: number[] = [];
  for (let i = 0; i < 16; i++) key[15 - i] = parseInt(guid.slice(i * 2, i * 2 + 2), 16);
  const result = data.slice();
  for (let i = 0; i < 32 && i < result.length; i++) result[i] = result[i]! ^ key[i % 16]!;
  return result;
}

class DocxRenderer {
  private theme: OfficeTheme = EMPTY_THEME;
  private parser = new PropertyParser(EMPTY_THEME);
  private styles = new Map<string, StyleDef>();
  private defaultParagraphStyle: string | undefined;
  private defaultTableStyle: string | undefined;
  private defaultRun: RunProps = {};
  private defaultPara: ParaProps = {};
  private numbering = new Map<string, NumberingDef>();
  private counters = new Map<string, number[]>();
  private usedLists = new Set<string>();
  private relationships = new Map<string, Relationship>();
  private footnotes: { root: XmlElement | null; relationships: Map<string, Relationship> } = {
    root: null,
    relationships: new Map(),
  };
  private endnotes = this.footnotes;
  private usedFootnotes: string[] = [];
  private usedEndnotes: string[] = [];
  private tasks: Promise<void>[] = [];
  private fields: Field[] = [];
  private paragraph: ParagraphState | null = null;
  private cells: CellContext[] = [];
  private sections: PageSetup[] = [];
  private sectionIndex = 0;
  private fontAliases = new Map<string, string>();
  private fontFamilies = new Map<string, string>();
  private faces: FontFace[] = [];
  private defaultTab = 48;
  private legacyTables = true;
  private tableDepth = 0;
  /** Inside tables, the size and alignment of the default paragraph style give way to the table style. */
  private tableStyleWins = true;
  private evenAndOdd = false;
  private noteNumber = '';
  /** Which part of the page is being rendered; floating objects are positioned against it. */
  private region: 'body' | 'header' = 'body';
  private static instances = 0;
  private readonly instance = ++DocxRenderer.instances;

  constructor(
    private readonly zip: ZipArchive,
    private readonly host: DocxHost,
  ) {}

  async render(): Promise<DocxDocument> {
    const rootRelationships = await readRelationships(this.zip, '');
    const mainPath =
      [...rootRelationships.values()].find((rel) => rel.type === 'officeDocument')?.target ??
      'word/document.xml';
    const document = await requireXmlPart(this.zip, mainPath, 'Word document');
    this.relationships = await readRelationships(this.zip, mainPath);
    const mainRelationships = this.relationships;
    const partOfType = (type: string, fallback: string): string =>
      [...mainRelationships.values()].find((rel) => rel.type === type && !rel.external)?.target ??
      fallback;

    const numberingPath = partOfType('numbering', 'word/numbering.xml');
    const footnotesPath = partOfType('footnotes', 'word/footnotes.xml');
    const endnotesPath = partOfType('endnotes', 'word/endnotes.xml');
    const fontTablePath = partOfType('fontTable', 'word/fontTable.xml');
    const [theme, styles, numbering, footnotes, endnotes, settings, fontTable] = await Promise.all([
      loadTheme(this.zip, partOfType('theme', 'word/theme/theme1.xml')),
      readXmlPart(this.zip, partOfType('styles', 'word/styles.xml')),
      readXmlPart(this.zip, numberingPath),
      readXmlPart(this.zip, footnotesPath),
      readXmlPart(this.zip, endnotesPath),
      readXmlPart(this.zip, partOfType('settings', 'word/settings.xml')),
      readXmlPart(this.zip, fontTablePath),
    ]);
    this.theme = theme;
    this.parser = new PropertyParser(theme);
    this.loadSettings(settings);
    this.loadStyles(styles);
    await this.loadFonts(fontTable, fontTablePath);
    this.loadNumbering(numbering, await readRelationships(this.zip, numberingPath));
    this.footnotes = {
      root: footnotes,
      relationships: await readRelationships(this.zip, footnotesPath),
    };
    this.endnotes = {
      root: endnotes,
      relationships: await readRelationships(this.zip, endnotesPath),
    };

    const body = child(document, 'body');
    // Sections: each one ends with the paragraph that carries its properties; the last is on the body.
    const sectionElements = childElements(body, 'p').flatMap((paragraph) => {
      const properties = find(paragraph, 'pPr', 'sectPr');
      return properties ? [properties] : [];
    });
    sectionElements.push(
      child(body, 'sectPr') ?? { local: 'sectPr', name: 'w:sectPr', attrs: {}, children: [] },
    );
    let previous: PageSetup | undefined;
    for (const element of sectionElements) {
      // The geometry is registered first: floating objects in headers are positioned against it.
      const setup = this.sectionGeometry(element, previous);
      this.sections.push(setup);
      this.sectionIndex = this.sections.length - 1;
      await this.loadHeaders(element, setup);
      previous = setup;
    }
    this.sectionIndex = 0;
    this.relationships = mainRelationships;

    const wrapper = h('div', {
      class: 'av-doc',
      style: {
        'font-size': `${this.defaultRun.size ?? 10}pt`,
        'font-family': this.fontCss(this.defaultRun.font ?? 'Times New Roman'),
      },
    });
    const separatorOf = (
      part: { root: XmlElement | null; relationships: Map<string, Relationship> },
      tag: string,
    ): HTMLElement => {
      const note = childElements(part.root, tag).find(
        (element) => attr(element, 'w:type') === 'separator',
      );
      const rendered = note ? this.renderNote(part, tag, attr(note, 'w:id') ?? '', '') : null;
      const separator =
        rendered ??
        h(
          'div',
          null,
          h('span', {
            style: { display: 'inline-block', width: '2in', 'border-top': '1px solid #000' },
          }),
        );
      separator.setAttribute('data-av-sep', '');
      return separator;
    };
    const separator = separatorOf(this.footnotes, 'footnote');
    const first = createPage(this.sections[0]!, 1, {
      paged: false,
      firstOfSection: true,
      evenAndOdd: this.evenAndOdd,
      separator,
    });
    wrapper.appendChild(first.page);
    this.renderBody(childElements(body), first.body);

    // Footnotes are rendered after the body; until pagination they all sit on the single page.
    const notes = new Map<string, HTMLElement>();
    this.usedFootnotes.forEach((id, index) => {
      const note = this.renderNote(this.footnotes, 'footnote', id, String(index + 1));
      if (!note) return;
      notes.set(id, note);
      first.notes.appendChild(note);
      first.notes.hidden = false;
    });
    if (this.usedEndnotes.length > 0) {
      const section = h('div', { class: 'av-doc-endnotes' }, separatorOf(this.endnotes, 'endnote'));
      this.usedEndnotes.forEach((id, index) => {
        const note = this.renderNote(
          this.endnotes,
          'endnote',
          id,
          formatCounter(index + 1, 'lowerRoman'),
        );
        if (note) section.appendChild(note);
      });
      first.body.appendChild(section);
    }
    fillFields(first.page, 1, 1);
    await Promise.all(this.tasks);
    for (const { image, key } of this.pendingPictures) {
      const url = this.pictureUrls.get(key);
      if (url) image.src = url;
    }

    wrapper.addEventListener('click', (event) => {
      const link = (event.target as Element | null)?.closest?.('[data-av-anchor]');
      if (!link) return;
      event.preventDefault();
      const name = link.getAttribute('data-av-anchor');
      const target = Array.from(wrapper.querySelectorAll('[data-av-bookmark]')).find(
        (bookmark) => bookmark.getAttribute('data-av-bookmark') === name,
      );
      target?.scrollIntoView({ block: 'start' });
    });

    let pages: HTMLElement[] = [first.page];
    let laidOut = false;
    return {
      element: wrapper,
      layout: async (): Promise<void> => {
        if (laidOut || !wrapper.isConnected) return;
        laidOut = true;
        const fonts = (globalThis.document as Document & { fonts?: FontFaceSet }).fonts;
        if (fonts) {
          await Promise.race([
            fonts.ready.catch(() => undefined),
            new Promise((resolve) => setTimeout(resolve, 1500)),
          ]);
        }
        if (!wrapper.isConnected) return;
        layoutTabs(wrapper);
        const blocks = Array.from(first.body.children) as HTMLElement[];
        const paginator = new Paginator(wrapper, this.sections, notes, this.evenAndOdd, separator);
        paginator.run(blocks);
        pages = paginator.pages.map((parts) => parts.page);
      },
      pageCount: () => pages.length,
      page: (index) => pages[index] ?? null,
      destroy: (): void => {
        const fonts = (globalThis.document as Document & { fonts?: FontFaceSet }).fonts;
        for (const face of this.faces) fonts?.delete(face);
        this.faces = [];
      },
    };
  }

  private loadSettings(root: XmlElement | null): void {
    const tab = num(attr(child(root, 'defaultTabStop'), 'w:val'));
    if (tab && tab > 0) this.defaultTab = twipsToPx(tab);
    this.evenAndOdd = isOn(child(root, 'evenAndOddHeaders'));
    for (const setting of childElements(child(root, 'compat'), 'compatSetting')) {
      if (attr(setting, 'w:name') === 'compatibilityMode') {
        this.legacyTables = (num(attr(setting, 'w:val')) ?? 12) < 15;
      } else if (attr(setting, 'w:name') === 'overrideTableStyleFontSizeAndJustification') {
        this.tableStyleWins = !isOn(setting);
      }
    }
  }

  private loadStyles(root: XmlElement | null): void {
    if (!root) return;
    this.defaultRun = this.parser.run(find(root, 'docDefaults', 'rPrDefault', 'rPr'));
    this.defaultPara = this.parser.para(find(root, 'docDefaults', 'pPrDefault', 'pPr'));
    for (const element of childElements(root, 'style')) {
      const style = this.parser.style(element);
      if (!style) continue;
      this.styles.set(style.id, style);
      if (attr(element, 'w:default') !== undefined && isOn(element, 'w:default')) {
        if (style.type === 'paragraph') this.defaultParagraphStyle = style.id;
        else if (style.type === 'table') this.defaultTableStyle = style.id;
      }
    }
  }

  /** Registers the fonts embedded in the document under private names. */
  private async loadFonts(root: XmlElement | null, path: string): Promise<void> {
    if (!root) return;
    const relationships = await readRelationships(this.zip, path);
    const fonts = (globalThis.document as Document & { fonts?: FontFaceSet }).fonts;
    const loads: Promise<void>[] = [];
    for (const font of childElements(root, 'font')) {
      const name = safeFontName(attr(font, 'w:name'));
      if (!name) continue;
      const family = attr(child(font, 'family'), 'w:val');
      const pitch = attr(child(font, 'pitch'), 'w:val');
      this.fontFamilies.set(
        name.toLowerCase(),
        pitch === 'fixed' || family === 'modern'
          ? 'monospace'
          : family === 'roman'
            ? 'serif'
            : family === 'script'
              ? 'cursive'
              : 'sans-serif',
      );
      if (!fonts || typeof FontFace === 'undefined') continue;
      const variants: [string, FontFaceDescriptors][] = [
        ['embedRegular', {}],
        ['embedBold', { weight: 'bold' }],
        ['embedItalic', { style: 'italic' }],
        ['embedBoldItalic', { weight: 'bold', style: 'italic' }],
      ];
      for (const [tag, descriptors] of variants) {
        const embedded = child(font, tag);
        const relationship = relationships.get(attr(embedded, 'r:id') ?? '');
        if (!embedded || !relationship || relationship.external) continue;
        const alias = `av${this.instance} ${name}`;
        const key = attr(embedded, 'w:fontKey');
        loads.push(
          (async (): Promise<void> => {
            const data = await this.zip.read(relationship.target);
            if (!data) return;
            const bytes = key ? deobfuscateFont(data, key) : data;
            try {
              const face = new FontFace(alias, toBlobPart(bytes), descriptors);
              await face.load();
              fonts.add(face);
              this.faces.push(face);
              this.fontAliases.set(name.toLowerCase(), alias);
            } catch {
              // A font the browser rejects is skipped; the text falls back to an installed font.
            }
          })(),
        );
      }
    }
    await Promise.all(loads);
  }

  private fontCss(name: string): string {
    const key = name.toLowerCase();
    const families: string[] = [];
    const alias = this.fontAliases.get(key);
    if (alias) families.push(`"${alias}"`);
    families.push(`"${name}"`);
    const fallback = FONT_FALLBACKS[key];
    if (fallback) families.push(`"${fallback}"`);
    families.push(
      this.fontFamilies.get(key) ??
        (MONO.test(name) ? 'monospace' : SERIF.test(name) ? 'serif' : 'sans-serif'),
    );
    return families.join(', ');
  }

  private loadNumbering(root: XmlElement | null, relationships: Map<string, Relationship>): void {
    if (!root) return;
    // Picture bullets: the image is loaded once and shared by every list item.
    const pictures = new Map<string, string>();
    for (const bullet of childElements(root, 'numPicBullet')) {
      const id = attr(bullet, 'w:numPicBulletId');
      const image = descendants(bullet, 'imagedata')[0] ?? descendants(bullet, 'blip')[0] ?? null;
      const relationship = relationships.get(attr(image, 'r:id') ?? attr(image, 'r:embed') ?? '');
      const mimeType = relationship && !relationship.external && imageMimeType(relationship.target);
      if (id === undefined || !relationship || !mimeType) continue;
      const key = `picture:${id}`;
      pictures.set(id, key);
      this.tasks.push(
        this.zip.read(relationship.target).then((data) => {
          if (data) {
            this.pictureUrls.set(
              key,
              this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType })),
            );
          }
        }),
      );
    }

    const abstracts = new Map<string, Map<number, NumberingLevel>>();
    const parseLevels = (parent: XmlElement, into: Map<number, NumberingLevel>): void => {
      for (const level of childElements(parent, 'lvl')) {
        const index = num(attr(level, 'w:ilvl')) ?? 0;
        into.set(index, {
          start: num(attr(child(level, 'start'), 'w:val')) ?? 1,
          format: attr(child(level, 'numFmt'), 'w:val') ?? 'decimal',
          text: attr(child(level, 'lvlText'), 'w:val') ?? '',
          suffix: attr(child(level, 'suff'), 'w:val') ?? 'tab',
          align: attr(child(level, 'lvlJc'), 'w:val') ?? 'left',
          para: this.parser.para(child(level, 'pPr')),
          run: this.parser.run(child(level, 'rPr')),
          picture: pictures.get(attr(child(level, 'lvlPicBulletId'), 'w:val') ?? ''),
        });
      }
    };
    for (const abstract of childElements(root, 'abstractNum')) {
      const levels = new Map<number, NumberingLevel>();
      parseLevels(abstract, levels);
      abstracts.set(attr(abstract, 'w:abstractNumId') ?? '', levels);
    }
    for (const list of childElements(root, 'num')) {
      const id = attr(list, 'w:numId');
      const abstractId = attr(child(list, 'abstractNumId'), 'w:val') ?? '';
      if (id === undefined) continue;
      const levels = new Map(abstracts.get(abstractId) ?? []);
      const restarts = new Map<number, number>();
      for (const override of childElements(list, 'lvlOverride')) {
        const index = num(attr(override, 'w:ilvl')) ?? 0;
        parseLevels(override, levels);
        const start = num(attr(child(override, 'startOverride'), 'w:val'));
        const level = levels.get(index);
        if (start !== undefined) {
          restarts.set(index, start);
          if (level) levels.set(index, { ...level, start });
        }
      }
      this.numbering.set(id, { levels, abstractId, restarts });
    }
  }

  private pictureUrls = new Map<string, string>();

  private async loadHeaders(element: XmlElement, setup: PageSetup): Promise<void> {
    const { headers, footers } = setup;
    for (const reference of childElements(element)) {
      if (reference.local !== 'headerReference' && reference.local !== 'footerReference') continue;
      const type = attr(reference, 'w:type') ?? 'default';
      const relationship = this.relationships.get(attr(reference, 'r:id') ?? '');
      if (!relationship || relationship.external) continue;
      if (type !== 'default' && type !== 'first' && type !== 'even') continue;
      const rendered = await this.renderPart(relationship.target);
      if (!rendered) continue;
      if (reference.local === 'headerReference') headers[type] = rendered;
      else footers[type] = rendered;
    }
  }

  private sectionGeometry(element: XmlElement, previous: PageSetup | undefined): PageSetup {
    const size = child(element, 'pgSz');
    const margins = child(element, 'pgMar');
    const margin = (side: string, fallback: number): number =>
      twipsToPx(Math.max(0, Math.abs(num(attr(margins, `w:${side}`)) ?? fallback)));
    const headers: HeaderSet = { ...previous?.headers };
    const footers: HeaderSet = { ...previous?.footers };
    return {
      width: twipsToPx(num(attr(size, 'w:w')) ?? 12240),
      height: twipsToPx(num(attr(size, 'w:h')) ?? 15840),
      top: margin('top', 1440),
      right: margin('right', 1440),
      bottom: margin('bottom', 1440),
      left: margin('left', 1440),
      header: margin('header', 720),
      footer: margin('footer', 720),
      continuous: attr(child(element, 'type'), 'w:val') === 'continuous',
      titlePage: isOn(child(element, 'titlePg')),
      headers,
      footers,
    };
  }

  /** Renders a header or footer part with its own relationships. */
  private async renderPart(path: string): Promise<HTMLElement | null> {
    const root = await readXmlPart(this.zip, path);
    if (!root) return null;
    const relationships = await readRelationships(this.zip, path);
    const outer = this.relationships;
    this.relationships = relationships;
    this.region = 'header';
    const container = h('div');
    try {
      this.renderBlocks(childElements(root), container);
    } finally {
      this.relationships = outer;
      this.region = 'body';
    }
    return container;
  }

  private renderNote(
    part: { root: XmlElement | null; relationships: Map<string, Relationship> },
    tag: string,
    id: string,
    label: string,
  ): HTMLElement | null {
    const note = childElements(part.root, tag).find((element) => attr(element, 'w:id') === id);
    if (!note) return null;
    const outer = this.relationships;
    this.relationships = part.relationships;
    this.noteNumber = label;
    const container = h('div', { class: 'av-doc-note' });
    try {
      this.renderBlocks(childElements(note), container);
    } finally {
      this.relationships = outer;
      this.noteNumber = '';
    }
    return container;
  }

  /** Style chain from the root ancestor down to `styleId`. */
  private styleChain(styleId: string | undefined): StyleDef[] {
    const chain: StyleDef[] = [];
    const seen = new Set<string>();
    let current = styleId;
    while (current && !seen.has(current)) {
      seen.add(current);
      const style = this.styles.get(current);
      if (!style) break;
      chain.unshift(style);
      current = style.basedOn;
    }
    return chain;
  }

  private nextLabel(
    numId: string,
    levelIndex: number,
  ): { text: string; level: NumberingLevel } | null {
    const definition = this.numbering.get(numId);
    const level = definition?.levels.get(levelIndex);
    if (!definition || !level) return null;
    // Lists based on the same abstract definition continue each other's numbering.
    const key = definition.abstractId || numId;
    let counters = this.counters.get(key);
    if (!counters) this.counters.set(key, (counters = []));
    if (!this.usedLists.has(numId)) {
      this.usedLists.add(numId);
      for (const [index, start] of definition.restarts) {
        counters[index] = start - 1;
        counters.length = Math.max(index + 1, 0);
      }
    }
    counters[levelIndex] = (counters[levelIndex] ?? level.start - 1) + 1;
    counters.length = levelIndex + 1;
    if (level.format === 'bullet') return { text: bulletGlyph(level.text, level.run.font), level };
    const text = level.text.replace(/%(\d)/g, (_match, digit: string) => {
      const index = Number(digit) - 1;
      const other = definition.levels.get(index);
      const value = counters[index] ?? other?.start ?? 1;
      return formatCounter(value, other?.format ?? 'decimal');
    });
    return { text, level };
  }

  /** Body content: tracks sections so that pagination knows where each one starts. */
  private renderBody(elements: XmlElement[], parent: HTMLElement): void {
    let startsSection = false;
    for (const element of elements) {
      const before = parent.childElementCount;
      this.renderBlocks([element], parent);
      if (startsSection) {
        const firstNew = parent.children[before];
        if (firstNew) {
          firstNew.setAttribute('data-av-sect', String(this.sectionIndex));
          startsSection = false;
        }
      }
      if (element.local === 'p' && find(element, 'pPr', 'sectPr')) {
        this.sectionIndex = Math.min(this.sectionIndex + 1, this.sections.length - 1);
        startsSection = true;
      }
    }
  }

  private lastParagraph: { style: string | undefined; element: HTMLElement } | null = null;

  private renderBlocks(elements: XmlElement[], parent: HTMLElement): void {
    for (const element of elements) {
      if (element.local === 'p') {
        parent.appendChild(this.renderParagraph(element));
      } else if (element.local === 'tbl') {
        this.lastParagraph = null;
        parent.appendChild(this.renderTable(element));
        this.lastParagraph = null;
      } else if (element.local === 'sdt') {
        this.renderBlocks(childElements(child(element, 'sdtContent')), parent);
      } else if (
        element.local === 'ins' ||
        element.local === 'customXml' ||
        element.local === 'smartTag'
      ) {
        this.renderBlocks(childElements(element), parent);
      }
    }
  }

  private section(): PageSetup {
    return this.sections[this.sectionIndex] ?? this.sections[0]!;
  }

  private renderParagraph(paragraph: XmlElement): HTMLElement {
    const pPr = child(paragraph, 'pPr');
    const styleId = attr(child(pPr, 'pStyle'), 'w:val') ?? this.defaultParagraphStyle;
    const chain = this.styleChain(styleId);
    const direct = this.parser.para(pPr);
    const context = this.cells[this.cells.length - 1];
    const yields = (style: StyleDef): boolean =>
      context !== undefined && this.tableStyleWins && style.id === this.defaultParagraphStyle;
    const layers = [
      this.defaultPara,
      context?.para,
      ...chain.map((style): ParaProps => {
        if (!yields(style) || style.para.align === undefined) return style.para;
        const rest = { ...style.para };
        delete rest.align;
        return rest;
      }),
      direct,
    ];
    const props = merge<ParaProps>(...layers);
    const styleRun = merge<RunProps>(
      this.defaultRun,
      context?.run,
      ...chain.map((style): RunProps => {
        if (!yields(style) || style.run.size === undefined) return style.run;
        const rest = { ...style.run };
        delete rest.size;
        return rest;
      }),
    );
    const markRun = merge<RunProps>(styleRun, this.parser.run(child(pPr, 'rPr')));
    const tabs = mergeTabs(layers.map((layer) => layer?.tabs));

    // Headings are found by outline level or by the built-in style names.
    let headingLevel =
      props.outlineLevel !== undefined && props.outlineLevel < 9 ? props.outlineLevel + 1 : 0;
    for (const style of chain) {
      const match = /^heading (\d)$/i.exec(style.name);
      if (match) headingLevel = Number(match[1]);
    }
    const tag = headingLevel >= 1 && headingLevel <= 6 ? (`h${headingLevel}` as 'h1') : 'p';

    let label: { text: string; level: NumberingLevel } | null = null;
    if (props.numId !== undefined && props.numId !== '0') {
      label = this.nextLabel(props.numId, props.level ?? 0);
      if (label) {
        // Level indents apply unless the paragraph sets its own.
        if (direct.indentLeft === undefined && label.level.para.indentLeft !== undefined) {
          props.indentLeft = label.level.para.indentLeft;
        }
        if (direct.firstLine === undefined && label.level.para.firstLine !== undefined) {
          props.firstLine = label.level.para.firstLine;
        }
      }
    }

    const align =
      props.align === 'center'
        ? 'center'
        : props.align === 'right' || props.align === 'end'
          ? 'right'
          : props.align === 'both' || props.align === 'distribute'
            ? 'justify'
            : props.align === 'left' || props.align === 'start'
              ? 'left'
              : undefined;
    const font = styleRun.font ?? this.defaultRun.font;
    let lineHeight: string | undefined;
    if (props.line) {
      if (props.lineRule === 'auto' || props.lineRule === undefined) {
        lineHeight = this.lineMultiple(props.line, font);
      } else {
        lineHeight = `${(props.line / 20).toFixed(2)}pt`;
      }
    }
    const element = h(tag, {
      style: {
        'text-align': align,
        'margin-left': twipsPx(props.indentLeft),
        'margin-right': twipsPx(props.indentRight),
        'text-indent': twipsPx(props.firstLine),
        'margin-top': twipsPx(props.spaceBefore),
        'margin-bottom': twipsPx(props.spaceAfter),
        'line-height': lineHeight,
        'background-color': props.background,
        'white-space': 'pre-wrap',
        direction: props.rtl ? 'rtl' : undefined,
        'font-family': font ? this.fontCss(font) : undefined,
      },
    });
    if (props.pageBreakBefore) {
      element.style.setProperty('break-before', 'page');
      element.setAttribute('data-av-break', 'before');
    }
    if (props.keepNext) element.setAttribute('data-av-kn', '');
    if (props.keepLines) element.setAttribute('data-av-kl', '');
    if (props.widowControl === false) element.setAttribute('data-av-nw', '');
    this.applyParagraphBorders(element, props.borders);

    // Paragraphs of the same style that ask for it have no space between them.
    const previous = this.lastParagraph;
    if (props.contextualSpacing && previous && previous.style === styleId) {
      element.style.marginTop = '0';
      previous.element.style.marginBottom = '0';
    }
    this.lastParagraph = props.contextualSpacing ? { style: styleId, element } : null;

    if (props.frame) this.applyFrame(element, props);

    const outer = this.paragraph;
    const state: ParagraphState = {
      sizes: [],
      hasTabs: false,
      font,
      line: props.line,
      lineRule: props.lineRule,
    };
    this.paragraph = state;

    if (label) element.appendChild(this.renderMarker(label, props, markRun));
    this.renderInline(childElements(paragraph), element, styleRun);
    this.paragraph = outer;

    // The line height of a paragraph follows its text, so the base size is that of its smallest run.
    const size = state.sizes.length > 0 ? Math.min(...state.sizes) : markRun.size;
    if (size) element.style.fontSize = `${size}pt`;
    const hasContent = state.sizes.length > 0 || element.querySelector('img,.av-doc-box') !== null;
    if (!hasContent && !label) element.appendChild(h('br'));
    if (state.hasTabs || tabs.length > 0) {
      element.setAttribute(
        'data-av-tabs',
        tabs
          .map((stop) => `${Math.round(stop.pos * 100) / 100}:${stop.align}:${stop.leader}`)
          .join(';'),
      );
      element.setAttribute(
        'data-av-ind',
        [twipsToPx(props.indentLeft ?? 0), twipsToPx(props.firstLine ?? 0), this.defaultTab].join(
          ',',
        ),
      );
    }
    return element;
  }

  private lineMultiple(line: number, font: string | undefined): string {
    return ((line / 240) * lineFactor(font)).toFixed(3);
  }

  private applyParagraphBorders(element: HTMLElement, borders: BorderSet | undefined): void {
    if (!borders) return;
    for (const side of SIDES) {
      const border = borders[side];
      if (!border || border.style === 'none') continue;
      element.style.setProperty(`border-${side}`, borderCss(border)!);
      if (border.space) element.style.setProperty(`padding-${side}`, px(border.space)!);
      // Side borders are drawn outside the text area, so the text keeps its position.
      if (side === 'left' || side === 'right') {
        const current = parseFloat(element.style.getPropertyValue(`margin-${side}`)) || 0;
        element.style.setProperty(`margin-${side}`, `${current - border.space - border.width}px`);
      }
    }
  }

  /** Drop caps and simple text frames become floats, so the following text wraps around them. */
  private applyFrame(element: HTMLElement, props: ParaProps): void {
    const frame = props.frame!;
    if (frame.dropCap === 'drop' || frame.dropCap === 'margin') {
      setStyles(element, {
        float: 'left',
        'margin-right': px(frame.hSpace) ?? '2px',
        'margin-top': '0',
        'margin-bottom': '0',
        'margin-left': frame.dropCap === 'margin' ? '-1.2em' : undefined,
        // Just under the height of the lines it spans, so the next line is not pushed aside.
        height:
          props.line && props.lineRule === 'exact'
            ? `${(twipsToPx(props.line) - 1).toFixed(2)}px`
            : undefined,
      });
      element.setAttribute('data-av-kn', '');
    } else if (frame.xAlign === 'right' || frame.xAlign === 'left') {
      setStyles(element, {
        float: frame.xAlign,
        width: frame.width ? twipsPx(frame.width) : undefined,
        [frame.xAlign === 'right' ? 'margin-left' : 'margin-right']: px(frame.hSpace),
      });
    }
  }

  private renderMarker(
    label: { text: string; level: NumberingLevel },
    props: ParaProps,
    markRun: RunProps,
  ): HTMLElement {
    const { level } = label;
    const hanging = props.firstLine !== undefined && props.firstLine < 0 ? -props.firstLine : 0;
    const marker = h('span', {
      class: 'av-doc-marker',
      style: {
        display: 'inline-block',
        'min-width': hanging ? twipsPx(hanging) : undefined,
        'padding-right': hanging || level.suffix === 'nothing' ? undefined : '0.5em',
        'text-indent': '0',
        'white-space': 'nowrap',
      },
    });
    const run = merge<RunProps>(markRun, level.run, { underline: 'none' });
    const pictureUrl = level.picture;
    let content: HTMLElement;
    if (pictureUrl) {
      const image = h('img', {
        attrs: { alt: '' },
        style: { height: '0.7em', width: 'auto', 'vertical-align': 'baseline' },
      });
      this.pendingPictures.push({ image, key: pictureUrl });
      content = image;
    } else {
      content = h('span', { text: label.text });
      // Bullets mapped from a symbol font to Unicode are drawn in a common font with similar glyphs.
      const symbolic = level.format === 'bullet' && label.text !== level.text;
      setStyles(content, this.runCss(symbolic ? { ...run, font: 'Arial' } : run));
    }
    if (level.align === 'right') {
      // A right-aligned number ends where the text indent begins, growing to the left.
      setStyles(content, { display: 'inline-block', transform: 'translateX(-100%)' });
    } else if (level.align === 'center') {
      setStyles(content, { display: 'inline-block', transform: 'translateX(-50%)' });
    }
    marker.appendChild(content);
    if (level.suffix === 'tab' && level.align === 'left')
      marker.setAttribute('data-av-suff', 'tab');
    if (this.paragraph) this.paragraph.hasTabs = true;
    if (this.paragraph && run.size) this.paragraph.sizes.push(run.size);
    return marker;
  }

  private pendingPictures: { image: HTMLImageElement; key: string }[] = [];

  private runCss(props: RunProps): Record<string, string | undefined> {
    const underlined = props.underline !== undefined && props.underline !== 'none';
    const struck = props.strike || props.doubleStrike;
    const decoration = [underlined && 'underline', struck && 'line-through']
      .filter(Boolean)
      .join(' ');
    const decorationStyle =
      underlined && props.underline !== 'solid'
        ? props.underline
        : props.doubleStrike
          ? 'double'
          : undefined;
    const subOrSuper = props.vertAlign === 'superscript' || props.vertAlign === 'subscript';
    const state = this.paragraph;
    let lineHeight: string | undefined;
    if (state?.line && props.size) {
      if (state.lineRule === 'atLeast') {
        lineHeight = `max(${(state.line / 20).toFixed(2)}pt, ${lineFactor(props.font).toFixed(3)}em)`;
      } else if (
        (state.lineRule === 'auto' || state.lineRule === undefined) &&
        props.font !== state.font
      ) {
        lineHeight = this.lineMultiple(state.line, props.font);
      }
    }
    return {
      'font-weight': props.bold ? 'bold' : props.bold === false ? 'normal' : undefined,
      'font-style': props.italic ? 'italic' : props.italic === false ? 'normal' : undefined,
      'text-decoration': decoration || undefined,
      'text-decoration-style': decoration ? decorationStyle : undefined,
      'text-decoration-color': underlined ? props.underlineColor : undefined,
      color: props.color,
      'background-color': props.highlight,
      'font-size': props.size
        ? `${subOrSuper ? Math.round(props.size * 65) / 100 : props.size}pt`
        : undefined,
      'font-family': props.font ? this.fontCss(props.font) : undefined,
      'vertical-align':
        props.vertAlign === 'superscript'
          ? 'super'
          : props.vertAlign === 'subscript'
            ? 'sub'
            : undefined,
      'text-transform': props.caps ? 'uppercase' : undefined,
      'font-variant': props.smallCaps ? 'small-caps' : undefined,
      'letter-spacing': props.spacing ? px(props.spacing) : undefined,
      position: props.position ? 'relative' : undefined,
      top: props.position ? `${-props.position}pt` : undefined,
      border: props.border ? borderCss(props.border) : undefined,
      'text-shadow': props.shadow ? '1px 1px 1px rgba(0,0,0,.4)' : undefined,
      // Raised and lowered text must not make its line taller.
      'line-height': subOrSuper ? '0' : lineHeight,
    };
  }

  /** True while the instruction of a field, or a replaced field result, is being read. */
  private suppressed(): boolean {
    return this.fields.some((field) => field.phase === 'instruction' || field.skip);
  }

  private makeLink(href: string | null, anchor: string | null): HTMLElement {
    if (href) return externalLink(href);
    if (anchor) return h('a', { attrs: { href: '#', 'data-av-anchor': anchor } });
    return h('span');
  }

  /** Where the content of a run goes: the paragraph, or the link of an enclosing HYPERLINK field. */
  private runTarget(parent: HTMLElement): HTMLElement {
    for (let i = this.fields.length - 1; i >= 0; i--) {
      const field = this.fields[i]!;
      if (field.phase !== 'result' || (!field.href && !field.anchor)) continue;
      if (field.linkParent !== parent || !field.link) {
        field.link = this.makeLink(field.href, field.anchor);
        field.linkParent = parent;
        parent.appendChild(field.link);
      }
      return field.link;
    }
    return parent;
  }

  private fieldSpan(kind: 'page' | 'pages'): HTMLElement {
    return h('span', { attrs: { 'data-av-field': kind }, text: '1' });
  }

  private static fieldKind(instruction: string): 'page' | 'pages' | null {
    const name = /^\s*(\w+)/.exec(instruction)?.[1]?.toUpperCase();
    if (name === 'PAGE') return 'page';
    if (name === 'NUMPAGES' || name === 'SECTIONPAGES') return 'pages';
    return null;
  }

  private renderInline(elements: XmlElement[], parent: HTMLElement, inherited: RunProps): void {
    for (const element of elements) {
      switch (element.local) {
        case 'r':
          this.renderRun(element, parent, inherited);
          break;
        case 'hyperlink': {
          const relationship = this.relationships.get(attr(element, 'r:id') ?? '');
          const href = relationship?.external ? safeLinkUrl(relationship.target) : null;
          const link = this.makeLink(href, attr(element, 'w:anchor') ?? null);
          this.renderInline(childElements(element), link, inherited);
          if (link.firstChild) parent.appendChild(link);
          break;
        }
        case 'bookmarkStart': {
          const name = attr(element, 'w:name');
          if (name && name !== '_GoBack') {
            parent.appendChild(h('span', { attrs: { 'data-av-bookmark': name } }));
          }
          break;
        }
        case 'fldSimple': {
          const kind = DocxRenderer.fieldKind(attr(element, 'w:instr') ?? '');
          if (kind) {
            if (this.suppressed()) break;
            const span = this.fieldSpan(kind);
            const firstRun = child(element, 'r');
            const props = merge<RunProps>(inherited, this.parser.run(child(firstRun, 'rPr')));
            setStyles(span, this.runCss(props));
            if (this.paragraph && props.size) this.paragraph.sizes.push(props.size);
            this.runTarget(parent).appendChild(span);
          } else {
            this.renderInline(childElements(element), parent, inherited);
          }
          break;
        }
        case 'oMath':
        case 'oMathPara': {
          // Equations are shown in linear form.
          const text = descendants(element, 't').map(textOf).join('');
          if (text && !this.suppressed()) {
            parent.appendChild(
              h('span', {
                text,
                style: { 'font-family': '"Cambria Math", Cambria, serif', 'font-style': 'italic' },
              }),
            );
            this.paragraph?.sizes.push(inherited.size ?? this.defaultRun.size ?? 10);
          }
          break;
        }
        case 'sdt':
          this.renderInline(childElements(child(element, 'sdtContent')), parent, inherited);
          break;
        case 'ins':
        case 'smartTag':
        case 'customXml':
        case 'moveTo':
          this.renderInline(childElements(element), parent, inherited);
          break;
        // w:del and w:moveFrom hold tracked deletions and are intentionally not shown.
      }
    }
  }

  private renderRun(run: XmlElement, parent: HTMLElement, inherited: RunProps): void {
    const rPr = child(run, 'rPr');
    let characterStyle = this.styleChain(attr(child(rPr, 'rStyle'), 'w:val'));
    // Word shows the entries of a table of contents in the TOC styles, not as blue links.
    if (
      characterStyle.some((style) => /^hyperlink$/i.test(style.name)) &&
      this.fields.some((field) => field.instruction.trim().toUpperCase().startsWith('TOC'))
    ) {
      characterStyle = [];
    }
    const props = merge<RunProps>(
      inherited,
      ...characterStyle.map((style) => style.run),
      this.parser.run(rPr),
    );

    const span = h('span');
    const text = (value: string): void => {
      if (value && !this.suppressed()) span.appendChild(document.createTextNode(value));
    };
    for (const node of childElements(run)) {
      switch (node.local) {
        case 'fldChar': {
          const type = attr(node, 'w:fldCharType');
          const field = this.fields[this.fields.length - 1];
          if (type === 'begin') {
            this.fields.push({
              instruction: '',
              phase: 'instruction',
              skip: false,
              href: null,
              anchor: null,
              link: null,
              linkParent: null,
            });
          } else if (type === 'separate' && field) {
            field.phase = 'result';
            const kind = DocxRenderer.fieldKind(field.instruction);
            if (kind) {
              if (!this.suppressed()) span.appendChild(this.fieldSpan(kind));
              field.skip = true;
            } else {
              const link = /^\s*HYPERLINK\s+(\\l\s+)?"([^"]*)"/i.exec(field.instruction);
              if (link?.[1]) field.anchor = link[2] ?? null;
              else if (link) field.href = safeLinkUrl(link[2]);
              const local = /\\l\s+"([^"]*)"/i.exec(field.instruction);
              if (link && !link[1] && local && !field.href) field.anchor = local[1] ?? null;
            }
          } else if (type === 'end' && field) {
            this.fields.pop();
            const kind = DocxRenderer.fieldKind(field.instruction);
            if (kind && field.phase === 'instruction' && !this.suppressed()) {
              span.appendChild(this.fieldSpan(kind));
            }
          }
          break;
        }
        case 'instrText': {
          const field = this.fields[this.fields.length - 1];
          if (field?.phase === 'instruction') field.instruction += textOf(node);
          break;
        }
        case 't':
          text(textOf(node));
          break;
        case 'tab':
        case 'ptab': {
          if (this.suppressed()) break;
          const tab = h('span', { class: 'av-doc-tab' }, h('span', { text: '\t' }));
          if (node.local === 'ptab') {
            tab.setAttribute('data-av-ptab', attr(node, 'w:alignment') ?? 'left');
            tab.setAttribute('data-av-leader', attr(node, 'w:leader') ?? 'none');
          }
          span.appendChild(tab);
          if (this.paragraph) this.paragraph.hasTabs = true;
          break;
        }
        case 'br':
          if (this.suppressed()) break;
          if (attr(node, 'w:type') === 'page') {
            span.appendChild(
              h('span', {
                attrs: { 'data-av-pb': '' },
                style: {
                  display: 'block',
                  'break-after': 'page',
                  'border-top': '1px dashed #c8c8c8',
                  margin: '16px 0',
                },
              }),
            );
          } else {
            span.appendChild(h('br'));
          }
          break;
        case 'cr':
          if (!this.suppressed()) span.appendChild(h('br'));
          break;
        case 'noBreakHyphen':
          text(NON_BREAKING_HYPHEN);
          break;
        case 'softHyphen':
          text(SOFT_HYPHEN);
          break;
        case 'sym': {
          const code = parseInt(attr(node, 'w:char') ?? '', 16);
          if (Number.isFinite(code)) {
            text(bulletGlyph(String.fromCodePoint(code), attr(node, 'w:font') ?? 'Symbol'));
          }
          break;
        }
        case 'footnoteReference':
        case 'endnoteReference': {
          const id = attr(node, 'w:id');
          if (id === undefined || this.suppressed()) break;
          const foot = node.local === 'footnoteReference';
          const part = foot ? this.footnotes : this.endnotes;
          const tag = foot ? 'footnote' : 'endnote';
          if (!childElements(part.root, tag).some((note) => attr(note, 'w:id') === id)) break;
          const used = foot ? this.usedFootnotes : this.usedEndnotes;
          if (!used.includes(id)) used.push(id);
          const number = used.indexOf(id) + 1;
          const reference = h('span', {
            text: foot ? String(number) : formatCounter(number, 'lowerRoman'),
            attrs: foot ? { 'data-av-fn': id } : {},
          });
          if (!props.vertAlign)
            setStyles(reference, { 'vertical-align': 'super', 'font-size': '65%' });
          span.appendChild(reference);
          break;
        }
        case 'footnoteRef':
        case 'endnoteRef':
          if (!props.vertAlign) {
            span.appendChild(
              h('span', {
                text: this.noteNumber,
                style: { 'vertical-align': 'super', 'font-size': '65%' },
              }),
            );
          } else {
            text(this.noteNumber);
          }
          break;
        case 'separator':
        case 'continuationSeparator':
          span.appendChild(
            h('span', {
              style: {
                display: 'inline-block',
                width: node.local === 'separator' ? '2in' : '100%',
                'border-top': '1px solid #000',
                'vertical-align': '0.3em',
              },
            }),
          );
          break;
        case 'drawing':
          if (!this.suppressed()) this.renderDrawing(node, span);
          break;
        case 'pict':
        case 'object':
          if (!this.suppressed()) this.renderVml(node, span);
          break;
        case 'AlternateContent': {
          // Prefer the modern representation; the fallback duplicates the same content.
          if (this.suppressed()) break;
          const choice = child(node, 'Choice');
          const drawing = descendants(choice, 'drawing')[0];
          if (drawing) this.renderDrawing(drawing, span);
          else {
            const fallback = child(node, 'Fallback');
            if (fallback) this.renderVml(fallback, span);
          }
          break;
        }
      }
    }
    if (!span.firstChild || props.hidden) return;
    setStyles(span, this.runCss(props));
    this.paragraph?.sizes.push(props.size ?? this.defaultRun.size ?? 10);
    this.runTarget(parent).appendChild(span);
  }

  private loadImage(relationshipId: string | undefined, image: HTMLImageElement): boolean {
    const relationship = relationshipId ? this.relationships.get(relationshipId) : undefined;
    if (!relationship || relationship.external) return false;
    const mimeType = imageMimeType(relationship.target);
    if (!mimeType) return false;
    this.tasks.push(
      this.zip.read(relationship.target).then((data) => {
        if (data) image.src = this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType }));
      }),
    );
    return true;
  }

  private readonly scheme = (name: string): string | undefined => {
    const aliases: Record<string, string> = { tx1: 'dk1', tx2: 'dk2', bg1: 'lt1', bg2: 'lt2' };
    return this.theme.colors.get(aliases[name] ?? name);
  };

  /** Object URL of an embedded picture, or `null` when it is missing or not displayable. */
  private imageUrl(relationshipId: string | undefined): Promise<string | null> {
    const relationship = relationshipId ? this.relationships.get(relationshipId) : undefined;
    const mimeType = relationship && !relationship.external && imageMimeType(relationship.target);
    if (!relationship || !mimeType) return Promise.resolve(null);
    const pending = this.zip
      .read(relationship.target)
      .then((data) =>
        data ? this.host.objectUrl(new Blob([toBlobPart(data)], { type: mimeType })) : null,
      );
    this.tasks.push(pending.then(() => undefined));
    return pending;
  }

  /** Fill and line of a shape, falling back to the theme style it references. */
  private shapePaint(shape: XmlElement): ShapePaint {
    const properties = child(shape, 'spPr');
    const style = child(shape, 'style');
    let fill: ShapePaint['fill'] | undefined;
    for (const element of childElements(properties)) {
      if (element.local === 'noFill') fill = null;
      else if (element.local === 'solidFill') fill = resolveColor(element, this.scheme) ?? null;
      else if (element.local === 'gradFill') fill = parseGradient(element, this.scheme);
      else if (element.local === 'blipFill') {
        fill = { image: this.imageUrl(attr(child(element, 'blip'), 'r:embed')) };
      } else continue;
      break;
    }
    if (fill === undefined) {
      const reference = child(style, 'fillRef');
      fill =
        reference && attr(reference, 'idx') !== '0'
          ? (resolveColor(reference, this.scheme) ?? null)
          : null;
    }
    const line = child(properties, 'ln');
    let stroke: string | null | undefined;
    if (line) {
      if (child(line, 'noFill')) stroke = null;
      else if (child(line, 'solidFill')) {
        stroke = resolveColor(child(line, 'solidFill'), this.scheme) ?? null;
      }
    }
    if (stroke === undefined) {
      const reference = child(style, 'lnRef');
      stroke =
        reference && attr(reference, 'idx') !== '0'
          ? (resolveColor(reference, this.scheme) ?? null)
          : null;
    }
    const ending = (name: string): boolean => {
      const type = attr(child(line, name), 'type');
      return type !== undefined && type !== 'none';
    };
    return {
      fill,
      stroke,
      strokeWidth: Math.max(0.75, emuToPx(num(attr(line, 'w')) ?? 9525)),
      dash: attr(child(line, 'prstDash'), 'val'),
      headEnd: ending('headEnd'),
      tailEnd: ending('tailEnd'),
    };
  }

  /** One DrawingML object (picture, shape, text box or group) as an element of the given size. */
  private renderShape(shape: XmlElement, width: number, height: number): HTMLElement | null {
    const size = { width: px(width) ?? '0', height: px(height) ?? '0' };
    const grouped = shape.local === 'wgp' || shape.local === 'grpSp';
    const frame = find(shape, grouped ? 'grpSpPr' : 'spPr', 'xfrm');
    const rotation = num(attr(frame, 'rot')) ?? 0;
    const transforms =
      [
        rotation ? `rotate(${rotation / 60000}deg)` : '',
        attr(frame, 'flipH') === '1' ? 'scaleX(-1)' : '',
        attr(frame, 'flipV') === '1' ? 'scaleY(-1)' : '',
      ]
        .filter(Boolean)
        .join(' ') || undefined;

    if (grouped) {
      // A group: children are positioned in the group's own coordinate space.
      const childOffset = child(frame, 'chOff');
      const childExtent = child(frame, 'chExt');
      const scaleX = width / (num(attr(childExtent, 'cx')) || 1);
      const scaleY = height / (num(attr(childExtent, 'cy')) || 1);
      const group = h('span', {
        class: 'av-doc-box',
        style: { display: 'inline-block', position: 'relative', transform: transforms, ...size },
      });
      for (const member of childElements(shape)) {
        if (member.local !== 'pic' && member.local !== 'wsp' && member.local !== 'grpSp') continue;
        const memberFrame = find(member, member.local === 'grpSp' ? 'grpSpPr' : 'spPr', 'xfrm');
        const offset = child(memberFrame, 'off');
        const extent = child(memberFrame, 'ext');
        const element = this.renderShape(
          member,
          (num(attr(extent, 'cx')) ?? 0) * scaleX,
          (num(attr(extent, 'cy')) ?? 0) * scaleY,
        );
        if (!element) continue;
        const left = ((num(attr(offset, 'x')) ?? 0) - (num(attr(childOffset, 'x')) ?? 0)) * scaleX;
        const top = ((num(attr(offset, 'y')) ?? 0) - (num(attr(childOffset, 'y')) ?? 0)) * scaleY;
        setStyles(element, {
          position: 'absolute',
          left: `${left.toFixed(2)}px`,
          top: `${top.toFixed(2)}px`,
        });
        group.appendChild(element);
      }
      return group.firstChild ? group : null;
    }

    if (shape.local !== 'wsp') {
      const blip = shape.local === 'blip' ? shape : descendants(shape, 'blip')[0];
      if (!blip) return null;
      const image = h('img', {
        attrs: { alt: '' },
        style: { ...size, 'max-width': 'none', transform: transforms },
      });
      if (this.loadImage(attr(blip, 'r:embed'), image)) return image;
      // EMF/WMF/TIFF cannot be drawn by browsers: keep the space so the layout is preserved.
      return h('span', {
        class: 'av-doc-box',
        style: { display: 'inline-block', border: '1px dashed #bbb', ...size },
      });
    }

    const properties = child(shape, 'spPr');
    const paint = this.shapePaint(shape);
    const textBox = descendants(shape, 'txbxContent')[0];
    const painted = paint.fill !== null || paint.stroke !== null;
    if (!textBox && !painted) return null;
    const body = child(shape, 'bodyPr');
    const anchor = attr(body, 'anchor');
    const box = h('span', {
      class: 'av-doc-box',
      style: {
        display: 'inline-flex',
        'flex-direction': 'column',
        'justify-content': anchor === 'ctr' ? 'center' : anchor === 'b' ? 'flex-end' : 'flex-start',
        position: 'relative',
        'vertical-align': 'top',
        'text-indent': '0',
        'text-align': 'left',
        'white-space': 'normal',
        width: size.width,
        [child(body, 'spAutoFit') ? 'min-height' : 'height']: size.height,
        transform: rotation ? `rotate(${rotation / 60000}deg)` : undefined,
      },
    });
    if (painted) {
      const svg = buildShapeSvg(shapeGeometry(properties, width, height), width, height, paint);
      // A flipped shape mirrors its outline but keeps its text readable.
      const flips = [
        attr(frame, 'flipH') === '1' ? 'scaleX(-1)' : '',
        attr(frame, 'flipV') === '1' ? 'scaleY(-1)' : '',
      ]
        .filter(Boolean)
        .join(' ');
      if (flips) svg.style.transform = flips;
      box.appendChild(svg);
    }
    if (textBox) {
      const inset = (name: string, fallback: number): string =>
        `${emuToPx(num(attr(body, name)) ?? fallback).toFixed(1)}px`;
      const reference = find(shape, 'style', 'fontRef');
      const content = h('span', {
        style: {
          display: 'block',
          position: 'relative',
          padding: [
            inset('tIns', 45720),
            inset('rIns', 91440),
            inset('bIns', 45720),
            inset('lIns', 91440),
          ].join(' '),
          color: reference ? resolveColor(reference, this.scheme) : undefined,
        },
      });
      this.renderNested(childElements(textBox), content);
      box.appendChild(content);
    }
    return box;
  }

  /** Renders blocks inside an inline container without disturbing the surrounding paragraph. */
  private renderNested(elements: XmlElement[], container: HTMLElement): void {
    const paragraph = this.paragraph;
    const last = this.lastParagraph;
    this.lastParagraph = null;
    this.renderBlocks(elements, container);
    this.paragraph = paragraph;
    this.lastParagraph = last;
  }

  private renderDrawing(drawing: XmlElement, parent: HTMLElement): void {
    const container = child(drawing, 'inline') ?? child(drawing, 'anchor');
    if (!container) return;
    const extent = child(container, 'extent');
    const width = emuToPx(num(attr(extent, 'cx')) ?? 0);
    const height = emuToPx(num(attr(extent, 'cy')) ?? 0);
    const data = find(container, 'graphic', 'graphicData');
    const shape = childElements(data)[0];
    if (!shape) return;
    const element = this.renderShape(shape, width, height);
    if (!element) return;
    if (container.local === 'anchor') {
      parent.appendChild(this.placeAnchored(element, container, width, height));
    } else {
      parent.appendChild(element);
    }
  }

  /** Positions a floating object: wrapped text becomes a CSS float, the rest is placed absolutely. */
  private placeAnchored(
    element: HTMLElement,
    anchor: XmlElement,
    width: number,
    height: number,
  ): HTMLElement {
    const section = this.section();
    const contentWidth = section.width - section.left - section.right;
    const distance = (name: string): number => emuToPx(num(attr(anchor, name)) ?? 0);
    const position = (
      name: string,
    ): { from: string; align: string | undefined; offset: number } => {
      const node = child(anchor, name);
      const offset = num(textOf(child(node, 'posOffset')));
      return {
        from: attr(node, 'relativeFrom') ?? 'column',
        align: child(node, 'align') ? textOf(child(node, 'align')) : undefined,
        offset: emuToPx(offset ?? 0),
      };
    };
    const horizontal = position('positionH');
    const vertical = position('positionV');
    const pageRelative = horizontal.from === 'page' || horizontal.from === 'leftMargin';
    // Horizontal position measured from the left edge of the text area.
    let x = horizontal.offset - (pageRelative ? section.left : 0);
    if (horizontal.align === 'center') {
      x = pageRelative ? (section.width - width) / 2 - section.left : (contentWidth - width) / 2;
    } else if (horizontal.align === 'right' || horizontal.align === 'outside') {
      x = pageRelative ? section.width - width - section.left : contentWidth - width;
    } else if (horizontal.align === 'left' || horizontal.align === 'inside') {
      x = pageRelative ? -section.left : 0;
    }
    const atPageTop =
      (vertical.from === 'margin' || vertical.from === 'page' || vertical.from === 'topMargin') &&
      (vertical.align === 'top' || (vertical.align === undefined && vertical.offset < 24));

    if (child(anchor, 'wrapTopAndBottom')) {
      // A full-width float: the text continues below it and keeps its first-line indent.
      const centred = horizontal.align === 'center';
      element.style.verticalAlign = 'top';
      return h(
        'span',
        {
          style: {
            float: 'left',
            width: '100%',
            'text-indent': '0',
            'line-height': '0',
            'text-align': centred ? 'center' : 'left',
            'padding-left': centred ? undefined : px(Math.max(0, x)),
            'margin-top': px(
              distance('distT') + Math.max(0, vertical.from === 'paragraph' ? vertical.offset : 0),
            ),
            'margin-bottom': px(distance('distB')),
          },
        },
        element,
      );
    }
    if (child(anchor, 'wrapSquare') || child(anchor, 'wrapTight') || child(anchor, 'wrapThrough')) {
      const right = x + width / 2 > contentWidth / 2;
      setStyles(element, {
        float: right ? 'right' : 'left',
        'margin-top': px(
          distance('distT') +
            (atPageTop ? 0 : Math.max(0, vertical.from === 'paragraph' ? vertical.offset : 0)),
        ),
        'margin-bottom': px(distance('distB')),
        'margin-left': right ? px(distance('distL')) : px(Math.max(0, x)),
        'margin-right': right ? px(Math.max(0, contentWidth - x - width)) : px(distance('distR')),
      });
      if (!right && x <= 0) element.style.marginRight = px(distance('distR')) ?? '';
      if (atPageTop && this.cells.length === 0) element.setAttribute('data-av-hoist', '');
      return element;
    }
    // In front of or behind the text: no wrapping.
    const style: Record<string, string | undefined> = {
      position: 'absolute',
      'z-index': attr(anchor, 'behindDoc') === '1' ? '-1' : '1',
    };
    // Offsets are relative to the text area (body) or to the page top (headers and footers).
    if (horizontal.from !== 'character' && this.cells.length === 0) {
      style['left'] = `${x.toFixed(1)}px`;
    }
    if (this.cells.length === 0 && (vertical.from === 'page' || vertical.from === 'margin')) {
      const base = vertical.from === 'page' ? 0 : section.top;
      const areaHeight =
        vertical.from === 'page' ? section.height : section.height - section.top - section.bottom;
      const y =
        vertical.align === 'center'
          ? (areaHeight - height) / 2
          : vertical.align === 'bottom'
            ? areaHeight - height
            : vertical.align === 'top'
              ? 0
              : vertical.offset;
      const origin = this.region === 'body' ? section.top : 0;
      style['top'] = `${(base + y - origin).toFixed(1)}px`;
    } else if (vertical.offset) {
      style['margin-top'] = `${vertical.offset.toFixed(1)}px`;
    }
    setStyles(element, style);
    return element;
  }

  /** Legacy VML pictures and text boxes. */
  private renderVml(graphic: XmlElement, parent: HTMLElement): void {
    const shape = descendants(graphic, 'shape')[0] ?? descendants(graphic, 'rect')[0] ?? null;
    const css = attr(shape, 'style') ?? '';
    const length = (name: string): string | undefined => {
      const match = new RegExp(`(?:^|;)\\s*${name}\\s*:\\s*([\\d.]+)(pt|in|px|cm|mm)`).exec(css);
      return match ? `${match[1]}${match[2]}` : undefined;
    };
    const size = { width: length('width'), height: length('height') };
    // Horizontal rules are stored as a VML rectangle flagged `o:hr`.
    if (descendants(graphic, 'rect').some((rect) => attr(rect, 'o:hr') === 't')) {
      parent.appendChild(
        h('span', {
          class: 'av-doc-box',
          style: { display: 'block', 'border-top': '1px solid #a0a0a0', margin: '0.45em 0' },
        }),
      );
      return;
    }
    for (const data of descendants(graphic, 'imagedata')) {
      const image = h('img', { attrs: { alt: '' }, style: size });
      if (this.loadImage(attr(data, 'r:id'), image)) parent.appendChild(image);
      else {
        parent.appendChild(
          h('span', {
            class: 'av-doc-box',
            style: { display: 'inline-block', border: '1px dashed #bbb', ...size },
          }),
        );
      }
    }
    const textBox = descendants(graphic, 'txbxContent')[0];
    if (textBox) {
      const box = h('span', {
        class: 'av-doc-box',
        style: {
          display: 'inline-block',
          'vertical-align': 'top',
          'text-indent': '0',
          'white-space': 'normal',
          width: size.width,
        },
      });
      this.renderNested(childElements(textBox), box);
      parent.appendChild(box);
    }
  }

  private tableLook(tblPr: XmlElement | null): TableLook {
    const look = child(tblPr, 'tblLook');
    const bits = parseInt(attr(look, 'w:val') ?? '', 16);
    const flag = (name: string, bit: number, fallback: boolean): boolean => {
      const value = attr(look, `w:${name}`);
      if (value !== undefined) return value === '1' || value === 'true';
      return Number.isFinite(bits) ? (bits & bit) !== 0 : fallback;
    };
    return {
      firstRow: flag('firstRow', 0x20, true),
      lastRow: flag('lastRow', 0x40, false),
      firstCol: flag('firstColumn', 0x80, true),
      lastCol: flag('lastColumn', 0x100, false),
      noHBand: flag('noHBand', 0x200, false),
      noVBand: flag('noVBand', 0x400, true),
    };
  }

  private renderTable(table: XmlElement): HTMLElement {
    const tblPr = child(table, 'tblPr');
    const chain = this.styleChain(
      attr(child(tblPr, 'tblStyle'), 'w:val') ?? this.defaultTableStyle,
    );
    const direct = this.parser.table(tblPr);
    const tableProps = merge<TableProps>(...chain.map((style) => style.table), direct);
    const borders = merge<BorderSet>(...chain.map((style) => style.table.borders), direct.borders);
    const margins = merge<Partial<Record<Side, number>>>(
      { top: 0, left: 7.2, bottom: 0, right: 7.2 },
      ...chain.map((style) => style.table.margins),
      direct.margins,
    );
    const styleCell = merge<CellProps>(...chain.map((style) => style.cell));
    const stylePara = merge<ParaProps>(...chain.map((style) => style.para));
    const styleRun = merge<RunProps>(...chain.map((style) => style.run));
    const conditional = (type: string): ConditionalFormat | undefined => {
      const formats = chain.flatMap((style) => {
        const format = style.conditionals.get(type);
        return format ? [format] : [];
      });
      if (formats.length === 0) return undefined;
      return {
        para: merge<ParaProps>(...formats.map((format) => format.para)),
        run: merge<RunProps>(...formats.map((format) => format.run)),
        cell: merge<CellProps>(...formats.map((format) => format.cell)),
      };
    };
    const look = this.tableLook(tblPr);
    const rowBand = tableProps.rowBandSize ?? 1;
    const colBand = tableProps.colBandSize ?? 1;

    // Width and position.
    const grid = childElements(child(table, 'tblGrid'), 'gridCol').map((column) =>
      twipsToPx(num(attr(column, 'w:w')) ?? 0),
    );
    const gridWidth = grid.reduce((sum, width) => sum + width, 0);
    const widthElement = child(tblPr, 'tblW');
    const widthType = attr(widthElement, 'w:type');
    const widthValue = num(attr(widthElement, 'w:w')) ?? 0;
    const shift = this.legacyTables ? (margins.left ?? 0) : 0;
    let width: string | undefined;
    if (widthType === 'pct' && widthValue > 0) {
      // Word measures a percentage against the text width plus the cell margins the table
      // hangs into; a table inside a cell is measured against the cell's own text width.
      const extra =
        this.legacyTables && this.tableDepth === 0 ? (margins.left ?? 0) + (margins.right ?? 0) : 0;
      width = `calc((100% + ${extra}px) * ${widthValue / 5000})`;
    } else if (widthType === 'dxa' && widthValue > 0) {
      width = px(twipsToPx(widthValue));
    } else if (gridWidth > 0) {
      width = px(gridWidth);
    }
    const floating = child(tblPr, 'tblpPr');
    const floatSide = floating
      ? attr(floating, 'w:tblpXSpec') === 'right'
        ? 'right'
        : attr(floating, 'w:tblpXSpec') === 'center'
          ? undefined
          : 'left'
      : undefined;
    const centred = tableProps.align === 'center' || attr(floating, 'w:tblpXSpec') === 'center';
    const rightAligned = tableProps.align === 'right' || tableProps.align === 'end';
    const fromText = (name: string): number => twipsToPx(num(attr(floating, `w:${name}`)) ?? 0);
    const element = h('table', {
      class: 'av-doc-table',
      style: {
        width,
        'table-layout': width && gridWidth > 0 ? 'fixed' : undefined,
        float: floatSide,
        'margin-left':
          centred || rightAligned
            ? 'auto'
            : `${(twipsToPx(tableProps.indent ?? 0) - shift).toFixed(2)}px`,
        'margin-right': centred
          ? 'auto'
          : floatSide === 'left'
            ? px(fromText('rightFromText'))
            : undefined,
        'margin-top': floatSide ? px(fromText('topFromText')) : undefined,
        'margin-bottom': floatSide ? px(fromText('bottomFromText')) : undefined,
      },
    });
    if (floatSide === 'right') element.style.marginLeft = px(fromText('leftFromText')) ?? '0';
    if (gridWidth > 0) {
      const group = h('colgroup');
      for (const column of grid) {
        group.appendChild(
          h('col', { style: { width: `${((column / gridWidth) * 100).toFixed(3)}%` } }),
        );
      }
      element.appendChild(group);
    }

    // First pass: build the cells and the grid they occupy.
    const rows = childElements(table, 'tr');
    const columnCount = Math.max(
      grid.length,
      ...rows.map((row) => {
        const before = num(attr(find(row, 'trPr', 'gridBefore'), 'w:val')) ?? 0;
        return childElements(row, 'tc').reduce(
          (sum, cell) => sum + Math.max(1, num(attr(find(cell, 'tcPr', 'gridSpan'), 'w:val')) ?? 1),
          before,
        );
      }),
    );
    const occupancy: (CellBox | undefined)[][] = rows.map(() => []);
    const boxes: CellBox[] = [];
    const openMerges = new Map<number, CellBox>();
    const tbody = h('tbody');
    const rendering: { box: CellBox; cell: XmlElement }[] = [];

    rows.forEach((row, rowIndex) => {
      const trPr = child(row, 'trPr');
      const tr = h('tr');
      const heightElement = child(trPr, 'trHeight');
      const rowHeight = num(attr(heightElement, 'w:val'));
      if (rowHeight) tr.style.height = px(twipsToPx(rowHeight)) ?? '';
      if (child(trPr, 'tblHeader') && isOn(child(trPr, 'tblHeader')))
        tr.setAttribute('data-av-th', '');
      let column = 0;
      const before = num(attr(child(trPr, 'gridBefore'), 'w:val')) ?? 0;
      if (before > 0) {
        tr.appendChild(
          h('td', { attrs: { colspan: before }, style: { border: 'none', padding: '0' } }),
        );
        column = before;
      }
      for (const cell of childElements(row, 'tc')) {
        const tcPr = child(cell, 'tcPr');
        const span = Math.max(1, num(attr(child(tcPr, 'gridSpan'), 'w:val')) ?? 1);
        const vMerge = child(tcPr, 'vMerge');
        if (vMerge && attr(vMerge, 'w:val') !== 'restart') {
          const origin = openMerges.get(column);
          if (origin) {
            origin.rowSpan += 1;
            origin.td.rowSpan = origin.rowSpan;
            for (let c = column; c < column + span; c++) occupancy[rowIndex]![c] = origin;
            column += span;
            continue;
          }
        }
        const td = h('td', { attrs: { colspan: span > 1 ? span : undefined } });
        const box: CellBox = {
          td,
          row: rowIndex,
          col: column,
          colSpan: span,
          rowSpan: 1,
          conditionals: [],
          direct: this.parser.cell(tcPr),
          borders: {
            top: { border: undefined, rank: 0 },
            left: { border: undefined, rank: 0 },
            bottom: { border: undefined, rank: 0 },
            right: { border: undefined, rank: 0 },
          },
        };
        if (vMerge) openMerges.set(column, box);
        else openMerges.delete(column);
        for (let c = column; c < column + span; c++) occupancy[rowIndex]![c] = box;
        boxes.push(box);
        rendering.push({ box, cell });
        tr.appendChild(td);
        column += span;
      }
      const after = num(attr(child(trPr, 'gridAfter'), 'w:val')) ?? 0;
      if (after > 0) {
        tr.appendChild(
          h('td', { attrs: { colspan: after }, style: { border: 'none', padding: '0' } }),
        );
      }
      tbody.appendChild(tr);
    });

    // Second pass: conditional formatting and borders, now that every span is known.
    const lastRow = rows.length - 1;
    for (const box of boxes) {
      const { row, col } = box;
      const endCol = col + box.colSpan - 1;
      const endRow = row + box.rowSpan - 1;
      const inFirstRow = look.firstRow && row === 0;
      const inLastRow = look.lastRow && endRow === lastRow && lastRow > 0;
      const inFirstCol = look.firstCol && col === 0;
      const inLastCol = look.lastCol && endCol === columnCount - 1 && columnCount > 1;
      const types: string[] = [];
      if (!look.noVBand && !inFirstCol && !inLastCol) {
        const index = col - (look.firstCol ? 1 : 0);
        types.push(Math.floor(index / colBand) % 2 === 0 ? 'band1Vert' : 'band2Vert');
      }
      if (!look.noHBand && !inFirstRow && !inLastRow) {
        const index = row - (look.firstRow ? 1 : 0);
        types.push(Math.floor(index / rowBand) % 2 === 0 ? 'band1Horz' : 'band2Horz');
      }
      if (inFirstCol) types.push('firstCol');
      if (inLastCol) types.push('lastCol');
      if (inFirstRow) types.push('firstRow');
      if (inLastRow) types.push('lastRow');
      if (inFirstRow && inLastCol) types.push('neCell');
      if (inFirstRow && inFirstCol) types.push('nwCell');
      if (inLastRow && inLastCol) types.push('seCell');
      if (inLastRow && inFirstCol) types.push('swCell');
      for (const type of types) {
        const format = conditional(type);
        if (format) box.conditionals.push({ type, format });
      }

      const atTop = row === 0;
      const atBottom = endRow === lastRow;
      const atLeft = col === 0;
      const atRight = endCol === columnCount - 1;
      box.borders.top = { border: atTop ? borders.top : borders.insideH, rank: 0 };
      box.borders.bottom = { border: atBottom ? borders.bottom : borders.insideH, rank: 0 };
      box.borders.left = { border: atLeft ? borders.left : borders.insideV, rank: 0 };
      box.borders.right = { border: atRight ? borders.right : borders.insideV, rank: 0 };
      const set = (side: Side, border: Border | undefined, rank: number): void => {
        if (border) box.borders[side] = { border, rank };
      };
      for (const { type, format } of box.conditionals) {
        const own = format.cell.borders;
        if (!own) continue;
        // A band that switches a border off does not erase the outer border of the table.
        const band = /^band/.test(type);
        const inner = (border: Border | undefined, edge: boolean): Border | undefined =>
          band && edge && border?.style === 'none' ? undefined : border;
        if (/Row$|Horz$/.test(type)) {
          set('top', inner(own.top, atTop), 1);
          set('bottom', inner(own.bottom, atBottom), 1);
          set('left', atLeft ? own.left : own.insideV, 1);
          set('right', atRight ? own.right : own.insideV, 1);
        } else if (/Col$|Vert$/.test(type)) {
          set('left', inner(own.left, atLeft), 1);
          set('right', inner(own.right, atRight), 1);
          set('top', atTop ? own.top : own.insideH, 1);
          set('bottom', atBottom ? own.bottom : own.insideH, 1);
        } else {
          for (const side of SIDES) set(side, own[side], 1);
        }
      }
      for (const side of SIDES) set(side, box.direct.borders?.[side], 2);
    }
    // Shared edges: both neighbours draw the border that wins.
    for (const box of boxes) {
      const below = occupancy[box.row + box.rowSpan]?.[box.col];
      if (below && below !== box) {
        const winner = strongerBorder(box.borders.bottom, below.borders.top);
        box.borders.bottom = winner;
        below.borders.top = winner;
      }
      const beside = occupancy[box.row]?.[box.col + box.colSpan];
      if (beside && beside !== box) {
        const winner = strongerBorder(box.borders.right, beside.borders.left);
        box.borders.right = winner;
        beside.borders.left = winner;
      }
    }

    // Third pass: styles and content.
    for (const { box, cell } of rendering) {
      const formats = box.conditionals.map((entry) => entry.format);
      const props = merge<CellProps>(
        styleCell,
        ...formats.map((format) => format.cell),
        box.direct,
      );
      const padding = merge<Partial<Record<Side, number>>>(
        margins,
        ...formats.map((format) => format.cell.margins),
        box.direct.margins,
      );
      setStyles(box.td, {
        'background-color': props.background,
        'vertical-align':
          props.vAlign === 'center' ? 'middle' : props.vAlign === 'bottom' ? 'bottom' : undefined,
        padding: (['top', 'right', 'bottom', 'left'] as const)
          .map((side) => `${padding[side] ?? 0}px`)
          .join(' '),
        'white-space': props.noWrap ? 'nowrap' : undefined,
        'writing-mode':
          props.textDirection && /^(tbRl|tbRlV|btLr)$/.test(props.textDirection)
            ? 'vertical-rl'
            : undefined,
        transform: props.textDirection === 'btLr' ? 'rotate(180deg)' : undefined,
      });
      for (const side of SIDES) {
        box.td.style.setProperty(`border-${side}`, borderCss(box.borders[side].border) ?? 'none');
      }
      this.cells.push({
        para: merge<ParaProps>(stylePara, ...formats.map((format) => format.para)),
        run: merge<RunProps>(styleRun, ...formats.map((format) => format.run)),
      });
      const outerParagraph = this.paragraph;
      this.lastParagraph = null;
      this.tableDepth++;
      this.renderBlocks(childElements(cell), box.td);
      this.tableDepth--;
      // The empty paragraph Word requires after a nested table takes no space.
      const closing = box.td.lastElementChild as HTMLElement | null;
      if (
        closing?.tagName === 'P' &&
        closing.previousElementSibling?.tagName === 'TABLE' &&
        !closing.textContent &&
        !closing.querySelector('img')
      ) {
        closing.style.display = 'none';
      }
      this.paragraph = outerParagraph;
      this.cells.pop();
    }
    element.appendChild(tbody);
    return element;
  }
}

/** Renders a DOCX package into a detached element; call `layout()` once it is attached. */
export function renderDocx(zip: ZipArchive, host: DocxHost): Promise<DocxDocument> {
  return new DocxRenderer(zip, host).render();
}
