import { h } from '../ui/dom';

/**
 * Layout for rendered Word documents: tab stops and pagination. Both need real
 * measurements, so they run after the document has been attached to the page.
 * Without a layout engine (tests, server rendering) every measurement is zero
 * and the document simply stays on the pages created by explicit breaks.
 */

export type HeaderSet = Partial<Record<'default' | 'first' | 'even', HTMLElement>>;

/** Page geometry of one section, in CSS pixels. */
export interface PageSetup {
  width: number;
  height: number;
  top: number;
  right: number;
  bottom: number;
  left: number;
  header: number;
  footer: number;
  /** A continuous section starts on the same page as the previous one. */
  continuous: boolean;
  titlePage: boolean;
  headers: HeaderSet;
  footers: HeaderSet;
}

export interface PageParts {
  page: HTMLElement;
  body: HTMLElement;
  notes: HTMLElement;
  setup: PageSetup;
}

const px = (value: number): string => `${Math.round(value * 100) / 100}px`;

function pick(set: HeaderSet, first: boolean, even: boolean, setup: PageSetup): HTMLElement | null {
  if (setup.titlePage && first) return set.first ?? null;
  if (even) return set.even ?? set.default ?? null;
  return set.default ?? null;
}

/** Builds an empty page. `paged` fixes the height; otherwise the page grows with its content. */
export function createPage(
  setup: PageSetup,
  pageNumber: number,
  options: {
    paged: boolean;
    firstOfSection: boolean;
    evenAndOdd: boolean;
    /** Line drawn above the footnotes of a page. */
    separator?: HTMLElement | undefined;
  },
): PageParts {
  const body = h('div', { class: 'av-doc-body' });
  const notes = h('div', { class: 'av-doc-footnotes', attrs: { hidden: true } });
  if (options.separator) notes.appendChild(options.separator.cloneNode(true));
  const page = h('div', {
    class: 'av-paper av-doc-page',
    attrs: { 'data-av-page': pageNumber },
    style: {
      width: px(setup.width),
      [options.paged ? 'height' : 'min-height']: px(setup.height),
      padding: [setup.top, setup.right, setup.bottom, setup.left].map(px).join(' '),
    },
  });
  const even = options.evenAndOdd && pageNumber % 2 === 0;
  const header = pick(setup.headers, options.firstOfSection, even, setup);
  const footer = pick(setup.footers, options.firstOfSection, even, setup);
  if (header) {
    const copy = header.cloneNode(true) as HTMLElement;
    copy.className = 'av-doc-header';
    // The header box starts at the top of the page so that floating objects can be placed from there.
    Object.assign(copy.style, {
      left: px(setup.left),
      right: px(setup.right),
      top: '0',
      paddingTop: px(setup.header),
    });
    page.appendChild(copy);
  }
  page.append(body, notes);
  if (footer) {
    const copy = h('div', { class: 'av-doc-footer' }, footer.cloneNode(true));
    Object.assign(copy.style, {
      left: px(setup.left),
      right: px(setup.right),
      top: '0',
      bottom: px(setup.footer),
    });
    page.appendChild(copy);
  }
  return { page, body, notes, setup };
}

/** Writes page numbers into `PAGE` and `NUMPAGES` fields. */
export function fillFields(root: HTMLElement, pageNumber: number, total: number): void {
  for (const field of Array.from(root.querySelectorAll<HTMLElement>('[data-av-field]'))) {
    field.textContent = String(
      field.getAttribute('data-av-field') === 'pages' ? total : pageNumber,
    );
  }
}

/** Ratio between layout pixels and on-screen pixels (CSS `zoom` on an ancestor). */
function measureScale(element: HTMLElement): number {
  const width = element.offsetWidth;
  const scale = width > 0 ? element.getBoundingClientRect().width / width : 1;
  return scale > 0 && Number.isFinite(scale) ? scale : 1;
}

interface Stop {
  pos: number;
  align: string;
  leader: string;
}

function parseStops(value: string | null): Stop[] {
  if (!value) return [];
  return value.split(';').flatMap((entry): Stop[] => {
    const [pos, align, leader] = entry.split(':');
    const position = Number(pos);
    return Number.isFinite(position) && pos !== ''
      ? [{ pos: position, align: align ?? 'left', leader: leader ?? 'none' }]
      : [];
  });
}

const LEADERS: Readonly<Record<string, string>> = {
  dot: 'dot',
  middleDot: 'dot',
  hyphen: 'hyphen',
  underscore: 'line',
  heavy: 'line',
};

/**
 * Gives every tab the width that takes the text to its next tab stop, like a
 * word processor does. Paragraphs carry their stops in `data-av-tabs`.
 */
export function layoutTabs(root: HTMLElement): void {
  const paragraphs = Array.from(root.querySelectorAll<HTMLElement>('[data-av-tabs]'));
  if (root.hasAttribute('data-av-tabs')) paragraphs.unshift(root);
  if (paragraphs.length === 0) return;
  const scale = measureScale(root);
  const range = document.createRange();

  for (const paragraph of paragraphs) {
    const tabs = Array.from(paragraph.querySelectorAll<HTMLElement>('.av-doc-tab')).filter(
      (tab) => tab.closest('[data-av-tabs]') === paragraph,
    );
    const marker = paragraph.querySelector<HTMLElement>(':scope > .av-doc-marker[data-av-suff]');
    const stops = parseStops(paragraph.getAttribute('data-av-tabs'));
    if (tabs.length === 0 && !marker) {
      alignDecimal(paragraph, stops[0], scale, range);
      continue;
    }
    const [indentLeft = 0, firstLine = 0, defaultStop = 48] = (
      paragraph.getAttribute('data-av-ind') ?? ''
    )
      .split(',')
      .map(Number);
    const step = defaultStop > 0 ? defaultStop : 48;
    for (const tab of tabs) tab.style.width = '0px';
    const origin = paragraph.getBoundingClientRect().left / scale - indentLeft;

    // A list number wider than its hanging indent pushes the text to the next tab stop.
    const label = marker?.firstElementChild;
    if (marker && label) {
      const start = indentLeft + firstLine;
      const end = start + label.getBoundingClientRect().width / scale;
      if (end > indentLeft + 0.5) {
        const stop =
          stops.find((candidate) => candidate.pos > end + 0.5)?.pos ??
          (Math.floor((end + 0.5) / step) + 1) * step;
        marker.style.minWidth = px(stop - start);
        marker.style.paddingRight = '0';
      }
    }

    tabs.forEach((tab, index) => {
      const x = tab.getBoundingClientRect().left / scale - origin;
      let stop: Stop | undefined;
      const absolute = tab.getAttribute('data-av-ptab');
      if (absolute) {
        const frame = tab.closest<HTMLElement>('.av-doc-header,.av-doc-footer,.av-doc-body,td');
        const width = frame ? frame.clientWidth : 0;
        const pos = (absolute === 'center' ? width / 2 : absolute === 'right' ? width : 0) + 0;
        stop = { pos, align: absolute, leader: tab.getAttribute('data-av-leader') ?? 'none' };
      } else {
        stop = stops.find((candidate) => candidate.pos > x + 0.5);
        if (!stop) {
          // A hanging indent acts as a tab stop; after that come the default stops.
          const pos =
            firstLine < 0 && indentLeft > x + 0.5
              ? indentLeft
              : (Math.floor((x + 0.5) / step) + 1) * step;
          stop = { pos, align: 'left', leader: 'none' };
        }
      }
      let width = stop.pos - x;
      if (stop.align === 'right' || stop.align === 'center' || stop.align === 'decimal') {
        // The text after the tab, up to the next tab, is aligned against the stop.
        range.setStartAfter(tab);
        const next = tabs[index + 1];
        if (next) range.setEndBefore(next);
        else range.setEnd(paragraph, paragraph.childNodes.length);
        const following = range.getBoundingClientRect().width / scale;
        width -= stop.align === 'center' ? following / 2 : following;
      }
      tab.style.width = px(Math.max(0, width));
      const leader = LEADERS[stop.leader];
      if (leader) tab.setAttribute('data-av-leader-style', leader);
    });
  }
}

/**
 * In a table cell, a paragraph whose first tab stop is a decimal stop is aligned on that stop
 * without a tab character: the decimal point (or the end of the text) sits at the stop.
 */
function alignDecimal(
  paragraph: HTMLElement,
  stop: Stop | undefined,
  scale: number,
  range: Range,
): void {
  if (stop?.align !== 'decimal' || paragraph.parentElement?.tagName !== 'TD') return;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  let first: Node | null = null;
  let end: Node | null = null;
  let offset = 0;
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const text = node.nodeValue ?? '';
    if (text === '') continue;
    first ??= node;
    end = node;
    const point = text.indexOf('.');
    offset = point === -1 ? text.length : point;
    if (point !== -1) break;
  }
  if (!first || !end) return;
  const [indentLeft = 0, firstLine = 0] = (paragraph.getAttribute('data-av-ind') ?? '')
    .split(',')
    .map(Number);
  paragraph.style.textIndent = px(firstLine);
  range.setStart(first, 0);
  range.setEnd(end, offset);
  const before = range.getBoundingClientRect().width / scale;
  const shift = stop.pos - indentLeft - firstLine - before;
  if (shift > 0.5) paragraph.style.textIndent = px(firstLine + shift);
}

/** Moves the rows that do not fit above `limit` into a new table. */
function splitTable(table: HTMLTableElement, limit: number): HTMLTableElement | null {
  const section = table.querySelector(':scope > tbody');
  if (!section) return null;
  const rows = Array.from(section.children) as HTMLElement[];
  let headers = 0;
  while (headers < rows.length && rows[headers]!.hasAttribute('data-av-th')) headers++;
  const index = rows.findIndex((row) => row.getBoundingClientRect().bottom > limit + 0.5);
  if (index <= headers || index === -1) return null;
  const rest = table.cloneNode(false) as HTMLTableElement;
  const columns = table.querySelector(':scope > colgroup');
  if (columns) rest.appendChild(columns.cloneNode(true));
  const body = h('tbody');
  for (let i = 0; i < headers; i++) body.appendChild(rows[i]!.cloneNode(true));
  for (let i = index; i < rows.length; i++) body.appendChild(rows[i]!);
  rest.appendChild(body);
  rest.style.marginTop = '0';
  rest.removeAttribute('data-av-break');
  rest.removeAttribute('data-av-sect');
  return rest;
}

interface Line {
  top: number;
  bottom: number;
}

/** Splits a paragraph before the first line that crosses `limit`; returns the remainder. */
function splitParagraph(paragraph: HTMLElement, limit: number): HTMLElement | null {
  if (!/^(P|H[1-6])$/.test(paragraph.tagName) || paragraph.hasAttribute('data-av-kl')) return null;
  const walker = document.createTreeWalker(paragraph, NodeFilter.SHOW_TEXT);
  const range = document.createRange();
  const fragments: { node: Text; top: number; bottom: number }[] = [];
  for (let node = walker.nextNode(); node; node = walker.nextNode()) {
    const owner = node.parentElement?.closest('p,h1,h2,h3,h4,h5,h6,table');
    if (owner !== paragraph) continue;
    range.selectNodeContents(node);
    for (const rect of Array.from(range.getClientRects())) {
      if (rect.height > 0)
        fragments.push({ node: node as Text, top: rect.top, bottom: rect.bottom });
    }
  }
  if (fragments.length === 0) return null;

  // Group fragments into lines: a fragment that starts below the middle of the current line opens a new one.
  const lines: Line[] = [];
  for (const fragment of [...fragments].sort((a, b) => a.top - b.top)) {
    const line = lines[lines.length - 1];
    if (line && fragment.top < (line.top + line.bottom) / 2) {
      line.bottom = Math.max(line.bottom, fragment.bottom);
    } else {
      lines.push({ top: fragment.top, bottom: fragment.bottom });
    }
  }
  let index = lines.findIndex((line) => line.bottom > limit + 0.5);
  if (index === -1) return null;
  // Widow and orphan control: keep at least two lines on each side of the break.
  const minimum = paragraph.hasAttribute('data-av-nw') ? 1 : 2;
  if (lines.length - index < minimum) index = lines.length - minimum;
  if (index < minimum) return null;
  const target = lines[index]!;

  // The first text node (in document order) with a fragment on or after the target line.
  const start = fragments.find((fragment) => fragment.top >= target.top - 0.5);
  if (!start) return null;
  const { node } = start;
  const onTargetLine = (offset: number): boolean => {
    range.setStart(node, offset);
    range.setEnd(node, Math.min(offset + 1, node.length));
    return range.getBoundingClientRect().top >= start.top - 0.5;
  };
  let low = 0;
  let high = node.length;
  while (low < high) {
    const middle = (low + high) >> 1;
    if (onTargetLine(middle)) high = middle;
    else low = middle + 1;
  }
  range.setStart(node, low);
  range.setEnd(paragraph, paragraph.childNodes.length);
  const rest = paragraph.cloneNode(false) as HTMLElement;
  rest.appendChild(range.extractContents());
  if (!rest.textContent) return null;
  rest.removeAttribute('data-av-break');
  rest.removeAttribute('data-av-sect');
  rest.style.textIndent = '0';
  rest.style.marginTop = '0';
  rest.style.borderTop = '0';
  paragraph.style.marginBottom = '0';
  paragraph.style.borderBottom = '0';
  if (getComputedStyle(paragraph).textAlign === 'justify')
    paragraph.style.textAlignLast = 'justify';
  return rest;
}

/** Splits paragraphs at explicit page breaks so that each part can start a page. */
function splitAtBreaks(blocks: HTMLElement[]): HTMLElement[] {
  const result: HTMLElement[] = [];
  let breakNext = false;
  for (const block of blocks) {
    if (breakNext) {
      block.setAttribute('data-av-break', 'before');
      breakNext = false;
    }
    let current = block;
    result.push(current);
    if (current.tagName === 'TABLE') continue;
    for (;;) {
      const marker = current.querySelector<HTMLElement>('[data-av-pb]');
      if (!marker || marker.closest('table')) break;
      const range = document.createRange();
      range.setStartAfter(marker);
      range.setEnd(current, current.childNodes.length);
      const rest = current.cloneNode(false) as HTMLElement;
      rest.appendChild(range.extractContents());
      marker.remove();
      if (!rest.textContent && !rest.querySelector('img,[data-av-pb]')) {
        breakNext = true;
        break;
      }
      rest.removeAttribute('data-av-sect');
      rest.setAttribute('data-av-break', 'before');
      rest.style.textIndent = '0';
      result.push(rest);
      current = rest;
    }
  }
  return result;
}

export class Paginator {
  readonly pages: PageParts[] = [];
  private current!: PageParts;
  private automatic = false;
  private firstOfSection = true;
  private readonly kept = new WeakSet<Element>();

  constructor(
    private readonly wrapper: HTMLElement,
    private readonly sections: readonly PageSetup[],
    private readonly notes: ReadonlyMap<string, HTMLElement>,
    private readonly evenAndOdd: boolean,
    private readonly separator?: HTMLElement,
  ) {}

  private open(setup: PageSetup, automatic: boolean): void {
    const parts = createPage(setup, this.pages.length + 1, {
      paged: true,
      firstOfSection: this.firstOfSection,
      evenAndOdd: this.evenAndOdd,
      separator: this.separator,
    });
    this.firstOfSection = false;
    this.wrapper.appendChild(parts.page);
    fillFields(parts.page, this.pages.length + 1, this.pages.length + 1);
    // A header or footer taller than its margin pushes the body, as in Word.
    const header = parts.page.querySelector<HTMLElement>(':scope > .av-doc-header');
    const footer = parts.page.querySelector<HTMLElement>(':scope > .av-doc-footer');
    if (header) {
      layoutTabs(header);
      const bottom = header.offsetHeight;
      if (bottom > setup.top) parts.page.style.paddingTop = px(bottom);
    }
    if (footer) {
      layoutTabs(footer);
      const inner = footer.firstElementChild as HTMLElement | null;
      const top = setup.footer + (inner?.offsetHeight ?? 0);
      if (top > setup.bottom) parts.page.style.paddingBottom = px(top);
    }
    this.pages.push(parts);
    this.current = parts;
    this.automatic = automatic;
  }

  /** Shows on the page exactly the footnotes that its text refers to. */
  private syncNotes(parts: PageParts): void {
    if (this.notes.size === 0) return;
    const wanted: HTMLElement[] = [];
    for (const reference of Array.from(parts.body.querySelectorAll('[data-av-fn]'))) {
      const note = this.notes.get(reference.getAttribute('data-av-fn') ?? '');
      if (note && !wanted.includes(note)) wanted.push(note);
    }
    for (const child of Array.from(parts.notes.children)) {
      if (!child.hasAttribute('data-av-sep') && !wanted.includes(child as HTMLElement))
        child.remove();
    }
    for (const note of wanted) parts.notes.appendChild(note);
    parts.notes.hidden = wanted.length === 0;
  }

  private limit(): number {
    const { body } = this.current;
    return body.getBoundingClientRect().top + body.clientHeight * measureScale(body);
  }

  private fits(): boolean {
    const last = this.current.body.lastElementChild;
    return !last || last.getBoundingClientRect().bottom <= this.limit() + 1;
  }

  private place(block: HTMLElement): void {
    const { body } = this.current;
    if (this.automatic && !body.firstChild) block.style.marginTop = '0';
    body.appendChild(block);
    // Pictures aligned with the top of the page start the page, and the text wraps around them.
    for (const float of Array.from(block.querySelectorAll('[data-av-hoist]'))) {
      float.removeAttribute('data-av-hoist');
      body.insertBefore(float, body.firstChild);
    }
    this.syncNotes(this.current);
  }

  run(source: HTMLElement[]): void {
    const blocks = splitAtBreaks(source);
    this.wrapper.textContent = '';
    let setup = this.sections[0]!;
    this.open(setup, false);
    const pending: HTMLElement[] = [];
    let next = 0;
    while (pending.length > 0 || next < blocks.length) {
      const block = pending.pop() ?? blocks[next++]!;
      const section = block.getAttribute('data-av-sect');
      if (section !== null) {
        block.removeAttribute('data-av-sect');
        const following = this.sections[Number(section)];
        if (following) {
          setup = following;
          this.firstOfSection = true;
          if (!following.continuous) {
            if (!this.current.body.firstChild) this.discard();
            this.open(setup, false);
          }
        }
      }
      if (block.getAttribute('data-av-break') === 'before') {
        block.removeAttribute('data-av-break');
        if (this.current.body.firstChild) this.open(setup, false);
      }
      this.place(block);
      if (this.fits()) continue;

      const limit = this.limit();
      const rest =
        block.tagName === 'TABLE'
          ? splitTable(block as HTMLTableElement, limit)
          : splitParagraph(block, limit);
      if (rest) {
        this.syncNotes(this.current);
        this.open(setup, true);
        pending.push(rest);
        continue;
      }
      const { body } = this.current;
      if (block === body.firstElementChild) continue;
      // Move the whole block, together with headings that must stay with it.
      const moved = [block];
      let previous = block.previousElementSibling;
      while (
        previous &&
        previous.hasAttribute('data-av-kn') &&
        previous !== body.firstElementChild &&
        !this.kept.has(previous) &&
        moved.length < 4
      ) {
        this.kept.add(previous);
        moved.push(previous as HTMLElement);
        previous = previous.previousElementSibling;
      }
      for (const element of moved) element.remove();
      this.syncNotes(this.current);
      this.open(setup, true);
      pending.push(...moved);
    }
    if (this.pages.length > 1 && !this.current.body.firstChild) this.discard();
    this.pages.forEach((parts, index) => fillFields(parts.page, index + 1, this.pages.length));
  }

  private discard(): void {
    const parts = this.pages.pop();
    parts?.page.remove();
    const last = this.pages[this.pages.length - 1];
    if (last) this.current = last;
  }
}
