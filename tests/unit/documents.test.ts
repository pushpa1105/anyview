import { afterEach, describe, expect, it } from 'vitest';
import { formatCounter, renderDocx } from '../../src/formats/docx';
import { renderOdt } from '../../src/formats/odt';
import { renderRtf } from '../../src/formats/rtf';
import { renderer as docxRenderer } from '../../src/renderers/docx';
import { renderer as odtRenderer } from '../../src/renderers/odt';
import { renderer as rtfRenderer } from '../../src/renderers/rtf';
import { openZip } from '../../src/util/zip';
import { instanceOf, renderContext } from '../helpers/dom';
import { docxFixture, odtFixture, RTF_FIXTURE } from '../helpers/fixtures';
import { makeZip, PNG_1X1 } from '../helpers/zip';

afterEach(() => {
  document.body.textContent = '';
});

function host() {
  const blobs: Blob[] = [];
  return {
    blobs,
    objectUrl(blob: Blob): string {
      blobs.push(blob);
      return `blob:test/${blobs.length}`;
    },
  };
}

const paragraphWith = (root: Element, text: string): HTMLElement => {
  const found = Array.from(root.querySelectorAll<HTMLElement>('p, h1, h2, h3')).find((element) =>
    element.textContent?.includes(text),
  );
  if (!found) throw new Error(`no paragraph contains "${text}"`);
  return found;
};

const spanWith = (root: Element, text: string): HTMLElement => {
  const found = Array.from(root.querySelectorAll<HTMLElement>('span')).find(
    (element) => element.textContent === text && element.children.length === 0,
  );
  if (!found) throw new Error(`no span has the text "${text}"`);
  return found;
};

describe('formatCounter', () => {
  it.each([
    [1, 'decimal', '1'],
    [3, 'lowerLetter', 'c'],
    [27, 'lowerLetter', 'aa'],
    [2, 'upperLetter', 'B'],
    [4, 'lowerRoman', 'iv'],
    [1994, 'upperRoman', 'MCMXCIV'],
    [4000, 'upperRoman', '4000'],
    [7, 'decimalZero', '07'],
    [7, 'none', ''],
    [0, 'lowerLetter', '0'],
    [5, 'somethingElse', '5'],
  ])('formats %d as %s → %s', (value, format, expected) => {
    expect(formatCounter(value, format)).toBe(expected);
  });
});

describe('DOCX', () => {
  it('applies page size, margins and document defaults', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const page = paper.querySelector<HTMLElement>('.av-doc-page')!;
    expect(parseFloat(page.style.width)).toBeCloseTo(11906 / 15, 2);
    expect(page.style.padding).toMatch(/^96px 72px/);
    expect(parseFloat(page.style.minHeight)).toBeGreaterThan(1000);
    expect(paper.style.fontSize).toBe('11pt');
    expect(paper.style.fontFamily).toContain('Calibri');
  });

  it('renders headings from style names and outline levels', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const heading = paper.querySelector('h1')!;
    expect(heading.textContent).toBe('Annual Report');
    const run = heading.querySelector<HTMLElement>('span')!;
    expect(run.style.fontWeight).toBe('bold');
    expect(run.style.fontSize).toBe('16pt');
    expect(paper.querySelector('h3')?.textContent).toBe('Outline heading');
  });

  it('applies character formatting, with direct formatting overriding styles', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const styled = spanWith(paper, 'styled');
    expect(styled.style.fontWeight).toBe('bold');
    expect(styled.style.fontStyle).toBe('italic');
    expect(styled.style.textDecoration).toBe('underline line-through');
    expect(styled.style.color.toLowerCase()).toBe('#ff0000');
    expect(styled.style.fontSize).toBe('14pt');
    expect(styled.style.backgroundColor).toBe('#ffff00');
    expect(styled.style.fontFamily).toContain('Arial');

    const superscript = spanWith(paper, 'sup');
    expect(superscript.style.verticalAlign).toBe('super');
    expect(superscript.style.fontWeight).not.toBe('bold');

    const characterStyle = spanWith(paper, 'char style');
    expect(characterStyle.style.fontStyle).toBe('italic');
    expect(characterStyle.style.textTransform).toBe('uppercase');
    expect(characterStyle.style.fontVariant).toBe('small-caps');
  });

  it('hides vanished text, deletions and field codes but keeps insertions', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const text = paper.textContent ?? '';
    expect(text).not.toContain('hidden text');
    expect(text).not.toContain('deleted');
    expect(text).not.toContain('FIELD CODE');
    expect(text).toContain('inserted');
    expect(text).toContain('content control');
    expect(text).toContain('block control');
  });

  it('renders tabs, breaks and symbols', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const paragraph = paragraphWith(paper, 'Plain');
    expect(paragraph.style.textAlign).toBe('center');
    expect(paragraph.textContent).toContain('\t');
    expect(paragraph.textContent).toContain('‑');
    expect(paragraph.textContent).toContain('✓');
    expect(paragraph.querySelectorAll('br')).toHaveLength(2);
  });

  it('applies paragraph indents, spacing, shading and alignment', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const paragraph = paragraphWith(paper, 'external link');
    expect(parseFloat(paragraph.style.marginLeft)).toBeCloseTo(48, 1);
    expect(parseFloat(paragraph.style.marginRight)).toBeCloseTo(24, 1);
    expect(parseFloat(paragraph.style.textIndent)).toBeCloseTo(16, 1);
    expect(parseFloat(paragraph.style.marginTop)).toBeCloseTo(8, 1);
    expect(parseFloat(paragraph.style.marginBottom)).toBeCloseTo(16, 1);
    // 1.5 lines of Calibri, whose single line is 1.2207 times the font size.
    expect(Number(paragraph.style.lineHeight)).toBeCloseTo(1.5 * 1.2207, 2);
    expect(paragraph.style.backgroundColor.toLowerCase()).toBe('#eeeeee');
    expect(paragraph.style.textAlign).toBe('justify');

    const exact = paragraphWith(paper, 'after break');
    expect(exact.style.lineHeight).toBe('14pt');
    expect(exact.style.textAlign).toBe('right');
    // Paragraph defaults apply where nothing overrides them.
    expect(parseFloat(paper.querySelector<HTMLElement>('h1')!.style.marginBottom)).toBeCloseTo(
      10.7,
      1,
    );
  });

  it('keeps safe external links and neutralises everything else', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const links = paper.querySelectorAll('a[target]');
    expect(links).toHaveLength(1);
    // Links to bookmarks stay inside the document.
    expect(paper.querySelectorAll('a[data-av-anchor]')).toHaveLength(1);
    expect(links[0]!.getAttribute('href')).toBe('https://example.com/');
    expect(links[0]!.getAttribute('rel')).toContain('noopener');
    expect(paper.textContent).toContain('unsafe link');
    expect(paper.textContent).toContain('internal link');
  });

  it('numbers lists across levels and restarts', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const label = (text: string): string =>
      paragraphWith(paper, text).querySelector('span')!.textContent ?? '';
    expect(label('first')).toBe('1.');
    expect(label('nested a')).toBe('1.a)');
    expect(label('nested b')).toBe('1.b)');
    expect(label('second')).toBe('2.');
    expect(label('bullet')).toBe('•');
    expect(label('restarted')).toBe('5.');
    expect(label('not numbered')).toBe('not numbered');
    expect(label('unknown list')).toBe('unknown list');

    // Level indents come from the numbering definition unless the paragraph sets its own.
    expect(parseFloat(paragraphWith(paper, 'first').style.marginLeft)).toBeCloseTo(48, 1);
    expect(parseFloat(paragraphWith(paper, 'first').style.textIndent)).toBeCloseTo(-24, 1);
    expect(parseFloat(paragraphWith(paper, 'bullet').style.marginLeft)).toBeCloseTo(96, 1);
  });

  it('renders tables with spans, merges, borders and shading', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const [first, second] = Array.from(paper.querySelectorAll('table'));
    expect(first!.querySelectorAll('col')).toHaveLength(3);
    expect(first!.style.marginLeft).toBe('auto');
    expect(first!.style.marginRight).toBe('auto');
    const rows = first!.querySelectorAll('tr');
    const wide = rows[0]!.children[0] as HTMLTableCellElement;
    const tall = rows[0]!.children[1] as HTMLTableCellElement;
    expect(wide.colSpan).toBe(2);
    expect(wide.style.backgroundColor.toLowerCase()).toBe('#ddeeff');
    expect(wide.style.verticalAlign).toBe('middle');
    expect(wide.style.borderTop).toContain('solid');
    expect(tall.rowSpan).toBe(2);
    expect(tall.style.verticalAlign).toBe('bottom');
    // The continuation cell of the vertical merge is not drawn.
    expect(rows[1]!.children).toHaveLength(2);

    expect(second!.style.marginLeft).toBe('auto');
    expect(second!.querySelector<HTMLElement>('td')!.style.borderTop).not.toContain('solid');
  });

  it('embeds supported pictures and reserves space for unsupported ones', async () => {
    const fake = host();
    const paper = (await renderDocx(openZip(docxFixture()), fake)).element;
    const images = paper.querySelectorAll('img');
    expect(images).toHaveLength(2);
    expect(images[0]!.getAttribute('src')).toBe('blob:test/1');
    expect(images[0]!.style.width).toBe('100px');
    expect(images[0]!.style.height).toBe('50px');
    expect(fake.blobs[0]!.type).toBe('image/png');
    const placeholder = Array.from(paper.querySelectorAll<HTMLElement>('span')).find((span) =>
      span.style.border.includes('dashed'),
    );
    expect(placeholder?.style.width).toBe('100px');
  });

  it('renders text boxes once', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    expect(paper.textContent?.match(/text box/g)).toHaveLength(1);
  });

  it('collects referenced footnotes at the end', async () => {
    const paper = (await renderDocx(openZip(docxFixture()), host())).element;
    const notes = paper.querySelector('.av-doc-footnotes')!;
    expect(notes.textContent).toBe('The footnote text.');
    expect(notes.querySelector('[data-av-sep]')).not.toBeNull();
    expect(paragraphWith(paper, 'external link').querySelector('[data-av-fn]')?.textContent).toBe(
      '1',
    );
  });

  it('renders a document without styles, numbering or footnotes', async () => {
    const document = await renderDocx(
      openZip(docxFixture({ minimal: true, body: '<w:p><w:r><w:t>bare</w:t></w:r></w:p>' })),
      host(),
    );
    const paper = document.element;
    expect(paper.textContent).toBe('bare');
    expect(document.pageCount()).toBe(1);
    expect(document.page(0)!.style.width).toBe('816px');
    expect(paper.querySelector<HTMLElement>('.av-doc-footnotes')!.hidden).toBe(true);
  });

  it('rejects packages without a main document', async () => {
    await expect(renderDocx(openZip(makeZip({ 'a.txt': 'x' })), host())).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('marks the document for table gridlines only when the option is set', async () => {
    const plain = renderContext(docxFixture(), 'report.docx', 'docx');
    instanceOf(await docxRenderer.render(plain.context)).destroy!();
    expect(plain.container.querySelector('.av-doc-gridlines')).toBeNull();

    const base = renderContext(docxFixture(), 'report.docx', 'docx');
    const context = { ...base.context, options: { ...base.context.options, tableGridlines: true } };
    instanceOf(await docxRenderer.render(context)).destroy!();
    expect(base.container.querySelector('.av-doc.av-doc-gridlines')).not.toBeNull();
  });

  it('renders on a zoomable sheet of paper', async () => {
    const { context, container } = renderContext(docxFixture(), 'report.docx', 'docx');
    const instance = instanceOf(await docxRenderer.render(context));
    const paper = container.querySelector<HTMLElement>('.av-doc')!;
    expect(paper.textContent).toContain('Annual Report');
    instance.zoom!.set(1.5);
    expect(instance.zoom!.get()).toBe(1.5);
    expect(paper.style.getPropertyValue('zoom')).toBe('1.5');
    expect(instance.printable!()).toBe(paper);
    // The fixture contains two page breaks.
    expect(instance.pages!.count()).toBe(3);
    expect(instance.pages!.current()).toBe(1);
    instance.pages!.goTo(1);
    instance.destroy!();
  });
});

describe('ODT', () => {
  it('applies the page layout', async () => {
    const paper = await renderOdt(openZip(odtFixture()), host());
    expect(paper.style.width).toBe('794px');
    expect(paper.style.minHeight).toBe('1123px');
    expect(paper.style.padding).toBe('76px 96px');
  });

  it('renders headings with inherited and automatic styles', async () => {
    const paper = await renderOdt(openZip(odtFixture()), host());
    const heading = paper.querySelector<HTMLElement>('h2')!;
    expect(heading.textContent).toBe('Chapter');
    expect(heading.style.textAlign).toBe('center');
    expect(heading.style.color).toBe('#336699');
    expect(heading.style.fontWeight).toBe('bold');
    expect(heading.style.fontSize).toBe('12pt');
    expect(parseFloat(heading.style.marginLeft)).toBeCloseTo(37.8, 1);
    expect(parseFloat(heading.style.textIndent)).toBeCloseTo(18.9, 1);
  });

  it('renders inline formatting, whitespace and links', async () => {
    const paper = await renderOdt(openZip(odtFixture()), host());
    const paragraph = paragraphWith(paper, 'Plain');
    expect(paragraph.style.textAlign).toBe('right');
    const styled = spanWith(paper, 'styled');
    expect(styled.style.fontWeight).toBe('bold');
    expect(styled.style.fontStyle).toBe('italic');
    expect(styled.style.textDecoration).toBe('underline line-through');
    expect(styled.style.backgroundColor).toBe('#ffff00');
    expect(styled.style.fontSize).toBe('14pt');
    expect(styled.style.verticalAlign).toBe('super');
    // Values that are not plain colours never reach the style attribute.
    const sub = spanWith(paper, 'sub');
    expect(sub.style.verticalAlign).toBe('sub');
    expect(sub.style.color).toBe('');
    expect(paragraph.textContent).toContain('sub   spaced\ttabbed');
    expect(paragraph.querySelectorAll('br')).toHaveLength(1);
    const links = paragraph.querySelectorAll('a');
    expect(links).toHaveLength(1);
    expect(links[0]!.getAttribute('href')).toBe('https://example.com/');
    expect(paragraph.querySelector('sup')?.textContent).toBe('1');
    expect(paragraph.textContent).not.toContain('note body');
    expect(paragraph.textContent).toContain('5');
    expect(paper.textContent).not.toContain('removed change');
    expect(paper.textContent).toContain('in section');
  });

  it('embeds packaged images only', async () => {
    const fake = host();
    const paper = await renderOdt(openZip(odtFixture()), fake);
    const images = paper.querySelectorAll('img');
    expect(images).toHaveLength(1);
    expect(images[0]!.getAttribute('src')).toBe('blob:test/1');
    expect(images[0]!.style.width).toBe('76px');
    expect(paper.textContent).toContain('boxed text');
  });

  it('renders ordered, bulleted and nested lists', async () => {
    const paper = await renderOdt(openZip(odtFixture()), host());
    const ordered = paper.querySelector('ol')!;
    expect(ordered.getAttribute('type')).toBe('a');
    expect(ordered.children).toHaveLength(2);
    expect(ordered.querySelector('li > ul')?.textContent).toBe('nested');
    expect(paper.querySelectorAll(':scope > ul')).toHaveLength(1);
  });

  it('renders tables with spans', async () => {
    const paper = await renderOdt(openZip(odtFixture()), host());
    const cells = paper.querySelectorAll('td');
    expect(cells).toHaveLength(3);
    expect(cells[0]!.colSpan).toBe(2);
    expect(cells[1]!.rowSpan).toBe(2);
  });

  it('works without styles.xml and rejects packages without content', async () => {
    const bare = makeZip({
      'content.xml':
        '<office:document-content><office:body><office:text><text:p>bare</text:p></office:text></office:body></office:document-content>',
    });
    expect((await renderOdt(openZip(bare), host())).textContent).toBe('bare');
    await expect(renderOdt(openZip(makeZip({ 'a.txt': 'x' })), host())).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('renders through the paper view', async () => {
    const { context, container } = renderContext(odtFixture(), 'doc.odt', 'odt');
    await odtRenderer.render(context);
    expect(container.querySelector('.av-paper h2')?.textContent).toBe('Chapter');
  });
});

describe('RTF', () => {
  const render = (source: string) => {
    const fake = host();
    return { paper: renderRtf(new TextEncoder().encode(source), fake), fake };
  };

  it('rejects files that are not RTF', () => {
    expect(() => render('hello')).toThrow('not a valid RTF');
  });

  it('skips metadata and unknown destinations', () => {
    const { paper } = render(RTF_FIXTURE);
    const text = paper.textContent ?? '';
    for (const hidden of [
      'Arial',
      'Normal',
      'Hidden title',
      'Nobody',
      'Test;',
      'ignore me',
      'HYPERLINK',
      'hidden',
    ]) {
      expect(text).not.toContain(hidden);
    }
    expect(text).toContain('link text');
  });

  it('applies character formatting scoped to groups', () => {
    const { paper } = render(RTF_FIXTURE);
    const heading = spanWith(paper, 'Heading');
    expect(heading.style.fontWeight).toBe('bold');
    expect(heading.style.fontSize).toBe('16pt');
    expect(heading.parentElement!.style.textAlign).toBe('center');

    expect(spanWith(paper, 'italic').style.fontStyle).toBe('italic');
    expect(spanWith(paper, 'under').style.textDecoration).toBe('underline');
    expect(spanWith(paper, ' none ').style.textDecoration).toBe('');
    expect(spanWith(paper, 'struck').style.textDecoration).toBe('line-through');
    expect(spanWith(paper, 'red').style.color).toBe('rgb(255, 0, 0)');
    expect(spanWith(paper, 'marked').style.backgroundColor).toBe('rgb(0, 0, 255)');
    expect(spanWith(paper, 'up').style.verticalAlign).toBe('super');
    expect(spanWith(paper, 'down').style.verticalAlign).toBe('sub');
    expect(spanWith(paper, 'flat').style.verticalAlign).toBe('');
    expect(spanWith(paper, 'bold after plain').style.fontWeight).toBe('bold');
    // Formatting ends with its group.
    expect(spanWith(paper, ' justified').style.fontWeight).toBe('');
  });

  it('applies paragraph alignment, indents and spacing', () => {
    const { paper } = render(RTF_FIXTURE);
    const paragraph = paragraphWith(paper, 'Plain');
    expect(parseFloat(paragraph.style.marginLeft)).toBeCloseTo(48, 1);
    expect(parseFloat(paragraph.style.textIndent)).toBeCloseTo(-24, 1);
    expect(parseFloat(paragraph.style.marginTop)).toBeCloseTo(8, 1);
    expect(parseFloat(paragraph.style.marginBottom)).toBeCloseTo(16, 1);
    expect(paragraphWith(paper, 'Caf').style.textAlign).toBe('right');
    expect(paragraphWith(paper, 'justified').style.textAlign).toBe('justify');
    expect(paragraphWith(paper, 'link text').style.textAlign).toBe('');
  });

  it('decodes code-page bytes, Unicode escapes and special characters', () => {
    const { paper } = render(RTF_FIXTURE);
    const paragraph = paragraphWith(paper, 'Caf');
    expect(paragraph.textContent).toContain('Café €uro 😀 {braces} \\ back nb‑hy—–‘’“”•\ttabbed');
    expect(paragraph.querySelectorAll('br')).toHaveLength(1);
    expect(paragraph.textContent).toContain('broken');
  });

  it('builds tables from rows and cells', () => {
    const { paper } = render(RTF_FIXTURE);
    const tables = paper.querySelectorAll('table');
    expect(tables).toHaveLength(1);
    const rows = tables[0]!.querySelectorAll('tr');
    expect(rows).toHaveLength(2);
    expect(Array.from(rows[0]!.querySelectorAll('td'), (cell) => cell.textContent)).toEqual([
      'Cell A',
      'Cell Bsecond line',
    ]);
    expect(rows[0]!.children[1]!.querySelectorAll('p')).toHaveLength(2);
    // Text after the table is back in the normal flow.
    expect(paragraphWith(paper, 'After table').closest('table')).toBeNull();
  });

  it('embeds PNG and JPEG pictures and ignores other kinds', () => {
    const { paper, fake } = render(RTF_FIXTURE);
    const images = paper.querySelectorAll('img');
    expect(images).toHaveLength(2);
    expect(fake.blobs.map((blob) => blob.type)).toEqual(['image/png', 'image/jpeg']);
    expect(fake.blobs[0]!.size).toBe(PNG_1X1.length);
    expect(images[0]!.style.width).toBe('100px');
    expect(images[0]!.style.height).toBe('50px');
  });

  it('handles section breaks, escaped newlines and binary data', () => {
    const { paper } = render(RTF_FIXTURE);
    expect(paragraphWith(paper, 'Section').textContent).toBe('Section');
    expect(paper.lastElementChild?.textContent).toBe('Line after backslash-newlineDone');
  });

  it('closes an unfinished table and survives unbalanced braces', () => {
    const { paper } = render(String.raw`{\rtf1 \intbl last cell\cell}}} trailing`);
    expect(paper.querySelector('td')?.textContent).toBe('last cell');
    expect(paper.textContent).not.toContain('trailing');
    expect(render(String.raw`{\rtf1 unclosed {\b bold`).paper.textContent).toBe('unclosed bold');
  });

  it('honours \\uc skip counts and other code pages', () => {
    // "u" + digits is spelled with a variable so it reaches the parser as RTF text.
    const zhe = String.raw`\u` + '1046';
    const { paper } = render(String.raw`{\rtf1\ansicpg1251\uc2 ${zhe} ??\'c6 \uc0${zhe} !}`);
    expect(paper.textContent).toBe('ЖЖ Ж!');
  });

  it('renders through the paper view', async () => {
    const { context, container } = renderContext(RTF_FIXTURE, 'doc.rtf', 'rtf');
    await rtfRenderer.render(context);
    expect(container.querySelector('.av-paper')?.textContent).toContain('Heading');
  });
});
