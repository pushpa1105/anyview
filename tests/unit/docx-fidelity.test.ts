import { afterEach, describe, expect, it, vi } from 'vitest';
import {
  deobfuscateFont,
  formatCounter,
  renderDocx,
  type DocxDocument,
} from '../../src/formats/docx';
import { fillFields, layoutTabs } from '../../src/formats/docx-layout';
import { lineFactor, mergeTabs, PropertyParser } from '../../src/formats/docx-props';
import {
  applySheetTint,
  parseTheme,
  safeFontName,
  shadeColor,
  tintColor,
} from '../../src/formats/theme';
import { parseXml } from '../../src/util/xml';
import { openZip } from '../../src/util/zip';
import { makeZip, PNG_1X1 } from '../helpers/zip';

const XML = '<?xml version="1.0" encoding="UTF-8"?>';
const NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape" ' +
  'xmlns:wpg="http://schemas.microsoft.com/office/word/2010/wordprocessingGroup" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

const THEME = `${XML}<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"><a:themeElements>
<a:clrScheme name="t"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:srgbClr val="FFFFFF"/></a:lt1>
<a:dk2><a:srgbClr val="1F497D"/></a:dk2><a:accent1><a:srgbClr val="4F81BD"/></a:accent1><a:hlink><a:srgbClr val="0000FF"/></a:hlink></a:clrScheme>
<a:fontScheme name="t"><a:majorFont><a:latin typeface="Cambria"/></a:majorFont><a:minorFont><a:latin typeface="Calibri"/></a:minorFont></a:fontScheme>
</a:themeElements></a:theme>`;

interface Parts {
  body: string;
  styles?: string;
  numbering?: string;
  settings?: string;
  footnotes?: string;
  endnotes?: string;
  fontTable?: string;
  /** Extra package files, e.g. headers. */
  files?: Record<string, string | Uint8Array>;
  /** Extra relationships of the main document: id, type, target, external. */
  relationships?: [string, string, string, boolean?][];
}

function relationshipsXml(entries: [string, string, string, boolean?][]): string {
  return `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries
    .map(
      ([id, type, target, external]) =>
        `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`,
    )
    .join('')}</Relationships>`;
}

function pack(parts: Parts): Uint8Array {
  const relationships: [string, string, string, boolean?][] = [
    ['rTheme', 'theme', 'theme/theme1.xml'],
    ['rStyles', 'styles', 'styles.xml'],
    ['rImage', 'image', 'media/image1.png'],
    ...(parts.relationships ?? []),
  ];
  const files: Record<string, string | Uint8Array> = {
    'word/document.xml': `${XML}<w:document ${NS}><w:body>${parts.body}</w:body></w:document>`,
    'word/theme/theme1.xml': THEME,
    'word/media/image1.png': PNG_1X1,
    'word/styles.xml': `${XML}<w:styles ${NS}>${
      parts.styles ??
      `<w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:asciiTheme="minorHAnsi"/><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
       <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:rPr><w:sz w:val="24"/></w:rPr><w:pPr><w:jc w:val="center"/></w:pPr></w:style>
       <w:style w:type="character" w:styleId="Hyperlink"><w:name w:val="Hyperlink"/><w:rPr><w:color w:val="0000FF"/><w:u w:val="single"/></w:rPr></w:style>`
    }</w:styles>`,
    ...parts.files,
  };
  const optional: [keyof Parts, string, string][] = [
    ['numbering', 'numbering', 'numbering.xml'],
    ['settings', 'settings', 'settings.xml'],
    ['footnotes', 'footnotes', 'footnotes.xml'],
    ['endnotes', 'endnotes', 'endnotes.xml'],
    ['fontTable', 'fontTable', 'fontTable.xml'],
  ];
  for (const [key, type, target] of optional) {
    const content = parts[key];
    if (typeof content !== 'string') continue;
    relationships.push([`r_${type}`, type, target]);
    files[`word/${target}`] =
      `${XML}<w:${type === 'fontTable' ? 'fonts' : type} ${NS}>${content}</w:${type === 'fontTable' ? 'fonts' : type}>`;
  }
  files['word/_rels/document.xml.rels'] = relationshipsXml(relationships);
  return makeZip(files);
}

let blobs = 0;
const host = { objectUrl: (): string => `blob:test/${++blobs}` };

async function render(parts: Parts): Promise<DocxDocument> {
  return renderDocx(openZip(pack(parts)), host);
}

const p = (text: string, properties = '', runProperties = ''): string =>
  `<w:p>${properties ? `<w:pPr>${properties}</w:pPr>` : ''}<w:r>${runProperties ? `<w:rPr>${runProperties}</w:rPr>` : ''}<w:t xml:space="preserve">${text}</w:t></w:r></w:p>`;

function paragraphWith(root: Element, text: string): HTMLElement {
  const found = Array.from(root.querySelectorAll<HTMLElement>('p,h1,h2,h3')).find((element) =>
    element.textContent?.includes(text),
  );
  if (!found) throw new Error(`No paragraph contains "${text}"`);
  return found;
}

function spanWith(root: Element, text: string): HTMLElement {
  const found = Array.from(root.querySelectorAll<HTMLElement>('span')).find(
    (element) => element.textContent === text && element.children.length === 0,
  );
  if (!found) throw new Error(`No span is "${text}"`);
  return found;
}

afterEach(() => {
  vi.restoreAllMocks();
  document.body.textContent = '';
});

describe('theme helpers', () => {
  it('parses colours and fonts and rejects unsafe font names', () => {
    const theme = parseTheme(parseXml(THEME));
    expect(theme.colors.get('dk1')).toBe('#000000');
    expect(theme.colors.get('accent1')).toBe('#4F81BD');
    expect(theme.majorFont).toBe('Cambria');
    expect(parseTheme(null).colors.size).toBe(0);
    expect(safeFontName('Segoe UI')).toBe('Segoe UI');
    expect(safeFontName('x"; background:url(evil)')).toBeUndefined();
    expect(safeFontName(undefined)).toBeUndefined();
  });

  it('tints and shades colours', () => {
    expect(tintColor('#000000', 0.5)).toBe('#808080');
    expect(shadeColor('#ffffff', 0.5)).toBe('#808080');
    expect(applySheetTint('#4f81bd', 0)).toBe('#4f81bd');
    expect(applySheetTint('#808080', 0.5)).toBe('#c0c0c0');
    expect(applySheetTint('#808080', -0.5)).toBe('#404040');
    // Hue and saturation survive a tint; only the lightness changes.
    const lighter = applySheetTint('#4f81bd', 0.4);
    expect(parseInt(lighter.slice(5, 7), 16)).toBeGreaterThan(0xbd);
    for (const colour of ['#ff0000', '#00ff00', '#0000ff', '#804020']) {
      expect(applySheetTint(colour, 0.2)).toMatch(/^#[0-9a-f]{6}$/);
      expect(applySheetTint(colour, -0.2)).toMatch(/^#[0-9a-f]{6}$/);
    }
  });
});

describe('property parsing', () => {
  const parser = new PropertyParser(parseTheme(parseXml(THEME)));
  const element = (xml: string) => parseXml(`<root ${NS}>${xml}</root>`).children[0] as never;

  it('resolves theme colours with tints, shades and fills', () => {
    expect(parser.color(element('<w:color w:val="FF0000"/>'))).toBe('#FF0000');
    expect(parser.color(element('<w:color w:val="auto"/>'))).toBeUndefined();
    expect(parser.color(element('<w:color w:themeColor="accent1"/>'))).toBe('#4F81BD');
    expect(parser.color(element('<w:color w:themeColor="text1" w:themeTint="80"/>'))).toBe(
      '#7f7f7f',
    );
    expect(parser.color(element('<w:color w:themeColor="background1" w:themeShade="80"/>'))).toBe(
      '#808080',
    );
    expect(parser.color(element('<w:color w:themeColor="unknown" w:val="00FF00"/>'))).toBe(
      '#00FF00',
    );
    expect(parser.shading(element('<w:shd w:val="clear" w:fill="EEEEEE"/>'))).toBe('#EEEEEE');
    expect(parser.shading(element('<w:shd w:val="solid" w:color="112233"/>'))).toBe('#112233');
    expect(
      parser.shading(element('<w:shd w:val="clear" w:themeFill="accent1" w:themeFillTint="33"/>')),
    ).toMatch(/^#/);
    expect(parser.color(null)).toBeUndefined();
    expect(parser.shading(null)).toBeUndefined();
  });

  it('reads fonts, borders and run properties', () => {
    expect(parser.font(element('<w:rFonts w:asciiTheme="majorHAnsi"/>'))).toBe('Cambria');
    expect(parser.font(element('<w:rFonts w:ascii="Arial"/>'))).toBe('Arial');
    expect(parser.font(null)).toBeUndefined();
    expect(parser.border(element('<w:top w:val="nil"/>'))?.style).toBe('none');
    expect(parser.border(element('<w:top w:val="double" w:sz="4"/>'))).toMatchObject({
      style: 'double',
      width: 3,
    });
    expect(
      parser.border(element('<w:top w:val="dashed" w:sz="12" w:space="4" w:color="FF0000"/>')),
    ).toMatchObject({
      style: 'dashed',
      width: 2,
      color: '#FF0000',
    });
    expect(parser.border(element('<w:top/>'))).toBeUndefined();
    const run = parser.run(
      element(
        '<w:rPr><w:u w:val="wave" w:color="FF0000"/><w:dstrike/><w:spacing w:val="30"/><w:position w:val="6"/><w:bdr w:val="single"/><w:shadow/><w:highlight w:val="none"/><w:shd w:val="clear" w:fill="ABCDEF"/><w:u w:val="none"/></w:rPr>',
      ),
    );
    expect(run).toMatchObject({
      underline: 'none',
      doubleStrike: true,
      spacing: 2,
      position: 3,
      shadow: true,
    });
    expect(run.highlight).toBe('#ABCDEF');
    expect(run.border?.style).toBe('solid');
  });

  it('reads paragraph, cell, table and style properties', () => {
    const para = parser.para(
      element(
        '<w:pPr><w:spacing w:beforeAutospacing="1" w:afterAutospacing="0" w:after="100"/><w:keepNext/><w:keepLines/><w:widowControl w:val="0"/><w:contextualSpacing/><w:bidi/><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/><w:tab w:val="left"/></w:tabs><w:framePr w:dropCap="drop" w:lines="3" w:hSpace="30"/><w:pBdr><w:start w:val="single"/><w:end w:val="single"/><w:between w:val="single"/></w:pBdr></w:pPr>',
      ),
    );
    expect(para).toMatchObject({
      spaceBefore: 280,
      spaceAfter: 100,
      keepNext: true,
      keepLines: true,
      widowControl: false,
      contextualSpacing: true,
      rtl: true,
    });
    expect(para.tabs).toEqual([{ pos: 600, align: 'right', leader: 'dot' }]);
    expect(para.frame).toMatchObject({ dropCap: 'drop', lines: 3, hSpace: 2 });
    expect(Object.keys(para.borders ?? {})).toEqual(['left', 'right', 'between']);

    const cell = parser.cell(
      element(
        '<w:tcPr><w:tcMar><w:start w:w="150" w:type="dxa"/><w:top w:w="30"/></w:tcMar><w:noWrap/><w:textDirection w:val="btLr"/><w:vAlign w:val="center"/></w:tcPr>',
      ),
    );
    expect(cell).toMatchObject({
      margins: { left: 10, top: 2 },
      noWrap: true,
      textDirection: 'btLr',
      vAlign: 'center',
    });
    const table = parser.table(
      element(
        '<w:tblPr><w:tblInd w:w="300" w:type="dxa"/><w:jc w:val="center"/><w:tblStyleRowBandSize w:val="2"/><w:tblStyleColBandSize w:val="3"/><w:tblCellMar><w:left w:w="15"/></w:tblCellMar></w:tblPr>',
      ),
    );
    expect(table).toMatchObject({
      indent: 300,
      align: 'center',
      rowBandSize: 2,
      colBandSize: 3,
      margins: { left: 1 },
    });
    expect(
      parser.style(element('<w:style w:type="table"><w:name w:val="x"/></w:style>')),
    ).toBeNull();
    expect(parser.cell(null)).toEqual({});
    expect(parser.table(null)).toEqual({});
  });

  it('merges tab stops and knows line heights of common fonts', () => {
    const merged = mergeTabs([
      [
        { pos: 100, align: 'left', leader: 'none' },
        { pos: 200, align: 'left', leader: 'none' },
      ],
      undefined,
      [
        { pos: 100, align: 'clear', leader: 'none' },
        { pos: 50, align: 'center', leader: 'dot' },
      ],
    ]);
    expect(merged.map((stop) => stop.pos)).toEqual([50, 200]);
    expect(lineFactor('Calibri')).toBeCloseTo(1.2207);
    expect(lineFactor('Some Unknown Font')).toBe(1.2);
    expect(lineFactor(undefined)).toBe(1.2);
    expect(formatCounter(2, 'ordinal')).toBe('2nd');
    expect(formatCounter(11, 'ordinal')).toBe('11th');
    expect(formatCounter(24, 'ordinal')).toBe('24th');
  });
});

describe('DOCX text and paragraph fidelity', () => {
  it('applies theme fonts, colours and character effects', async () => {
    const { element } = await render({
      body:
        p(
          'themed',
          '',
          '<w:color w:themeColor="accent1"/><w:rFonts w:asciiTheme="majorHAnsi"/><w:u w:val="wave" w:color="FF0000"/><w:dstrike/><w:spacing w:val="30"/><w:position w:val="6"/><w:bdr w:val="single" w:sz="4"/><w:shadow/>',
        ) +
        p('sup', '', '<w:vertAlign w:val="superscript"/>') +
        '<w:p><w:r><w:t>a</w:t><w:softHyphen/><w:noBreakHyphen/><w:sym w:font="Wingdings" w:char="F0FC"/><w:cr/><w:br/></w:r><m:oMath><m:r><m:t>x+1</m:t></m:r></m:oMath></w:p>',
    });
    expect(element.style.fontFamily).toContain('Calibri');
    const themed = spanWith(element, 'themed');
    expect(themed.style.color.toLowerCase()).toBe('#4f81bd');
    expect(themed.style.fontFamily).toContain('Cambria');
    expect(themed.style.fontFamily).toContain('Caladea');
    expect(themed.style.textDecoration).toBe('underline line-through');
    expect(themed.style.textDecorationStyle).toBe('wavy');
    expect(themed.style.letterSpacing).toBe('2px');
    expect(themed.style.top).toBe('-3pt');
    expect(themed.style.border).toContain('solid');
    expect(themed.style.textShadow).not.toBe('');
    expect(spanWith(element, 'sup').style.lineHeight).toBe('0');
    expect(element.textContent).toContain('✓');
    expect(element.textContent).toContain('x+1');
    expect(element.querySelectorAll('br').length).toBe(2);
  });

  it('lays out paragraph borders, frames, breaks and spacing rules', async () => {
    const { element } = await render({
      body:
        p(
          'bordered',
          '<w:pBdr><w:left w:val="single" w:sz="8" w:space="4"/><w:bottom w:val="single" w:sz="8" w:space="1"/><w:top w:val="nil"/></w:pBdr><w:ind w:left="150"/>',
        ) +
        p(
          'D',
          '<w:framePr w:dropCap="drop" w:lines="3" w:hSpace="30"/><w:spacing w:line="900" w:lineRule="exact"/>',
        ) +
        p('side', '<w:framePr w:xAlign="right" w:w="1500" w:hSpace="150"/>') +
        p(
          'atleast',
          '<w:spacing w:line="300" w:lineRule="atLeast"/><w:keepNext/><w:keepLines/><w:widowControl w:val="0"/><w:pageBreakBefore/><w:bidi/>',
          '<w:sz w:val="40"/>',
        ) +
        p(
          'ctx one',
          '<w:pStyle w:val="Normal"/><w:contextualSpacing/><w:spacing w:before="200" w:after="200"/>',
        ) +
        p(
          'ctx two',
          '<w:pStyle w:val="Normal"/><w:contextualSpacing/><w:spacing w:before="200" w:after="200"/>',
        ) +
        '<w:p/>',
    });
    const bordered = paragraphWith(element, 'bordered');
    expect(bordered.style.borderLeft).toContain('solid');
    expect(bordered.style.borderTop).toBe('');
    expect(parseFloat(bordered.style.paddingLeft)).toBeCloseTo(5.33, 1);
    // The border sits outside the text area: 10px indent minus space and width.
    expect(parseFloat(bordered.style.marginLeft)).toBeLessThan(10);
    const cap = paragraphWith(element, 'D');
    expect(cap.style.float).toBe('left');
    expect(parseFloat(cap.style.height)).toBeCloseTo(59, 0);
    expect(parseFloat(cap.style.lineHeight)).toBe(45);
    const side = paragraphWith(element, 'side');
    expect(side.style.float).toBe('right');
    expect(side.style.width).toBe('100px');
    const atLeast = paragraphWith(element, 'atleast');
    expect(atLeast.getAttribute('data-av-break')).toBe('before');
    expect(atLeast.hasAttribute('data-av-kn')).toBe(true);
    expect(atLeast.hasAttribute('data-av-kl')).toBe(true);
    expect(atLeast.hasAttribute('data-av-nw')).toBe(true);
    expect(atLeast.style.direction).toBe('rtl');
    expect(atLeast.style.fontSize).toBe('20pt');
    expect(paragraphWith(element, 'ctx one').style.marginBottom).toBe('0px');
    expect(paragraphWith(element, 'ctx two').style.marginTop).toBe('0px');
    // An empty paragraph still occupies a line.
    expect(element.querySelector('p:last-of-type br')).not.toBeNull();
  });

  it('numbers lists like Word: shared counters, alignment, pictures and symbols', async () => {
    const level = (index: number, inner: string): string =>
      `<w:lvl w:ilvl="${index}">${inner}</w:lvl>`;
    const { element } = await render({
      numbering:
        `<w:numPicBullet w:numPicBulletId="0"><w:pict><v:shape><v:imagedata r:id="rPic"/></v:shape></w:pict></w:numPicBullet>` +
        `<w:abstractNum w:abstractNumId="0">${level(0, '<w:start w:val="1"/><w:numFmt w:val="lowerRoman"/><w:lvlText w:val="%1."/><w:lvlJc w:val="right"/><w:pPr><w:ind w:left="720" w:hanging="180"/></w:pPr>')}${level(1, '<w:numFmt w:val="bullet"/><w:lvlText w:val="o"/><w:suff w:val="space"/><w:rPr><w:rFonts w:ascii="Courier New"/></w:rPr>')}${level(2, '<w:numFmt w:val="bullet"/><w:lvlText w:val=""/><w:lvlPicBulletId w:val="0"/>')}${level(3, '<w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0A7;"/><w:lvlJc w:val="center"/><w:rPr><w:rFonts w:ascii="Wingdings"/></w:rPr>')}</w:abstractNum>` +
        `<w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num><w:num w:numId="2"><w:abstractNumId w:val="0"/></w:num>` +
        `<w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="1"/></w:lvlOverride></w:num>`,
      files: {
        'word/_rels/numbering.xml.rels': relationshipsXml([['rPic', 'image', 'media/image1.png']]),
      },
      body:
        p('one', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
        p('two', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr>') +
        p('letter', '<w:numPr><w:ilvl w:val="1"/><w:numId w:val="2"/></w:numPr>') +
        p('picture', '<w:numPr><w:ilvl w:val="2"/><w:numId w:val="2"/></w:numPr>') +
        p('square', '<w:numPr><w:ilvl w:val="3"/><w:numId w:val="2"/></w:numPr>') +
        p('again', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr>'),
    });
    const marker = (text: string): HTMLElement =>
      paragraphWith(element, text).querySelector<HTMLElement>('.av-doc-marker')!;
    expect(marker('one').textContent).toBe('i.');
    // A second list on the same definition continues the numbering.
    expect(marker('two').textContent).toBe('ii.');
    expect((marker('one').firstElementChild as HTMLElement).style.transform).toBe(
      'translateX(-100%)',
    );
    expect(marker('letter').textContent).toBe('o');
    expect(marker('letter').hasAttribute('data-av-suff')).toBe(false);
    expect(marker('picture').querySelector('img')!.getAttribute('src')).toMatch(/^blob:test\//);
    expect(marker('square').textContent).toBe('▪');
    expect((marker('square').firstElementChild as HTMLElement).style.transform).toBe(
      'translateX(-50%)',
    );
    expect(marker('again').textContent).toBe('i.');
    expect(paragraphWith(element, 'one').hasAttribute('data-av-tabs')).toBe(true);
  });

  it('handles fields, bookmarks, links and tab stops', async () => {
    const field = (instruction: string, result: string): string =>
      `<w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText xml:space="preserve">${instruction}</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r>${result}<w:r><w:fldChar w:fldCharType="end"/></w:r>`;
    const document = await render({
      body:
        `<w:p><w:pPr><w:tabs><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr>${field(
          ' TOC \\o "1-3" ',
          '<w:hyperlink w:anchor="target"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>Entry</w:t></w:r><w:r><w:tab/></w:r><w:r><w:t>7</w:t></w:r></w:hyperlink>',
        )}</w:p>` +
        `<w:p>${field(' HYPERLINK "https://example.com/a" ', '<w:r><w:t>web</w:t></w:r>')}${field(' HYPERLINK \\l "target" ', '<w:r><w:t>jump</w:t></w:r>')}${field(' HYPERLINK "javascript:alert(1)" ', '<w:r><w:t>bad</w:t></w:r>')}</w:p>` +
        `<w:p>${field(' PAGE ', '<w:r><w:t>99</w:t></w:r>')} of <w:fldSimple w:instr=" NUMPAGES "><w:r><w:rPr><w:b/></w:rPr><w:t>88</w:t></w:r></w:fldSimple><w:fldSimple w:instr=" AUTHOR "><w:r><w:t>Ann</w:t></w:r></w:fldSimple><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText> PAGE </w:instrText></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r></w:p>` +
        `<w:p><w:bookmarkStart w:id="1" w:name="target"/><w:bookmarkStart w:id="2" w:name="_GoBack"/><w:r><w:t>Destination</w:t><w:ptab w:alignment="right" w:leader="dot"/></w:r><w:hyperlink r:id="rLink"><w:r><w:rPr><w:rStyle w:val="Hyperlink"/></w:rPr><w:t>site</w:t></w:r></w:hyperlink><w:hyperlink><w:r><w:t>plain</w:t></w:r></w:hyperlink></w:p>`,
      relationships: [['rLink', 'hyperlink', 'https://example.com/', true]],
    });
    const { element } = document;
    // Entries of a table of contents are not shown as blue links.
    const entry = spanWith(element, 'Entry');
    expect(entry.style.color).toBe('');
    expect(entry.closest('a')!.getAttribute('data-av-anchor')).toBe('target');
    const toc = paragraphWith(element, 'Entry');
    expect(toc.getAttribute('data-av-tabs')).toBe('600:right:dot');
    expect(toc.querySelector('.av-doc-tab')!.textContent).toBe('\t');

    expect(spanWith(element, 'web').closest('a')!.getAttribute('href')).toBe(
      'https://example.com/a',
    );
    expect(spanWith(element, 'jump').closest('a')!.getAttribute('data-av-anchor')).toBe('target');
    expect(spanWith(element, 'bad').closest('a')).toBeNull();
    expect(spanWith(element, 'site').style.color.toLowerCase()).toBe('#0000ff');
    expect(spanWith(element, 'site').closest('a')!.getAttribute('target')).toBe('_blank');
    expect(spanWith(element, 'plain').closest('a')).toBeNull();

    // Page fields show live numbers instead of the stale cached result.
    const fields = Array.from(element.querySelectorAll('[data-av-field]'));
    expect(fields.map((item) => item.getAttribute('data-av-field'))).toEqual([
      'page',
      'pages',
      'page',
    ]);
    expect(element.textContent).not.toContain('99');
    expect(element.textContent).not.toContain('88');
    expect(element.textContent).toContain('Ann');
    expect(element.querySelectorAll('[data-av-bookmark]')).toHaveLength(1);
    expect(element.querySelector('[data-av-ptab]')!.getAttribute('data-av-leader')).toBe('dot');

    // Clicking an internal link scrolls to its bookmark.
    const bookmark = element.querySelector<HTMLElement>('[data-av-bookmark]')!;
    const scroll = vi.fn();
    bookmark.scrollIntoView = scroll;
    spanWith(element, 'jump').dispatchEvent(
      new MouseEvent('click', { bubbles: true, cancelable: true }),
    );
    expect(scroll).toHaveBeenCalledTimes(1);
    spanWith(element, 'plain').dispatchEvent(new MouseEvent('click', { bubbles: true }));
    expect(scroll).toHaveBeenCalledTimes(1);
    fillFields(element, 3, 9);
    expect(fields.map((item) => item.textContent)).toEqual(['3', '9', '3']);
  });
});

describe('DOCX tables', () => {
  const cell = (text: string, properties = ''): string =>
    `<w:tc>${properties ? `<w:tcPr>${properties}</w:tcPr>` : ''}${p(text)}</w:tc>`;
  const STYLES = `<w:docDefaults><w:rPrDefault><w:rPr><w:sz w:val="22"/></w:rPr></w:rPrDefault></w:docDefaults>
    <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/><w:pPr><w:jc w:val="center"/></w:pPr><w:rPr><w:sz w:val="24"/></w:rPr></w:style>
    <w:style w:type="table" w:default="1" w:styleId="TableNormal"><w:name w:val="Normal Table"/><w:tblPr><w:tblCellMar><w:left w:w="108" w:type="dxa"/><w:right w:w="108" w:type="dxa"/></w:tblCellMar></w:tblPr></w:style>
    <w:style w:type="table" w:styleId="Fancy"><w:name w:val="Fancy"/><w:basedOn w:val="TableNormal"/>
      <w:pPr><w:spacing w:after="0"/></w:pPr><w:rPr><w:color w:val="333333"/></w:rPr>
      <w:tblPr><w:tblBorders><w:top w:val="single" w:sz="8" w:color="00AA00"/><w:bottom w:val="single" w:sz="8" w:color="00AA00"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/><w:insideV w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBorders></w:tblPr>
      <w:tcPr><w:vAlign w:val="bottom"/></w:tcPr>
      <w:tblStylePr w:type="firstRow"><w:rPr><w:b/><w:color w:val="FFFFFF"/></w:rPr><w:tcPr><w:shd w:val="clear" w:fill="0000AA"/><w:tcBorders><w:bottom w:val="single" w:sz="16" w:color="000000"/><w:insideV w:val="nil"/></w:tcBorders></w:tcPr></w:tblStylePr>
      <w:tblStylePr w:type="lastRow"><w:rPr><w:i/></w:rPr></w:tblStylePr>
      <w:tblStylePr w:type="firstCol"><w:rPr><w:caps/></w:rPr><w:tcPr><w:tcBorders><w:right w:val="single" w:sz="8" w:color="FF0000"/><w:insideH w:val="nil"/></w:tcBorders></w:tcPr></w:tblStylePr>
      <w:tblStylePr w:type="lastCol"><w:tcPr><w:shd w:val="clear" w:fill="EEEEEE"/></w:tcPr></w:tblStylePr>
      <w:tblStylePr w:type="band1Horz"><w:tcPr><w:shd w:val="clear" w:fill="DDEEFF"/></w:tcPr></w:tblStylePr>
      <w:tblStylePr w:type="band2Vert"><w:tcPr><w:shd w:val="clear" w:fill="FFEEDD"/></w:tcPr></w:tblStylePr>
      <w:tblStylePr w:type="nwCell"><w:tcPr><w:tcBorders><w:top w:val="double" w:sz="4"/></w:tcBorders></w:tcPr></w:tblStylePr>
      <w:tblStylePr w:type="seCell"><w:rPr><w:strike/></w:rPr></w:tblStylePr>
      <w:tblStylePr w:type="neCell"><w:rPr><w:strike/></w:rPr></w:tblStylePr>
      <w:tblStylePr w:type="swCell"><w:rPr><w:strike/></w:rPr></w:tblStylePr>
    </w:style>`;

  it('applies table styles with conditional formatting and resolves shared borders', async () => {
    const row = (cells: string, properties = ''): string =>
      `<w:tr>${properties ? `<w:trPr>${properties}</w:trPr>` : ''}${cells}</w:tr>`;
    const { element } = await render({
      styles: STYLES,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="Fancy"/><w:tblW w:w="5000" w:type="pct"/><w:tblLook w:firstRow="1" w:lastRow="1" w:firstColumn="1" w:lastColumn="1" w:noHBand="0" w:noVBand="0"/></w:tblPr>
        <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
        ${row(cell('h1') + cell('h2') + cell('h3') + cell('h4'), '<w:tblHeader/><w:trHeight w:val="600"/>')}
        ${row(cell('a1') + cell('a2', '<w:tcBorders><w:bottom w:val="nil"/></w:tcBorders><w:noWrap/><w:textDirection w:val="btLr"/>') + cell('a3') + cell('a4'))}
        ${row(cell('b1') + cell('b2') + cell('b3') + cell('b4'))}
        ${row(cell('t1') + cell('t2') + cell('t3') + cell('t4'))}
      </w:tbl>`,
    });
    const table = element.querySelector<HTMLTableElement>('table')!;
    expect(table.style.width).toContain('100%');
    expect(table.style.tableLayout).toBe('fixed');
    expect(table.querySelectorAll('col')).toHaveLength(4);
    // Older documents start the table at the margin minus the cell padding.
    expect(parseFloat(table.style.marginLeft)).toBeCloseTo(-7.2, 1);
    const td = (text: string): HTMLTableCellElement =>
      spanWith(element, text).closest('td') as HTMLTableCellElement;

    expect(td('h1').parentElement!.hasAttribute('data-av-th')).toBe(true);
    expect((td('h1').parentElement as HTMLElement).style.height).toBe('40px');
    expect(td('h2').style.backgroundColor.toLowerCase()).toBe('#0000aa');
    expect(spanWith(element, 'h2').style.fontWeight).toBe('bold');
    expect(spanWith(element, 'h2').style.color.toLowerCase()).toBe('#ffffff');
    expect(td('h2').style.borderBottom).toContain('2.75px solid');
    expect(td('h1').style.borderTop).toContain('double');
    // The conditional format removes the vertical lines inside the header row.
    expect(td('h2').style.borderRight).toMatch(/^none/);
    expect(td('h2').style.borderTop).toContain('#00aa00'.toUpperCase());
    expect(td('a2').style.backgroundColor.toLowerCase()).toBe('#ddeeff');
    expect(td('b2').style.backgroundColor.toLowerCase()).toBe('');
    expect(td('b3').style.backgroundColor.toLowerCase()).toBe('#ffeedd');
    expect(td('a4').style.backgroundColor.toLowerCase()).toBe('#eeeeee');
    expect(spanWith(element, 'a1').style.textTransform).toBe('uppercase');
    expect(td('a1').style.borderRight).toContain('#FF0000');
    expect(td('a2').style.borderLeft).toContain('#FF0000');
    // An explicit "no border" on one cell wins over the table grid of its neighbour.
    expect(td('a2').style.borderBottom).toMatch(/^none/);
    expect(td('b2').style.borderTop).toMatch(/^none/);
    expect(td('a2').style.whiteSpace).toBe('nowrap');
    expect(td('a2').style.writingMode).toBe('vertical-rl');
    expect(td('a2').style.verticalAlign).toBe('bottom');
    expect(spanWith(element, 't2').style.fontStyle).toBe('italic');
    expect(spanWith(element, 't4').style.textDecoration).toContain('line-through');
    expect(td('t2').style.borderBottom).toContain('solid');
    // Edges without a border are recorded for the optional gridlines: the header cells have no
    // lines between them and no line at the sides of the table.
    expect(td('h1').getAttribute('data-av-g')).toBe('l');
    expect(td('h2').getAttribute('data-av-g')).toBe('l');
    expect(td('h4').getAttribute('data-av-g')).toBe('lr');
    // A shared edge is recorded once, on the cell below it.
    expect(td('a2').getAttribute('data-av-g')).toBeNull();
    expect(td('b2').getAttribute('data-av-g')).toBe('t');
    expect(td('a1').style.padding).toMatch(/^0px 7.2px/);
    // In tables the table style wins over the size and alignment of the default paragraph style.
    expect(spanWith(element, 'b2').style.fontSize).toBe('11pt');
    expect(paragraphWith(element, 'b2').style.textAlign).toBe('');
    expect(spanWith(element, 'b2').style.color).toBe('#333333');
  });

  it('positions floating, indented, nested and irregular tables', async () => {
    const { element } = await render({
      styles: STYLES,
      settings:
        '<w:evenAndOddHeaders/><w:defaultTabStop w:val="360"/><w:compat><w:compatSetting w:name="compatibilityMode" w:val="15"/><w:compatSetting w:name="overrideTableStyleFontSizeAndJustification" w:val="1"/></w:compat>',
      body: `<w:tbl><w:tblPr><w:tblpPr w:rightFromText="180" w:bottomFromText="90"/><w:tblW w:w="3000" w:type="dxa"/><w:tblLook w:val="04A0"/></w:tblPr><w:tblGrid><w:gridCol w:w="1500"/><w:gridCol w:w="1500"/></w:tblGrid>
          <w:tr><w:trPr><w:gridBefore w:val="1"/></w:trPr>${cell('float')}</w:tr>
          <w:tr>${cell('wide', '<w:gridSpan w:val="2"/><w:vMerge w:val="restart"/>')}<w:trPr><w:gridAfter w:val="1"/></w:trPr></w:tr>
          <w:tr>${cell('', '<w:gridSpan w:val="2"/><w:vMerge/>')}</w:tr></w:tbl>
        <w:tbl><w:tblPr><w:tblpPr w:tblpXSpec="right" w:leftFromText="150"/><w:jc w:val="right"/></w:tblPr><w:tr>${cell('right')}</w:tr></w:tbl>
        <w:tbl><w:tblPr><w:tblpPr w:tblpXSpec="center"/><w:tblInd w:w="300" w:type="dxa"/></w:tblPr><w:tr><w:tc><w:tbl><w:tblPr/><w:tr>${cell('inner')}</w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>`,
    });
    const [floating, right, centred] = Array.from(
      element.querySelectorAll<HTMLTableElement>('.av-doc-page > .av-doc-body > table'),
    );
    expect(floating!.style.float).toBe('left');
    expect(floating!.style.width).toBe('200px');
    expect(floating!.style.marginRight).toBe('12px');
    expect(floating!.style.marginBottom).toBe('6px');
    // Word 2013+ documents do not shift tables into the margin.
    expect(parseFloat(floating!.style.marginLeft)).toBe(0);
    expect(floating!.querySelector<HTMLElement>('td')!.style.padding).toBe('0px');
    expect(spanWith(element, 'wide').closest('td')!.rowSpan).toBe(2);
    expect(floating!.querySelectorAll('tr')[2]!.children).toHaveLength(0);
    expect(right!.style.float).toBe('right');
    expect(right!.style.marginLeft).toBe('10px');
    expect(centred!.style.marginLeft).toBe('auto');
    expect(centred!.style.marginRight).toBe('auto');
    // The compulsory empty paragraph after a nested table takes no room.
    const closing = spanWith(element, 'inner').closest('table')!.nextElementSibling as HTMLElement;
    expect(closing.style.display).toBe('none');
    // With the override setting, the default paragraph style keeps its size in tables.
    expect(spanWith(element, 'float').style.fontSize).toBe('12pt');
  });

  it('measures percentage widths like Word and keeps the outer border under bands', async () => {
    const { element } = await render({
      styles: `${STYLES}<w:style w:type="table" w:styleId="Band"><w:name w:val="Band"/><w:basedOn w:val="TableNormal"/>
        <w:tblPr><w:tblBorders><w:bottom w:val="single" w:sz="8" w:color="00AA00"/><w:right w:val="single" w:sz="8" w:color="00AA00"/><w:insideH w:val="single" w:sz="4" w:color="CCCCCC"/></w:tblBorders></w:tblPr>
        <w:tblStylePr w:type="band1Horz"><w:tcPr><w:tcBorders><w:top w:val="nil"/><w:bottom w:val="nil"/></w:tcBorders></w:tcPr></w:tblStylePr>
        <w:tblStylePr w:type="band1Vert"><w:tcPr><w:tcBorders><w:left w:val="nil"/><w:right w:val="nil"/></w:tcBorders></w:tcPr></w:tblStylePr></w:style>`,
      body: `<w:tbl><w:tblPr><w:tblStyle w:val="Band"/><w:tblW w:w="3500" w:type="pct"/><w:tblLook w:firstRow="0" w:firstColumn="0" w:noHBand="0" w:noVBand="0"/></w:tblPr>
        <w:tr>${cell('x')}</w:tr><w:tr>${cell('y')}</w:tr>
        <w:tr><w:tc><w:tbl><w:tblPr><w:tblW w:w="4000" w:type="pct"/></w:tblPr><w:tr>${cell('inner')}</w:tr></w:tbl><w:p/></w:tc></w:tr></w:tbl>`,
    });
    const td = (text: string): HTMLTableCellElement =>
      spanWith(element, text).closest('td') as HTMLTableCellElement;
    const outer = element.querySelector<HTMLTableElement>('table')!;
    // The table hangs into the margins by its cell padding, and the percentage includes that.
    expect(outer.style.width).toContain('14.4px) * 0.7');
    // A table inside a cell is measured against the text width of the cell.
    expect(td('inner').closest('table')!.style.width).toContain('+ 0px) * 0.8');
    // Inside the table a band removes the lines; at the edge the table border stays.
    expect(td('x').style.borderBottom).toMatch(/^none/);
    expect(td('x').style.borderRight).toContain('#00AA00');
    expect(
      (td('inner').closest('table')!.parentElement as HTMLElement).style.borderBottom,
    ).toContain('#00AA00');
  });
});

describe('DOCX drawings', () => {
  const anchor = (inner: string, attributes = ''): string =>
    `<w:r><w:drawing><wp:anchor distT="0" distB="0" distL="114300" distR="114300" ${attributes}>${inner}</wp:anchor></w:drawing></w:r>`;
  const picture =
    '<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rImage"/></pic:blipFill><pic:spPr><a:xfrm rot="5400000" flipH="1"/></pic:spPr></pic:pic></a:graphicData></a:graphic>';
  const extent = '<wp:extent cx="952500" cy="476250"/>';

  it('floats, centres and absolutely positions anchored pictures', async () => {
    const { element } = await render({
      body:
        `<w:p>${anchor(`<wp:positionH relativeFrom="margin"><wp:align>right</wp:align></wp:positionH><wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV>${extent}<wp:wrapSquare wrapText="bothSides"/>${picture}`)}<w:r><w:t>right float</w:t></w:r></w:p>` +
        `<w:p>${anchor(`<wp:positionH relativeFrom="column"><wp:posOffset>95250</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>190500</wp:posOffset></wp:positionV>${extent}<wp:wrapTight/>${picture}`)}<w:r><w:t>left float</w:t></w:r></w:p>` +
        `<w:p>${anchor(`<wp:positionH relativeFrom="page"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>9525</wp:posOffset></wp:positionV>${extent}<wp:wrapTopAndBottom/>${picture}`)}<w:r><w:t>block</w:t></w:r></w:p>` +
        `<w:p>${anchor(`<wp:positionH relativeFrom="column"><wp:posOffset>190500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>${extent}<wp:wrapTopAndBottom/>${picture}`)}<w:r><w:t>block left</w:t></w:r></w:p>` +
        `<w:p>${anchor(`<wp:positionH relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>952500</wp:posOffset></wp:positionV>${extent}<wp:wrapNone/>${picture}`, 'behindDoc="1"')}<w:r><w:t>behind</w:t></w:r></w:p>` +
        `<w:p>${anchor(`<wp:positionH relativeFrom="character"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="margin"><wp:align>bottom</wp:align></wp:positionV>${extent}<wp:wrapNone/>${picture}`)}${anchor(`<wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="page"><wp:align>center</wp:align></wp:positionV>${extent}<wp:wrapNone/>${picture}`)}${anchor(`<wp:positionH relativeFrom="margin"><wp:align>center</wp:align></wp:positionH><wp:positionV relativeFrom="line"><wp:posOffset>95250</wp:posOffset></wp:positionV>${extent}<wp:wrapNone/>${picture}`)}<w:r><w:t>front</w:t></w:r></w:p>` +
        `<w:p><w:r><w:drawing><wp:inline>${extent}${picture}</wp:inline></w:drawing></w:r><w:r><w:drawing><wp:inline>${extent}<a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rMissing"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing><wp:inline>${extent}<a:graphic><a:graphicData/></a:graphic></wp:inline></w:drawing></w:r><w:r><w:drawing/></w:r></w:p>`,
    });
    const images = Array.from(element.querySelectorAll<HTMLElement>('img'));
    expect(images).toHaveLength(9);
    const [right, left, block, blockLeft, behind, bottom, middle, lineRelative, inline] = images;
    expect(right!.style.float).toBe('right');
    expect(right!.hasAttribute('data-av-hoist')).toBe(true);
    expect(right!.style.transform).toBe('rotate(90deg) scaleX(-1)');
    expect(left!.style.float).toBe('left');
    expect(left!.style.marginLeft).toBe('10px');
    expect(left!.style.marginTop).toBe('20px');
    expect(left!.style.marginRight).toBe('12px');
    expect(block!.parentElement!.style.float).toBe('left');
    expect(block!.parentElement!.style.width).toBe('100%');
    expect(block!.parentElement!.style.textAlign).toBe('center');
    expect(blockLeft!.parentElement!.style.paddingLeft).toBe('20px');
    expect(behind!.style.position).toBe('absolute');
    expect(behind!.style.zIndex).toBe('-1');
    // Page-relative offsets are converted to the text area: 100px - 96px margins.
    expect(parseFloat(behind!.style.left)).toBe(4);
    expect(parseFloat(behind!.style.top)).toBe(4);
    expect(bottom!.style.left).toBe('');
    expect(parseFloat(bottom!.style.top)).toBeCloseTo(1056 - 192 - 50, 0);
    expect(parseFloat(middle!.style.top)).toBeCloseTo((1056 - 50) / 2 - 96, 0);
    expect(parseFloat(middle!.style.left)).toBe(0);
    expect(parseFloat(lineRelative!.style.left)).toBeCloseTo((624 - 100) / 2, 0);
    expect(parseFloat(lineRelative!.style.marginTop)).toBe(10);
    expect(inline!.style.position).toBe('');
    expect(inline!.style.width).toBe('100px');
    // A picture the browser cannot show keeps its space.
    expect(element.querySelectorAll('.av-doc-box')).toHaveLength(1);
  });

  it('draws shapes, groups and text boxes from DrawingML', async () => {
    const shape = (inner: string): string =>
      `<w:r><w:drawing><wp:inline><wp:extent cx="1905000" cy="952500"/><a:graphic><a:graphicData>${inner}</a:graphicData></a:graphic></wp:inline></w:drawing></w:r>`;
    const custom =
      '<a:custGeom><a:gdLst><a:gd name="half" fmla="*/ w 1 2"/></a:gdLst><a:pathLst><a:path w="200" h="100"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="half" y="0"/></a:lnTo><a:lnTo><a:pt x="0" y="100"/></a:lnTo><a:close/></a:path></a:pathLst></a:custGeom>';
    const { element } = await render({
      body:
        `<w:p>${shape(`<wps:wsp><wps:spPr><a:xfrm flipH="1" rot="2700000"/>${custom}<a:solidFill><a:schemeClr val="accent1"><a:alpha val="50000"/></a:schemeClr></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/><a:tailEnd type="triangle"/></a:ln></wps:spPr><wps:txbx><w:txbxContent>${p('inside')}</w:txbxContent></wps:txbx><wps:bodyPr anchor="ctr" lIns="0"><a:spAutoFit/></wps:bodyPr></wps:wsp>`)}</w:p>` +
        `<w:p>${shape(`<wps:wsp><wps:spPr><a:prstGeom prst="ellipse"/><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill><a:ln><a:noFill/></a:ln></wps:spPr><wps:bodyPr anchor="b"/></wps:wsp>`)}</w:p>` +
        `<w:p>${shape(`<wps:wsp><wps:spPr><a:prstGeom prst="rect"/><a:blipFill><a:blip r:embed="rImage"/></a:blipFill></wps:spPr><wps:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></wps:style><wps:txbx><w:txbxContent>${p('styled')}</w:txbxContent></wps:txbx></wps:wsp>`)}</w:p>` +
        `<w:p>${shape(`<wps:wsp><wps:spPr><a:prstGeom prst="rect"/></wps:spPr><wps:style><a:lnRef idx="0"/><a:fillRef idx="1"><a:schemeClr val="tx1"/></a:fillRef></wps:style></wps:wsp>`)}${shape('<wps:wsp><wps:spPr><a:noFill/></wps:spPr></wps:wsp>')}</w:p>` +
        `<w:p>${shape(`<wpg:wgp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/><a:chOff x="0" y="0"/><a:chExt cx="200" cy="100"/></a:xfrm></wpg:grpSpPr>
            <pic:pic><pic:blipFill><a:blip r:embed="rImage"/></pic:blipFill><pic:spPr><a:xfrm><a:off x="100" y="50"/><a:ext cx="100" cy="50"/></a:xfrm></pic:spPr></pic:pic>
            <wpg:grpSp><wpg:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="100" cy="50"/><a:chOff x="0" y="0"/><a:chExt cx="10" cy="10"/></a:xfrm></wpg:grpSpPr><wps:wsp><wps:spPr><a:xfrm><a:off x="5" y="5"/><a:ext cx="5" cy="5"/></a:xfrm><a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></wps:spPr></wps:wsp></wpg:grpSp>
            <wps:wsp><wps:spPr><a:noFill/></wps:spPr></wps:wsp><wpg:other/></wpg:wgp>`)}${shape('<wpg:wgp><wpg:grpSpPr/></wpg:wgp>')}</w:p>` +
        `<w:p><w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><a:graphic><a:graphicData><wps:wsp><wps:spPr><a:solidFill><a:srgbClr val="FFFF00"/></a:solidFill></wps:spPr><wps:txbx><w:txbxContent>${p('choice')}</w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:inline></w:drawing></mc:Choice><mc:Fallback><w:pict><v:shape><v:textbox><w:txbxContent>${p('fallback')}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>` +
        `<w:r><mc:AlternateContent><mc:Choice Requires="x"/><mc:Fallback><w:pict><v:shape style="width:72pt;height:36pt"><v:imagedata r:id="rImage"/><v:textbox><w:txbxContent>${p('legacy box')}</w:txbxContent></v:textbox></v:shape></w:pict></mc:Fallback></mc:AlternateContent></w:r>` +
        `<w:r><w:pict><v:rect style="width:468pt;height:.05pt" o:hr="t"/></w:pict></w:r><w:r><w:pict><v:shape><v:imagedata r:id="rMissing"/></v:shape></w:pict></w:r></w:p>`,
    });
    const boxes = Array.from(element.querySelectorAll<HTMLElement>('.av-doc-box'));
    const first = boxes[0]!;
    expect(first.style.justifyContent).toBe('center');
    expect(first.style.minHeight).toBe('100px');
    expect(first.style.transform).toBe('rotate(45deg)');
    const svg = first.querySelector<SVGSVGElement>('svg')!;
    // The outline is mirrored; the text is not.
    expect(svg.style.transform).toBe('scaleX(-1)');
    const path = svg.querySelector('path:not(marker path)')!;
    expect(path.getAttribute('d')).toBe('M0 0L100 0L0 100Z');
    expect(path.getAttribute('fill')).toBe('rgba(79, 129, 189, 0.50)');
    expect(path.getAttribute('stroke')).toBe('#ff0000');
    expect(path.getAttribute('stroke-dasharray')).toBe('8 6');
    expect(first.textContent).toBe('inside');

    const gradient = boxes[1]!;
    expect(gradient.style.justifyContent).toBe('flex-end');
    expect(gradient.querySelector('linearGradient stop')).not.toBeNull();
    expect(gradient.querySelector('path')!.getAttribute('fill')).toMatch(/^url\(#av-grad-/);
    expect(gradient.querySelector('path')!.hasAttribute('stroke')).toBe(false);

    const styled = boxes[2]!;
    await Promise.resolve();
    expect(styled.querySelector('image')!.getAttribute('href')).toMatch(/^blob:test\//);
    expect(styled.querySelector('svg > path')!.getAttribute('stroke')).toBe('#4f81bd');
    expect((styled.lastElementChild as HTMLElement).style.color.toLowerCase()).toBe('#ffffff');
    expect(boxes[3]!.querySelector('path')!.getAttribute('fill')).toBe('#000000');

    const group = boxes[4]!;
    expect(group.style.position).toBe('relative');
    const member = group.querySelector<HTMLElement>('img')!;
    expect(parseFloat(member.style.left)).toBe(100);
    expect(parseFloat(member.style.top)).toBe(50);
    expect(member.style.width).toBe('100px');
    const nested = group.querySelectorAll<HTMLElement>('.av-doc-box')[1]!;
    expect(parseFloat(nested.style.left)).toBe(50);
    expect(nested.style.width).toBe('50px');

    expect(element.textContent).toContain('choice');
    expect(element.textContent).not.toContain('fallback');
    expect(element.textContent).toContain('legacy box');
    const legacy = Array.from(element.querySelectorAll<HTMLElement>('img')).find(
      (image) => image.style.width === '72pt',
    );
    expect(legacy).toBeDefined();
    const rule = boxes.find((box) => box.style.borderTop.includes('solid'));
    expect(rule?.style.display).toBe('block');
  });
});

describe('DOCX pages', () => {
  const header = (text: string): string =>
    `${XML}<w:hdr ${NS}><w:p><w:r><w:t>${text} </w:t></w:r><w:r><w:fldChar w:fldCharType="begin"/></w:r><w:r><w:instrText>PAGE</w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>0</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r><w:r><w:drawing><wp:anchor><wp:positionH relativeFrom="page"><wp:posOffset>0</wp:posOffset></wp:positionH><wp:positionV relativeFrom="page"><wp:posOffset>95250</wp:posOffset></wp:positionV><wp:extent cx="95250" cy="95250"/><wp:wrapNone/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rLogo"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r></w:p></w:hdr>`;
  const footer = `${XML}<w:ftr ${NS}><w:p><w:r><w:t>Page </w:t></w:r><w:fldSimple w:instr="PAGE"><w:r><w:t>0</w:t></w:r></w:fldSimple><w:r><w:t> of </w:t></w:r><w:fldSimple w:instr="NUMPAGES"><w:r><w:t>0</w:t></w:r></w:fldSimple></w:p></w:ftr>`;
  const note = (id: number, text: string): string =>
    `<w:footnote w:id="${id}"><w:p><w:r><w:rPr><w:vertAlign w:val="superscript"/></w:rPr><w:footnoteRef/></w:r><w:r><w:t>${text}</w:t></w:r></w:p></w:footnote>`;
  const reference = (id: number): string => `<w:r><w:footnoteReference w:id="${id}"/></w:r>`;

  const parts: Parts = {
    settings: '<w:evenAndOddHeaders/>',
    footnotes:
      '<w:footnote w:type="separator" w:id="0"><w:p><w:r><w:separator/></w:r><w:r><w:continuationSeparator/></w:r></w:p></w:footnote>' +
      note(2, 'first note') +
      note(3, 'second note'),
    endnotes:
      '<w:endnote w:id="2"><w:p><w:r><w:endnoteRef/></w:r><w:r><w:t>the endnote</w:t></w:r></w:p></w:endnote>',
    files: {
      'word/header1.xml': header('Default'),
      'word/header2.xml': header('First'),
      'word/header3.xml': header('Even'),
      'word/footer1.xml': footer,
      'word/_rels/header1.xml.rels': relationshipsXml([['rLogo', 'image', 'media/image1.png']]),
    },
    relationships: [
      ['rH1', 'header', 'header1.xml'],
      ['rH2', 'header', 'header2.xml'],
      ['rH3', 'header', 'header3.xml'],
      ['rF1', 'footer', 'footer1.xml'],
      ['rExternal', 'header', 'https://example.com/h.xml', true],
    ],
    body:
      `<w:p><w:r><w:t>page one</w:t></w:r>${reference(2)}<w:r><w:br w:type="page"/><w:t>page two</w:t></w:r>${reference(3)}<w:r><w:br w:type="page"/></w:r></w:p>` +
      `<w:p><w:r><w:t>page three</w:t></w:r><w:r><w:endnoteReference w:id="2"/></w:r>${reference(99)}</w:p>` +
      `<w:p><w:pPr><w:sectPr><w:headerReference w:type="default" r:id="rH1"/><w:headerReference w:type="first" r:id="rH2"/><w:headerReference w:type="even" r:id="rH3"/><w:headerReference w:type="default" r:id="rExternal"/><w:headerReference w:type="odd" r:id="rH1"/><w:footerReference w:type="default" r:id="rF1"/><w:titlePg/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr><w:r><w:t>end of first section</w:t></w:r></w:p>` +
      p('continuous section') +
      `<w:p><w:pPr><w:sectPr><w:type w:val="continuous"/><w:pgSz w:w="12240" w:h="15840"/></w:sectPr></w:pPr></w:p>` +
      p('landscape', '<w:pageBreakBefore/>') +
      `<w:tbl><w:tr><w:tc>${p('in table')}</w:tc></w:tr></w:tbl>` +
      `<w:sectPr><w:pgSz w:w="15840" w:h="12240" w:orient="landscape"/><w:pgMar w:top="720" w:right="720" w:bottom="720" w:left="720" w:header="360" w:footer="360"/></w:sectPr>`,
  };

  it('shows one page with every note before layout', async () => {
    const document = await render(parts);
    expect(document.pageCount()).toBe(1);
    expect(document.page(0)!.querySelector('.av-doc-header')!.textContent).toBe('First 1');
    expect(document.page(1)).toBeNull();
    const notes = document.element.querySelector('.av-doc-footnotes')!;
    expect(notes.textContent).toBe('1first note2second note');
    expect(notes.querySelectorAll('[data-av-sep] span span')).toHaveLength(2);
    expect(document.element.querySelector('.av-doc-endnotes')!.textContent).toBe('ithe endnote');
    // Nothing happens until the document is attached.
    await document.layout();
    expect(document.pageCount()).toBe(1);
  });

  it('splits pages at breaks and sections with headers, footers and footnotes', async () => {
    const document = await render(parts);
    globalThis.document.body.appendChild(document.element);
    await document.layout();
    await document.layout();
    const pages = Array.from({ length: document.pageCount() }, (_unused, index) =>
      document.page(index)!,
    );
    const text = (index: number): string =>
      pages[index]!.querySelector('.av-doc-body')!.textContent ?? '';
    expect(pages).toHaveLength(4);
    expect(text(0)).toContain('page one');
    expect(text(1)).toBe('page two2');
    expect(text(2)).toContain('page three');
    expect(text(2)).toContain('continuous section');
    // Endnotes close the document.
    expect(text(3)).toBe('landscapein tableithe endnote');

    // First page, even page and default headers; the page number is live.
    expect(pages.map((page) => page.querySelector('.av-doc-header')!.textContent)).toEqual([
      'First 1',
      'Even 2',
      'Default 3',
      'Even 4',
    ]);
    expect(pages[1]!.querySelector('.av-doc-footer')!.textContent).toBe('Page 2 of 4');
    expect(pages[2]!.querySelector('.av-doc-header img')!.getAttribute('src')).toMatch(
      /^blob:test\//,
    );
    const logo = pages[2]!.querySelector<HTMLElement>('.av-doc-header img')!;
    expect(parseFloat(logo.style.left)).toBe(-96);
    expect(parseFloat(logo.style.top)).toBe(10);

    // Each page carries the footnotes its text refers to.
    const notes = (index: number): string =>
      Array.from(pages[index]!.querySelectorAll('.av-doc-note:not([data-av-sep])'))
        .map((item) => item.textContent)
        .join('|');
    expect(notes(0)).toBe('1first note');
    expect(notes(1)).toBe('2second note');
    expect(notes(2)).toBe('');
    expect(notes(3)).toBe('ithe endnote');
    expect(pages[3]!.querySelector<HTMLElement>('.av-doc-footnotes')!.hidden).toBe(true);

    expect(pages[0]!.style.width).toBe('816px');
    expect(pages[0]!.style.height).toBe('1056px');
    expect(pages[3]!.style.width).toBe('1056px');
    expect(pages[3]!.style.padding).toMatch(/^48px/);
    expect(pages[3]!.querySelector<HTMLElement>('.av-doc-header')!.style.paddingTop).toBe('24px');
    document.destroy();
  });
});

describe('DOCX layout with measurements', () => {
  /**
   * A tiny layout engine for the tests: blocks are stacked in their parent,
   * a block is as tall as the number in its `H…` text (default 100), table
   * rows are 100 tall and text wraps every 10 characters into 20px lines.
   */
  function fakeLayout(pageHeight: number): void {
    const heightOf = (element: Element): number => {
      if (element.tagName === 'TABLE') return element.querySelectorAll('tr').length * 100;
      if (element.tagName === 'TR') return 100;
      const marked = /H(\d+)/.exec(element.textContent ?? '');
      if (marked) return Number(marked[1]);
      return Math.max(1, Math.ceil((element.textContent ?? '').length / 10)) * 20;
    };
    const topOf = (element: Element): number => {
      if (element.classList.contains('av-doc-body')) return 0;
      const parent = element.parentElement;
      if (!parent) return 0;
      let top = topOf(parent.tagName === 'TBODY' ? parent.parentElement! : parent);
      for (
        let sibling = element.previousElementSibling;
        sibling;
        sibling = sibling.previousElementSibling
      ) {
        top += heightOf(sibling);
      }
      return top;
    };
    const rect = (top: number, height: number, left = 0, width = 100): DOMRect =>
      ({
        top,
        bottom: top + height,
        left,
        right: left + width,
        width,
        height,
        x: left,
        y: top,
      }) as DOMRect;
    vi.spyOn(HTMLElement.prototype, 'getBoundingClientRect').mockImplementation(function (
      this: HTMLElement,
    ) {
      if (this.classList.contains('av-doc-body')) return rect(0, pageHeight, 0, 600);
      if (this.classList.contains('av-doc-tab')) {
        return rect(
          0,
          20,
          Number(this.getAttribute('data-x') ?? 0),
          parseFloat(this.style.width) || 0,
        );
      }
      if (
        this.classList.contains('av-doc-marker') ||
        this.parentElement?.classList.contains('av-doc-marker')
      ) {
        return rect(0, 20, 0, 60);
      }
      if (this.classList.contains('av-doc')) return rect(0, 0, 0, 800);
      return rect(topOf(this), heightOf(this), Number(this.getAttribute('data-left') ?? 0), 500);
    });
    vi.spyOn(HTMLElement.prototype, 'clientHeight', 'get').mockImplementation(function (
      this: HTMLElement,
    ) {
      return this.classList.contains('av-doc-body') ? pageHeight : 0;
    });
    vi.spyOn(HTMLElement.prototype, 'clientWidth', 'get').mockReturnValue(600);
    vi.spyOn(HTMLElement.prototype, 'offsetWidth', 'get').mockImplementation(function (
      this: HTMLElement,
    ) {
      return this.classList.contains('av-doc') ? 800 : 0;
    });
    vi.spyOn(HTMLElement.prototype, 'offsetHeight', 'get').mockImplementation(function (
      this: HTMLElement,
    ) {
      return this.classList.contains('av-doc-header') ||
        this.parentElement?.classList.contains('av-doc-footer')
        ? 150
        : 0;
    });
    const lineRects = (node: Node, from: number, to: number): DOMRect[] => {
      const paragraph = node.parentElement!.closest('p,h1')!;
      const top = topOf(paragraph);
      const rects: DOMRect[] = [];
      for (
        let line = Math.floor(from / 10);
        line <= Math.floor(Math.max(from, to - 1) / 10);
        line++
      ) {
        rects.push(rect(top + line * 20, 20));
      }
      return rects;
    };
    vi.spyOn(Range.prototype, 'getClientRects').mockImplementation(function (this: Range) {
      const node = this.startContainer;
      const rects = node.nodeType === 3 ? lineRects(node, this.startOffset, this.endOffset) : [];
      return rects as unknown as DOMRectList;
    });
    vi.spyOn(Range.prototype, 'getBoundingClientRect').mockImplementation(function (this: Range) {
      const node = this.startContainer;
      if (node.nodeType === 3) return lineRects(node, this.startOffset, this.startOffset + 1)[0]!;
      return rect(0, 20, 0, 40);
    });
  }

  async function paginate(
    parts: Parts,
    pageHeight: number,
  ): Promise<{ document: DocxDocument; pages: string[] }> {
    const document = await render(parts);
    globalThis.document.body.appendChild(document.element);
    fakeLayout(pageHeight);
    await document.layout();
    const pages = Array.from({ length: document.pageCount() }, (_unused, index) =>
      (document.page(index)!.querySelector('.av-doc-body')!.textContent ?? '').replace(/\s+/g, ' '),
    );
    return { document, pages };
  }

  it('moves blocks that do not fit, keeping headings with their text', async () => {
    const { document, pages } = await paginate(
      {
        body:
          p('A H300', '<w:spacing w:before="200"/>') +
          p('Heading H100', '<w:keepNext/>') +
          p('B H400', '<w:keepLines/><w:spacing w:before="200"/>') +
          p('C H900', '<w:keepLines/>') +
          p('D H100'),
      },
      700,
    );
    expect(pages).toEqual(['A H300', 'Heading H100B H400', 'C H900', 'D H100']);
    // Space before a paragraph is dropped at the top of an automatic page.
    const moved = paragraphWith(document.page(1)!, 'Heading');
    expect(moved.style.marginTop).toBe('0px');
    expect(paragraphWith(document.page(0)!, 'A H300').style.marginTop).not.toBe('0px');
  });

  it('splits tables between rows and repeats header rows', async () => {
    const row = (text: string, header = false): string =>
      `<w:tr>${header ? '<w:trPr><w:tblHeader/></w:trPr>' : ''}<w:tc>${p(text)}</w:tc></w:tr>`;
    const { document, pages } = await paginate(
      {
        body: `<w:tbl><w:tblGrid><w:gridCol w:w="3000"/></w:tblGrid>${row('head', true)}${row('r1')}${row('r2')}${row('r3')}${row('r4')}</w:tbl>${p('tail H100')}`,
      },
      350,
    );
    expect(pages).toEqual(['headr1r2', 'headr3r4', 'tail H100']);
    expect(document.page(1)!.querySelector('colgroup')).not.toBeNull();
    expect(document.page(1)!.querySelector<HTMLElement>('table')!.style.marginTop).toBe('0px');
  });

  it('splits paragraphs between lines with widow and orphan control', async () => {
    const long = 'abcdefghij'.repeat(8);
    const { document, pages } = await paginate(
      {
        body:
          p('H60 first') +
          p(long, '<w:jc w:val="both"/>') +
          p('0123456789'.repeat(3), '<w:widowControl w:val="0"/>') +
          p('x'.repeat(25)),
      },
      180,
    );
    // 60px block, then six 20px lines fit before the break; the rest continues.
    expect(pages[0]).toBe(`H60 first${'abcdefghij'.repeat(6)}`);
    expect(pages[1]!.startsWith('abcdefghij'.repeat(2))).toBe(true);
    const rest = document.page(1)!.querySelector<HTMLElement>('p')!;
    expect(rest.style.textIndent).toBe('0px');
    expect(rest.style.marginTop).toBe('0px');
    const head = document.page(0)!.querySelectorAll<HTMLElement>('p')[1]!;
    expect(head.style.marginBottom).toBe('0px');
    expect(pages.join('')).toContain('x'.repeat(25));
    expect(pages.join('').replace(/\D/g, '').length).toBeGreaterThanOrEqual(30);
  });

  it('starts a page with pictures aligned to its top and grows margins for tall headers', async () => {
    const hoisted = `<w:r><w:drawing><wp:anchor><wp:positionH relativeFrom="margin"><wp:align>left</wp:align></wp:positionH><wp:positionV relativeFrom="margin"><wp:align>top</wp:align></wp:positionV><wp:extent cx="952500" cy="476250"/><wp:wrapSquare/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="rImage"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:anchor></w:drawing></w:r>`;
    const { document } = await paginate(
      {
        files: {
          'word/header1.xml': `${XML}<w:hdr ${NS}>${p('tall header')}</w:hdr>`,
          'word/footer1.xml': `${XML}<w:ftr ${NS}>${p('tall footer')}</w:ftr>`,
        },
        relationships: [
          ['rH1', 'header', 'header1.xml'],
          ['rF1', 'footer', 'footer1.xml'],
        ],
        body:
          p('Title H50') +
          `<w:p>${hoisted}<w:r><w:t>text H50</w:t></w:r></w:p>` +
          '<w:sectPr><w:headerReference w:type="default" r:id="rH1"/><w:footerReference w:type="default" r:id="rF1"/></w:sectPr>',
      },
      700,
    );
    const page = document.page(0)!;
    expect(page.querySelector('.av-doc-body')!.firstElementChild!.tagName).toBe('IMG');
    expect(page.style.paddingTop).toBe('150px');
    expect(page.style.paddingBottom).toBe('198px');
  });

  it('positions tab stops, leaders and list numbers', async () => {
    const { element } = await render({
      numbering:
        '<w:abstractNum w:abstractNumId="0"><w:lvl w:ilvl="0"><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%1.%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl></w:abstractNum><w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>',
      body:
        `<w:p><w:pPr><w:tabs><w:tab w:val="left" w:pos="1500"/><w:tab w:val="center" w:leader="hyphen" w:pos="4500"/><w:tab w:val="right" w:leader="dot" w:pos="9000"/></w:tabs></w:pPr><w:r><w:t>a</w:t><w:tab/><w:t>b</w:t><w:tab/><w:t>c</w:t><w:tab/><w:t>d</w:t><w:tab/><w:t>e</w:t><w:ptab w:alignment="center" w:leader="underscore"/><w:ptab w:alignment="right"/><w:ptab w:alignment="left"/></w:r></w:p>` +
        p('numbered', '<w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr>') +
        `<w:p><w:pPr><w:ind w:left="720" w:hanging="720"/></w:pPr><w:r><w:t>x</w:t><w:tab/><w:t>hanging</w:t></w:r></w:p>`,
    });
    globalThis.document.body.appendChild(element);
    fakeLayout(700);
    const tabs = Array.from(element.querySelectorAll<HTMLElement>('.av-doc-tab'));
    [20, 150, 320, 610, 0, 0, 0, 10].forEach((x, index) =>
      tabs[index]!.setAttribute('data-x', String(x)),
    );
    layoutTabs(element);
    // Left stop at 100px, centre stop at 300px (minus half the 40px text), right stop at 600px.
    expect(tabs[0]!.style.width).toBe('80px');
    expect(tabs[1]!.style.width).toBe('130px');
    expect(tabs[1]!.getAttribute('data-av-leader-style')).toBe('hyphen');
    expect(tabs[2]!.style.width).toBe('240px');
    expect(tabs[2]!.getAttribute('data-av-leader-style')).toBe('dot');
    // Past the last stop, default stops every half inch apply.
    expect(tabs[3]!.style.width).toBe('14px');
    expect(tabs[4]!.style.width).toBe('280px');
    expect(tabs[4]!.getAttribute('data-av-leader-style')).toBe('line');
    expect(tabs[5]!.style.width).toBe('560px');
    expect(tabs[6]!.style.width).toBe('0px');
    // A hanging indent is a tab stop of its own.
    expect(tabs[7]!.style.width).toBe('38px');
    // A number wider than its hanging indent moves the text to the next stop.
    const marker = element.querySelector<HTMLElement>('.av-doc-marker')!;
    expect(marker.style.minWidth).toBe('72px');
    layoutTabs(element.querySelector<HTMLElement>('[data-av-tabs]')!);
    layoutTabs(globalThis.document.createElement('div'));
  });

  it('aligns numbers in table cells on a decimal tab stop', async () => {
    const decimal = (text: string, pos: number): string =>
      `<w:p><w:pPr><w:tabs><w:tab w:val="decimal" w:pos="${pos}"/></w:tabs></w:pPr>${text ? `<w:r><w:t>${text}</w:t></w:r>` : ''}</w:p>`;
    const { element } = await render({
      body: `<w:tbl><w:tr><w:tc>${decimal('12.5', 3000)}${decimal('7', 3000)}${decimal('9', 360)}${decimal('', 3000)}</w:tc></w:tr></w:tbl>${decimal('8', 3000)}`,
    });
    globalThis.document.body.appendChild(element);
    fakeLayout(700);
    layoutTabs(element);
    // The stop is at 200px and the fake text before the decimal point is 100px wide.
    expect(paragraphWith(element, '12.5').style.textIndent).toBe('100px');
    expect(paragraphWith(element, '7').style.textIndent).toBe('100px');
    // Text wider than the distance to the stop stays where it is.
    expect(paragraphWith(element, '9').style.textIndent).toBe('0px');
    // Outside a table a decimal stop needs a tab character.
    expect(paragraphWith(element, '8').style.textIndent).toBe('');
  });
});

describe('embedded fonts', () => {
  it('reverses the font obfuscation', () => {
    const key = '{00112233-4455-6677-8899-AABBCCDDEEFF}';
    const data = Uint8Array.from({ length: 40 }, (_unused, index) => index);
    const result = deobfuscateFont(data, key);
    expect(result[0]).toBe(0 ^ 0xff);
    expect(result[15]).toBe(15 ^ 0x00);
    expect(result[16]).toBe(16 ^ 0xff);
    expect(result[32]).toBe(32);
    expect(Array.from(deobfuscateFont(result, key))).toEqual(Array.from(data));
    expect(deobfuscateFont(data, 'not-a-guid')).toBe(data);
  });

  it('registers embedded fonts under private names and releases them', async () => {
    const added: unknown[] = [];
    const deleted: unknown[] = [];
    class FakeFace {
      constructor(
        readonly family: string,
        readonly source: ArrayBuffer,
        readonly descriptors: FontFaceDescriptors,
      ) {
        if (descriptors.style === 'italic') throw new Error('rejected');
      }
      load(): Promise<this> {
        return Promise.resolve(this);
      }
    }
    vi.stubGlobal('FontFace', FakeFace);
    Object.defineProperty(document, 'fonts', {
      configurable: true,
      value: {
        add: (face: unknown) => added.push(face),
        delete: (face: unknown) => deleted.push(face),
        ready: Promise.resolve(),
      },
    });
    try {
      const rendered = await render({
        fontTable:
          '<w:font w:name="Custom Sans"><w:family w:val="swiss"/><w:embedRegular r:id="rFont" w:fontKey="{00112233-4455-6677-8899-AABBCCDDEEFF}"/><w:embedBold r:id="rFont"/><w:embedItalic r:id="rFont"/><w:embedBoldItalic r:id="rGone"/></w:font>' +
          '<w:font w:name="Mono"><w:pitch w:val="fixed"/></w:font><w:font w:name="Old"><w:family w:val="roman"/></w:font><w:font w:name="Hand"><w:family w:val="script"/></w:font><w:font/>',
        files: {
          'word/fonts/font1.odttf': new Uint8Array(64),
          'word/_rels/fontTable.xml.rels': relationshipsXml([
            ['rFont', 'font', 'fonts/font1.odttf'],
            ['rGone', 'font', 'fonts/missing.odttf'],
          ]),
        },
        body:
          p('custom', '', '<w:rFonts w:ascii="Custom Sans"/>') +
          p('mono', '', '<w:rFonts w:ascii="Mono"/>') +
          p('old', '', '<w:rFonts w:ascii="Old"/>') +
          p('hand', '', '<w:rFonts w:ascii="Hand"/>') +
          p('guess', '', '<w:rFonts w:ascii="Courier Prime"/>') +
          p('serif', '', '<w:rFonts w:ascii="Georgia"/>'),
      });
      expect(added).toHaveLength(2);
      expect((added[0] as FakeFace).family).toMatch(/^av\d+ Custom Sans$/);
      const family = (text: string): string => spanWith(rendered.element, text).style.fontFamily;
      expect(family('custom')).toMatch(/^"av\d+ Custom Sans", "Custom Sans", sans-serif$/);
      expect(family('mono')).toContain('monospace');
      expect(family('old')).toContain('serif');
      expect(family('hand')).toContain('cursive');
      expect(family('guess')).toContain('monospace');
      expect(family('serif')).toContain('serif');
      document.body.appendChild(rendered.element);
      await rendered.layout();
      rendered.destroy();
      expect(deleted).toHaveLength(2);
    } finally {
      vi.unstubAllGlobals();
      Reflect.deleteProperty(document, 'fonts');
    }
  });
});
