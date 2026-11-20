import { afterEach, describe, expect, it, vi } from 'vitest';
import { parseChart } from '../../src/formats/chart';
import {
  applyColorModifiers,
  buildShapeSvg,
  isLineGeometry,
  parseGradient,
  resolveColor,
  shapeGeometry,
} from '../../src/formats/drawingml';
import { openPresentation } from '../../src/formats/pptx';
import { EMPTY_THEME } from '../../src/formats/theme';
import { renderChart } from '../../src/ui/chart';
import { createPaperView } from '../../src/ui/paper';
import { parseXml, type XmlElement } from '../../src/util/xml';
import { openZip } from '../../src/util/zip';
import { renderContext } from '../helpers/dom';
import { makeZip, PNG_1X1 } from '../helpers/zip';

const XML = '<?xml version="1.0" encoding="UTF-8"?>';
const A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"';
const node = (xml: string): XmlElement =>
  parseXml(`<a:root ${A}>${xml}</a:root>`).children[0] as XmlElement;

afterEach(() => {
  vi.restoreAllMocks();
  document.body.textContent = '';
});

describe('DrawingML colours', () => {
  const scheme = (name: string): string | undefined => (name === 'accent1' ? '#4472c4' : undefined);

  it('resolves every colour model', () => {
    const color = (xml: string): string | undefined =>
      resolveColor(node(`<a:solidFill>${xml}</a:solidFill>`), scheme);
    expect(color('<a:srgbClr val="FF0000"/>')).toBe('#ff0000');
    expect(color('<a:sysClr val="window" lastClr="FFFFFF"/>')).toBe('#ffffff');
    expect(color('<a:prstClr val="navy"/>')).toBe('#000080');
    expect(color('<a:schemeClr val="accent1"/>')).toBe('#4472c4');
    expect(color('<a:schemeClr val="missing"/>')).toBeUndefined();
    expect(color('<a:scrgbClr r="100000" g="50000" b="0"/>')).toBe('#ff8000');
    expect(color('<a:other/>')).toBeUndefined();
    expect(resolveColor(null, scheme)).toBeUndefined();
  });

  it('applies modifiers', () => {
    const modifiers = (xml: string): XmlElement[] =>
      node(`<a:c>${xml}</a:c>`).children as XmlElement[];
    expect(applyColorModifiers('#808080', modifiers('<a:lumMod val="50000"/>'))).toBe('#404040');
    expect(applyColorModifiers('#000000', modifiers('<a:lumOff val="50000"/>'))).toBe('#808080');
    expect(applyColorModifiers('#000000', modifiers('<a:tint val="50000"/>'))).toBe('#808080');
    expect(applyColorModifiers('#ffffff', modifiers('<a:shade val="50000"/>'))).toBe('#808080');
    expect(
      applyColorModifiers('#ff0000', modifiers('<a:alpha val="25000"/><a:satMod val="1"/>')),
    ).toBe('rgba(255, 0, 0, 0.25)');
  });

  it('parses gradients', () => {
    const gradient = parseGradient(
      node(
        '<a:gradFill><a:gsLst><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="50000"><a:schemeClr val="none"/></a:gs></a:gsLst><a:lin ang="2700000"/></a:gradFill>',
      ),
      scheme,
    );
    expect(gradient).toEqual({
      angle: 45,
      stops: [
        { position: 0, color: '#ff0000' },
        { position: 100, color: '#0000ff' },
      ],
    });
    expect(parseGradient(node('<a:gradFill/>'), scheme)).toBeNull();
    expect(
      parseGradient(
        node(
          '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill>',
        ),
        scheme,
      )?.angle,
    ).toBe(90);
  });
});

describe('shape geometry', () => {
  const PRESETS = [
    'rect',
    'line',
    'straightConnector1',
    'bentConnector2',
    'bentConnector3',
    'curvedConnector2',
    'curvedConnector3',
    'ellipse',
    'arc',
    'donut',
    'roundRect',
    'flowChartTerminator',
    'round1Rect',
    'round2SameRect',
    'round2DiagRect',
    'snip1Rect',
    'snip2SameRect',
    'snip2DiagRect',
    'triangle',
    'flowChartMerge',
    'rtTriangle',
    'diamond',
    'parallelogram',
    'trapezoid',
    'flowChartManualOperation',
    'pentagon',
    'hexagon',
    'heptagon',
    'octagon',
    'decagon',
    'dodecagon',
    'plus',
    'rightArrow',
    'leftArrow',
    'upArrow',
    'downArrow',
    'leftRightArrow',
    'upDownArrow',
    'chevron',
    'homePlate',
    'star4',
    'star5',
    'star6',
    'star7',
    'star8',
    'star10',
    'star12',
    'star16',
    'frame',
    'heart',
    'can',
    'bracketPair',
    'leftBracket',
    'rightBracket',
    'bracePair',
    'somethingUnknown',
  ];

  it.each(PRESETS)('draws the %s preset inside its box', (preset) => {
    const paths = shapeGeometry(node(`<a:spPr><a:prstGeom prst="${preset}"/></a:spPr>`), 200, 100);
    expect(paths.length).toBeGreaterThan(0);
    for (const path of paths) {
      expect(path.d).toMatch(/^M/);
      expect(path.d).not.toMatch(/NaN|Infinity/);
      expect(path.filled || path.stroked).toBe(true);
    }
    expect(paths[0]!.filled).toBe(!isLineGeometry(preset));
  });

  it('honours adjustment values and defaults to a rectangle', () => {
    const rounded = (adjust: string): string =>
      shapeGeometry(
        node(
          `<a:spPr><a:prstGeom prst="roundRect"><a:avLst>${adjust}</a:avLst></a:prstGeom></a:spPr>`,
        ),
        200,
        100,
      )[0]!.d;
    expect(rounded('<a:gd name="adj" fmla="val 50000"/>')).toContain('A50 50');
    expect(rounded('<a:gd name="adj" fmla="val 10000"/>')).toContain('A10 10');
    expect(rounded('<a:gd fmla="bad"/>')).toContain('A16.67 16.67');
    expect(shapeGeometry(null, 20, 10)[0]!.d).toBe('M0 0L20 0L20 10L0 10Z');
    // Regular polygons and stars are stretched to fill the box, like Office does.
    const star = shapeGeometry(node('<a:spPr><a:prstGeom prst="star5"/></a:spPr>'), 200, 100)[0]!.d;
    expect(star).toContain('L200 ');
    expect(star).toContain(' 100');
  });

  it('follows custom geometry, including guides, curves and arcs', () => {
    const paths = shapeGeometry(
      node(`<a:spPr><a:custGeom>
        <a:avLst><a:gd name="adj" fmla="val 25"/></a:avLst>
        <a:gdLst>
          <a:gd name="x1" fmla="*/ w adj 100"/><a:gd name="x2" fmla="+- r 0 x1"/><a:gd name="mid" fmla="+/ l r 2"/>
          <a:gd name="pick" fmla="?: adj x1 x2"/><a:gd name="ab" fmla="abs -5"/><a:gd name="mn" fmla="min 1 2"/>
          <a:gd name="mx" fmla="max 1 2"/><a:gd name="pn" fmla="pin 0 5 3"/><a:gd name="sq" fmla="sqrt 16"/>
          <a:gd name="md" fmla="mod 3 4 0"/><a:gd name="sn" fmla="sin 10 5400000"/><a:gd name="cs" fmla="cos 10 0"/>
          <a:gd name="tn" fmla="tan 10 2700000"/><a:gd name="at" fmla="at2 1 1"/><a:gd name="ca" fmla="cat2 10 1 0"/>
          <a:gd name="sa" fmla="sat2 10 0 1"/><a:gd name="zero" fmla="*/ 1 1 0"/><a:gd name="avg0" fmla="+/ 1 1 0"/><a:gd fmla="val 1"/><a:gd name="odd" fmla="nope 1"/>
        </a:gdLst>
        <a:pathLst>
          <a:path w="100" h="50">
            <a:moveTo><a:pt x="x1" y="0"/></a:moveTo><a:lnTo><a:pt x="x2" y="t"/></a:lnTo>
            <a:cubicBezTo><a:pt x="100" y="0"/><a:pt x="100" y="50"/><a:pt x="mid" y="b"/></a:cubicBezTo>
            <a:quadBezTo><a:pt x="0" y="50"/><a:pt x="0" y="hd2"/></a:quadBezTo>
            <a:arcTo wR="10" hR="10" stAng="cd2" swAng="5400000"/>
            <a:arcTo wR="5" hR="5" stAng="0" swAng="-21600000"/>
            <a:lnTo/><a:moveTo/><a:cubicBezTo><a:pt x="1" y="1"/></a:cubicBezTo><a:quadBezTo/><a:close/>
          </a:path>
          <a:path fill="none" stroke="0"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="unknown" y="50"/></a:lnTo></a:path>
          <a:path/>
        </a:pathLst></a:custGeom></a:spPr>`),
      200,
      100,
    );
    expect(paths).toHaveLength(2);
    // The path space (100 x 50) is scaled to the 200 x 100 box.
    expect(
      paths[0]!.d.startsWith('M50 0L150 0C200 0 200 100 100 100Q0 100 0 50A20 20 0 0 1 20 30'),
    ).toBe(true);
    expect(paths[0]!.d.match(/A10 10/g)).toHaveLength(2);
    expect(paths[0]!.d.endsWith('Z')).toBe(true);
    expect(paths[1]).toEqual({ d: 'M0 0L0 50', filled: false, stroked: false });
    // An empty custom geometry falls back to the preset or a rectangle.
    expect(
      shapeGeometry(node('<a:spPr><a:custGeom><a:pathLst/></a:custGeom></a:spPr>'), 10, 10)[0]!.d,
    ).toContain('L10 10');
  });
});

describe('shape SVG', () => {
  const box = [{ d: 'M0 0L10 0L10 10Z', filled: true, stroked: true }];

  it('paints solid fills, dashes and arrow heads', () => {
    const svg = buildShapeSvg(
      [
        ...box,
        { d: 'M0 0L10 10', filled: false, stroked: true },
        { d: 'M0 5L10 5', filled: false, stroked: false },
      ],
      10,
      10,
      {
        fill: '#ff0000',
        stroke: '#000000',
        strokeWidth: 2,
        dash: 'sysDash',
        headEnd: true,
        tailEnd: true,
      },
    );
    const paths = Array.from(svg.querySelectorAll(':scope > path'));
    expect(svg.getAttribute('viewBox')).toBe('0 0 10 10');
    expect(paths[0]!.getAttribute('fill')).toBe('#ff0000');
    expect(paths[0]!.getAttribute('stroke-dasharray')).toBe('6 2');
    expect(paths[0]!.hasAttribute('marker-end')).toBe(false);
    expect(paths[1]!.getAttribute('fill')).toBe('none');
    expect(paths[1]!.getAttribute('marker-end')).toMatch(/^url\(#av-arrow-/);
    expect(paths[1]!.getAttribute('marker-start')).toMatch(/^url\(#av-arrow-/);
    expect(paths[2]!.hasAttribute('stroke')).toBe(false);
    expect(svg.querySelector('marker path')!.getAttribute('fill')).toBe('#000000');
  });

  it('paints gradients, pictures and nothing at all', async () => {
    const gradient = buildShapeSvg(box, 10, 10, {
      fill: {
        angle: 0,
        stops: [
          { position: 0, color: '#fff' },
          { position: 100, color: '#000' },
        ],
      },
      stroke: null,
      strokeWidth: 1,
    });
    expect(gradient.querySelector('linearGradient')!.getAttribute('x2')).toBe('1');
    expect(gradient.querySelectorAll('stop')).toHaveLength(2);
    expect(gradient.querySelector(':scope > path')!.hasAttribute('stroke')).toBe(false);

    const picture = buildShapeSvg(
      [...box, { d: 'M0 0L1 1', filled: false, stroked: true }],
      10,
      10,
      {
        fill: { image: Promise.resolve('blob:picture') },
        stroke: '#000',
        strokeWidth: 1,
        dash: 'unknown',
      },
    );
    await Promise.resolve();
    expect(picture.querySelector('image')!.getAttribute('href')).toBe('blob:picture');
    expect(picture.querySelectorAll('clipPath path')).toHaveLength(1);
    const missing = buildShapeSvg(box, 10, 10, {
      fill: { image: Promise.resolve(null) },
      stroke: null,
      strokeWidth: 1,
    });
    await Promise.resolve();
    expect(missing.querySelector('image')!.hasAttribute('href')).toBe(false);

    // A flat line still gets a drawable area.
    const flat = buildShapeSvg([{ d: 'M0 0L50 0', filled: false, stroked: true }], 50, 0, {
      fill: null,
      stroke: '#000',
      strokeWidth: 1,
    });
    expect(flat.getAttribute('viewBox')).toBe('0 0 50 1');
    expect(flat.style.height).toBe('1px');
    expect(flat.querySelector('path')!.getAttribute('fill')).toBe('none');
  });
});

const P =
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"';
const REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const rels = (entries: [string, string, string][]): string =>
  `${XML}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">${entries
    .map(
      ([id, type, target]) => `<Relationship Id="${id}" Type="${REL}/${type}" Target="${target}"/>`,
    )
    .join('')}</Relationships>`;

const PIE = `${XML}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" ${A}><c:chart><c:plotArea><c:pieChart><c:varyColors val="1"/>
  <c:ser><c:idx val="0"/><c:order val="0"/><c:tx><c:strRef><c:strCache><c:pt idx="0"><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>
  <c:cat><c:strRef><c:strCache><c:ptCount val="3"/><c:pt idx="0"><c:v>A</c:v></c:pt><c:pt idx="1"><c:v>B</c:v></c:pt><c:pt idx="2"><c:v>C</c:v></c:pt></c:strCache></c:strRef></c:cat>
  <c:val><c:numRef><c:numCache><c:ptCount val="3"/><c:pt idx="0"><c:v>3</c:v></c:pt><c:pt idx="1"><c:v>1</c:v></c:pt><c:pt idx="2"><c:v>0</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>
  <c:dLbls><c:showPercent val="1"/></c:dLbls></c:pieChart></c:plotArea></c:chart></c:chartSpace>`;

describe('chart slice labels', () => {
  it('labels pie slices with their share or value', () => {
    const model = parseChart(parseXml(PIE), EMPTY_THEME)!;
    expect(model.groups[0]!.showPercent).toBe(true);
    const labels = (): string[] =>
      Array.from(renderChart(model, 300, 200).querySelectorAll('text'))
        .map((text) => text.textContent ?? '')
        .filter((text) => /^\d+%?$/.test(text));
    expect(labels()).toEqual(['75%', '25%']);
    model.groups[0]!.showPercent = false;
    model.groups[0]!.showValues = true;
    expect(labels()).toEqual(['3', '1']);
    model.groups[0]!.kind = 'doughnut';
    expect(labels()).toEqual(['3', '1']);
  });
});

describe('PPTX shapes, charts, pictures and tables', () => {
  const shape = (properties: string, extra = ''): string =>
    `<p:sp><p:nvSpPr><p:cNvPr id="1" name="s"/><p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="1905000" cy="952500"/></a:xfrm>${properties}</p:spPr>${extra}</p:sp>`;
  const slide = `${XML}<p:sld ${P} ${A}><p:cSld><p:spTree>
    ${shape('<a:prstGeom prst="triangle"/><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill><a:ln w="19050"><a:solidFill><a:srgbClr val="0000FF"/></a:solidFill><a:tailEnd type="arrow"/></a:ln>', '<p:txBody><a:bodyPr/><a:p><a:r><a:t>inside</a:t></a:r></a:p></p:txBody>')}
    ${shape('<a:custGeom><a:pathLst><a:path w="2" h="1"><a:moveTo><a:pt x="0" y="0"/></a:moveTo><a:lnTo><a:pt x="2" y="1"/></a:lnTo></a:path></a:pathLst></a:custGeom><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst></a:gradFill><a:ln><a:noFill/></a:ln>')}
    ${shape('<a:prstGeom prst="star5"/><a:blipFill><a:blip r:embed="rImage"/></a:blipFill>', '<p:style><a:lnRef idx="1"><a:srgbClr val="112233"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="445566"/></a:fillRef></p:style>')}
    ${shape('<a:prstGeom prst="star5"/><a:blipFill><a:blip r:embed="rGone"/></a:blipFill>')}
    ${shape('<a:prstGeom prst="diamond"/>', '<p:style><a:lnRef idx="1"><a:srgbClr val="112233"/></a:lnRef><a:fillRef idx="1"><a:srgbClr val="445566"/></a:fillRef></p:style>')}
    ${shape('<a:prstGeom prst="diamond"/><a:noFill/>')}
    ${shape('<a:prstGeom prst="line"/><a:ln w="12700"><a:prstDash val="dash"/></a:ln>')}
    ${shape('<a:prstGeom prst="bentConnector3"/><a:ln w="12700"><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill></a:ln>')}
    ${shape('<a:prstGeom prst="rect"/><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln><a:prstDash val="sysDot"/><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>')}
    ${shape('<a:prstGeom prst="roundRect"><a:avLst><a:gd name="adj" fmla="val 50000"/></a:avLst></a:prstGeom><a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill><a:ln><a:prstDash val="solid"/></a:ln>')}
    <p:pic><p:nvPicPr><p:cNvPr id="9" name="pic"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rImage"/><a:srcRect l="25000" t="10000" r="25000" b="40000"/></p:blipFill><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="952500" cy="952500"/></a:xfrm></p:spPr></p:pic>
    <p:graphicFrame><p:xfrm><a:off x="0" y="0"/><a:ext cx="2857500" cy="1905000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rChart"/></a:graphicData></a:graphic></p:graphicFrame>
    <p:graphicFrame><p:xfrm><a:off x="0" y="0"/><a:ext cx="2857500" cy="1905000"/></p:xfrm><a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"><c:chart r:id="rBroken"/></a:graphicData></a:graphic></p:graphicFrame>
    <p:graphicFrame><p:xfrm><a:off x="0" y="0"/><a:ext cx="2857500" cy="952500"/></p:xfrm><a:graphic><a:graphicData><a:tbl><a:tblGrid><a:gridCol w="952500"/></a:tblGrid>
      <a:tr h="381000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>cell</a:t></a:r></a:p></a:txBody><a:tcPr marL="95250" anchor="ctr"><a:lnL w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:lnL><a:lnR><a:noFill/></a:lnR><a:lnT><a:solidFill><a:srgbClr val="00FF00"/></a:solidFill></a:lnT></a:tcPr></a:tc>
      <a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>low</a:t></a:r></a:p></a:txBody><a:tcPr anchor="b"/></a:tc></a:tr></a:tbl></a:graphicData></a:graphic></p:graphicFrame>
  </p:spTree></p:cSld></p:sld>`;

  const deck = (): Uint8Array =>
    makeZip({
      'ppt/presentation.xml': `${XML}<p:presentation ${P}><p:sldIdLst><p:sldId id="256" r:id="rSlide"/></p:sldIdLst><p:sldSz cx="9144000" cy="5143500"/></p:presentation>`,
      'ppt/_rels/presentation.xml.rels': rels([['rSlide', 'slide', 'slides/slide1.xml']]),
      'ppt/slides/slide1.xml': slide,
      'ppt/slides/_rels/slide1.xml.rels': rels([
        ['rImage', 'image', '../media/image1.png'],
        ['rChart', 'chart', '../charts/chart1.xml'],
        ['rBroken', 'chart', '../charts/chart2.xml'],
      ]),
      'ppt/media/image1.png': PNG_1X1,
      'ppt/charts/chart1.xml': PIE,
      'ppt/charts/chart2.xml': `${XML}<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"/>`,
    });

  it('draws shapes from their outlines and embeds charts', async () => {
    const presentation = await openPresentation(openZip(deck()), {
      objectUrl: () => 'blob:test/image',
    });
    const element = await presentation.renderSlide(0);
    const shapes = Array.from(element.querySelectorAll<HTMLElement>(':scope > .av-slide-shape'));
    const outline = (index: number): SVGPathElement | null =>
      shapes[index]!.querySelector('svg.av-shape-svg > path');

    expect(outline(0)!.getAttribute('d')).toBe('M0 100L100 0L200 100Z');
    expect(outline(0)!.getAttribute('fill')).toBe('#00ff00');
    expect(outline(0)!.getAttribute('stroke')).toBe('#0000ff');
    expect(Number(outline(0)!.getAttribute('stroke-width'))).toBe(2);
    // The text sits above the outline.
    expect(shapes[0]!.textContent).toBe('inside');
    expect((shapes[0]!.lastElementChild as HTMLElement).style.position).toBe('relative');

    expect(outline(1)!.getAttribute('d')).toBe('M0 0L200 100');
    expect(outline(1)!.getAttribute('fill')).toMatch(/^url\(#av-grad-/);
    expect(outline(1)!.hasAttribute('stroke')).toBe(false);
    expect(shapes[2]!.querySelector('image')!.getAttribute('href')).toBe('blob:test/image');
    expect(outline(2)!.getAttribute('stroke')).toBe('#112233');
    expect(shapes[3]!.querySelector('image')!.hasAttribute('href')).toBe(false);
    expect(outline(4)!.getAttribute('fill')).toBe('#445566');
    expect(shapes[5]!.querySelector('svg')).toBeNull();

    // Lines without a colour are black; dashes and elbows are drawn.
    expect(outline(6)!.getAttribute('stroke')).toBe('#000000');
    expect(outline(6)!.getAttribute('stroke-dasharray')).not.toBeNull();
    expect(outline(6)!.getAttribute('fill')).toBe('none');
    expect(outline(7)!.getAttribute('d')).toBe('M0 0L100 0L100 100L200 100');
    expect(outline(8)!.getAttribute('stroke-dasharray')).not.toBeNull();
    // Plain boxes stay CSS boxes; the corner radius follows the adjustment value.
    expect(shapes[9]!.querySelector('svg')).toBeNull();
    expect(shapes[9]!.style.borderRadius).toBe('50px');

    const picture = shapes[10]!;
    expect(picture.style.overflow).toBe('hidden');
    const image = picture.querySelector<HTMLElement>('img')!;
    expect(parseFloat(image.style.width)).toBeCloseTo(200, 1);
    expect(parseFloat(image.style.left)).toBeCloseTo(-50, 1);
    expect(parseFloat(image.style.height)).toBeCloseTo(200, 1);
    expect(parseFloat(image.style.top)).toBeCloseTo(-20, 1);

    const chart = shapes[11]!.querySelector('svg')!;
    expect(chart.getAttribute('width')).toBe('100%');
    expect(chart.textContent).toContain('75%');
    expect(shapes[12]!.querySelector('svg')).toBeNull();

    const [cell, low] = Array.from(element.querySelectorAll<HTMLElement>('td'));
    expect(cell!.style.borderLeft).toContain('2.67px solid');
    expect(cell!.style.borderRight).toMatch(/^none/);
    expect(cell!.style.borderTop).toContain('solid');
    expect(cell!.style.verticalAlign).toBe('middle');
    expect(cell!.style.padding).toContain('10px');
    expect(low!.style.verticalAlign).toBe('bottom');
  });
});

describe('paged paper view', () => {
  it('tracks the visible page and scrolls to a page', () => {
    const { context, container } = renderContext(new Uint8Array(), 'a.docx', 'docx');
    const pages = [0, 1, 2].map((index) => {
      const page = document.createElement('div');
      vi.spyOn(page, 'getBoundingClientRect').mockImplementation(
        () => ({ top: index * 1000, bottom: index * 1000 + 990 }) as DOMRect,
      );
      return page;
    });
    const stack = document.createElement('div');
    stack.className = 'av-doc';
    stack.append(...pages);
    const notify = vi.spyOn(context, 'notify');
    const view = createPaperView(context, stack, {
      pageCount: () => pages.length,
      page: (index) => pages[index] ?? null,
    });
    const scroller = container.querySelector<HTMLElement>('.av-pages')!;
    expect(stack.classList.contains('av-paper')).toBe(false);
    vi.spyOn(scroller, 'clientHeight', 'get').mockReturnValue(600);
    let top = 0;
    vi.spyOn(scroller, 'getBoundingClientRect').mockImplementation(() => ({ top }) as DOMRect);

    expect(view.pages!.count()).toBe(3);
    expect(view.pages!.current()).toBe(1);
    top = 1200;
    scroller.dispatchEvent(new Event('scroll'));
    expect(view.pages!.current()).toBe(2);
    expect(notify).toHaveBeenCalledWith('page');
    scroller.dispatchEvent(new Event('scroll'));
    expect(notify).toHaveBeenCalledTimes(1);
    top = 5000;
    expect(view.pages!.current()).toBe(3);

    top = 0;
    view.pages!.goTo(3);
    expect(scroller.scrollTop).toBe(2000 - 16);
    view.pages!.goTo(99);
    view.pages!.goTo(-1);
    expect(notify.mock.calls.length).toBeGreaterThanOrEqual(3);
    view.zoom!.set(2);
    expect(stack.style.getPropertyValue('zoom')).toBe('2');
    expect(view.printable!()).toBe(stack);
  });

  it('shows a single sheet without page navigation', () => {
    const { context } = renderContext(new Uint8Array(), 'a.rtf', 'rtf');
    const sheet = document.createElement('div');
    const view = createPaperView(context, sheet);
    expect(sheet.classList.contains('av-paper')).toBe(true);
    expect(view.pages).toBeUndefined();
  });
});
