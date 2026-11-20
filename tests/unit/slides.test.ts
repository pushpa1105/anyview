import { afterEach, describe, expect, it, vi } from 'vitest';
import { openOdpPresentation } from '../../src/formats/odp';
import { applyColorModifiers, openPresentation } from '../../src/formats/pptx';
import { renderer as odpRenderer } from '../../src/renderers/odp';
import { renderer as pptxRenderer } from '../../src/renderers/pptx';
import { parseXml } from '../../src/util/xml';
import { openZip } from '../../src/util/zip';
import { instanceOf, renderContext, waitFor } from '../helpers/dom';
import { odpFixture, pptxFixture } from '../helpers/fixtures';
import { makeZip } from '../helpers/zip';

afterEach(() => {
  vi.unstubAllGlobals();
  document.body.textContent = '';
});

/** The test DOM's IntersectionObserver never fires; without it every slide is built eagerly. */
function withoutIntersectionObserver(): void {
  vi.stubGlobal('IntersectionObserver', undefined);
}

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

const shapeWith = (slide: Element, text: string): HTMLElement => {
  const found = Array.from(slide.querySelectorAll<HTMLElement>('.av-slide-shape')).find((shape) =>
    shape.textContent?.includes(text),
  );
  if (!found) throw new Error(`no shape contains "${text}"`);
  return found;
};

const px = (value: string): number => parseFloat(value);

describe('applyColorModifiers', () => {
  const modifiers = (xml: string) =>
    parseXml(`<c>${xml}</c>`).children.filter((node) => typeof node !== 'string');

  it('applies luminance, tint, shade and alpha', () => {
    expect(applyColorModifiers('#4472C4', modifiers(''))).toBe('#4472c4');
    expect(applyColorModifiers('#808080', modifiers('<a:lumMod val="50000"/>'))).toBe('#404040');
    expect(applyColorModifiers('#000000', modifiers('<a:lumOff val="50000"/>'))).toBe('#808080');
    expect(applyColorModifiers('#000000', modifiers('<a:tint val="50000"/>'))).toBe('#808080');
    expect(applyColorModifiers('#ffffff', modifiers('<a:shade val="50000"/>'))).toBe('#808080');
    expect(applyColorModifiers('#ff0000', modifiers('<a:alpha val="25000"/>'))).toBe(
      'rgba(255, 0, 0, 0.25)',
    );
    expect(applyColorModifiers('#ffffff', modifiers('<a:lumOff val="90000"/><a:other/>'))).toBe(
      '#ffffff',
    );
  });
});

describe('PPTX', () => {
  const open = () => openPresentation(openZip(pptxFixture()), host());

  it('reads the slide size and count', async () => {
    const deck = await open();
    expect(deck.count).toBe(4);
    expect(deck.width).toBe(960);
    expect(deck.height).toBe(720);
  });

  it('rejects packages that are not presentations or have no slides', async () => {
    await expect(
      openPresentation(openZip(makeZip({ 'a.txt': 'x' })), host()),
    ).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
    const empty = makeZip({
      'ppt/presentation.xml': '<p:presentation><p:sldIdLst/></p:presentation>',
    });
    await expect(openPresentation(openZip(empty), host())).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
  });

  it('reports a slide whose part is missing', async () => {
    const deck = await open();
    await expect(deck.renderSlide(3)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('paints a gradient background with theme colours', async () => {
    const slide = await (await open()).renderSlide(0);
    expect(slide.style.background).toContain('linear-gradient');
    expect(slide.style.background).toContain('#ffffff 0%');
    // accent1 (#4472C4) at 50% luminance.
    expect(slide.style.background).toContain('#223962 100%');
  });

  it('inherits placeholder position and text style from layout and master', async () => {
    const slide = await (await open()).renderSlide(0);
    const title = shapeWith(slide, 'Quarterly Review');
    // Position comes from the layout placeholder, not the master.
    expect(px(title.style.left)).toBeCloseTo(72, 1);
    expect(px(title.style.top)).toBeCloseTo(42, 0);
    expect(px(title.style.width)).toBeCloseTo(816, 0);
    const paragraph = title.querySelector<HTMLElement>('p')!;
    expect(paragraph.style.textAlign).toBe('center');
    const run = paragraph.querySelector<HTMLElement>('span')!;
    expect(run.style.fontSize).toBe('44pt');
    expect(run.style.fontWeight).toBe('bold');
    expect(run.style.color).toBe('#000000');
    expect(run.style.fontFamily).toContain('Arial');
    // The master's vertical anchor reaches the slide through the placeholder chain.
    expect(title.firstElementChild!.getAttribute('style')).toContain('margin-top: auto');
  });

  it('does not show prompt text from layout or master placeholders', async () => {
    const slide = await (await open()).renderSlide(0);
    expect(slide.textContent).not.toContain('prompt');
  });

  it('draws decorative master and layout shapes behind the slide', async () => {
    const slide = await (await open()).renderSlide(0);
    expect(shapeWith(slide, 'Master logo').style.background).toBe('#00aa00');
    const stripe = Array.from(slide.querySelectorAll<HTMLElement>('.av-slide-shape')).find(
      (shape) => shape.style.background === '#ed7d31' && px(shape.style.width) === 960,
    );
    expect(stripe).toBeDefined();
  });

  it('hides master and layout shapes when the slide opts out', async () => {
    const slide = await (await open()).renderSlide(1);
    expect(slide.textContent).toBe('Second slide');
    expect(slide.querySelectorAll('.av-slide-shape')).toHaveLength(1);
  });

  it('renders bullets, numbering and auto-fit scaling', async () => {
    const slide = await (await open()).renderSlide(0);
    const body = shapeWith(slide, 'First point');
    const paragraphs = body.querySelectorAll<HTMLElement>('p');
    expect(paragraphs[0]!.textContent).toBe('•First point');
    expect(px(paragraphs[0]!.style.marginLeft)).toBeCloseTo(36, 0);
    expect(px(paragraphs[0]!.style.textIndent)).toBeCloseTo(-36, 0);
    // 32pt body text scaled to 50% by normAutofit.
    expect(paragraphs[0]!.querySelectorAll<HTMLElement>('span')[1]!.style.fontSize).toBe('16pt');
    expect(paragraphs[1]!.textContent).toBe('1.Sub one');
    expect(paragraphs[2]!.textContent).toBe('2.Sub two');
    expect(paragraphs[1]!.querySelectorAll<HTMLElement>('span')[1]!.style.fontSize).toBe('14pt');
  });

  it('renders run formatting, links, breaks and slide numbers', async () => {
    const slide = await (await open()).renderSlide(0);
    const link = slide.querySelector('a')!;
    expect(link.textContent).toBe('Styled link');
    expect(link.getAttribute('href')).toBe('https://example.com/');
    expect(link.style.fontWeight).toBe('bold');
    expect(link.style.fontStyle).toBe('italic');
    expect(link.style.textDecoration).toBe('underline line-through');
    expect(link.style.color).toBe('rgba(255, 0, 0, 0.50)');
    expect(link.style.verticalAlign).toBe('super');
    expect(link.style.textTransform).toBe('uppercase');
    expect(link.style.fontFamily).toContain('Georgia');
    // 20pt × 50% auto-fit × 70% for superscript.
    expect(link.style.fontSize).toBe('7pt');
    const paragraph = link.parentElement!;
    expect(paragraph.querySelectorAll('br')).toHaveLength(1);
    expect(paragraph.textContent).toBe('Styled link1');
    // buNone switches the inherited bullet off.
    expect(paragraph.textContent).not.toContain('•');
  });

  it('keeps empty paragraphs one line tall', async () => {
    const slide = await (await open()).renderSlide(0);
    const body = shapeWith(slide, 'First point');
    const last = Array.from(body.querySelectorAll<HTMLElement>('p')).pop()!;
    expect(last.textContent).toBe('​');
    expect(last.style.fontSize).toBe('6pt');
  });

  it('draws preset shapes with fills, outlines, rotation and flips', async () => {
    const slide = await (await open()).renderSlide(0);
    const oval = shapeWith(slide, 'In oval');
    expect(oval.style.borderRadius).toBe('50%');
    // accent1 with a 50% tint.
    expect(oval.style.background).toBe('#a2b9e2');
    expect(oval.style.border).toContain('4px solid #ff0000');
    expect(oval.style.transform).toBe('rotate(45deg) scaleX(-1)');
    const text = oval.firstElementChild as HTMLElement;
    expect(text.style.transform).toContain('scaleX(-1)');
    expect(text.style.whiteSpace).toBe('nowrap');
    expect(text.getAttribute('style')).toContain('margin-top: auto');
    expect(text.style.padding.startsWith('0px')).toBe(true);
    const paragraph = text.querySelector<HTMLElement>('p')!;
    expect(paragraph.style.textAlign).toBe('right');
    expect(Number(paragraph.style.lineHeight)).toBeCloseTo(1.8, 2);
    expect(paragraph.style.marginTop).toBe('6pt');
  });

  it('honours explicit "no fill" over style references and uses them otherwise', async () => {
    const slide = await (await open()).renderSlide(0);
    const shapes = Array.from(slide.querySelectorAll<HTMLElement>('.av-slide-shape'));
    const rounded = shapes.find((shape) => shape.style.borderRadius.endsWith('px'))!;
    expect(rounded.style.background).toBe('');
    expect(rounded.style.border).toBe('');

    const themed = shapeWith(slide, 'Themed');
    // accent2 lightened by a 20% luminance offset.
    expect(themed.style.background).toBe('#ffb064');
    expect(themed.style.border).toBe('');
    expect(themed.querySelector<HTMLElement>('span')!.style.color).toBe('#ffffff');
  });

  it('draws connectors as lines', async () => {
    const slide = await (await open()).renderSlide(0);
    const line = slide.querySelector('svg path')!;
    expect(line.getAttribute('stroke')).toBe('#333333');
    expect(Number(line.getAttribute('stroke-width'))).toBeCloseTo(1.33, 1);
  });

  it('embeds pictures and skips formats browsers cannot draw', async () => {
    const fake = host();
    const deck = await openPresentation(openZip(pptxFixture()), fake);
    const slide = await deck.renderSlide(0);
    const images = slide.querySelectorAll('img');
    // The picture and the OLE preview; the WMF and the unplaced picture are skipped.
    expect(images).toHaveLength(2);
    expect(images[0]!.alt).toBe('A logo');
    expect(images[0]!.getAttribute('src')).toMatch(/^blob:test\//);
    expect(fake.blobs.every((blob) => blob.type === 'image/png')).toBe(true);
    expect(px(images[0]!.parentElement!.style.left)).toBeCloseTo(630, 0);
  });

  it('applies group transforms to child shapes', async () => {
    const slide = await (await open()).renderSlide(0);
    const child = Array.from(slide.querySelectorAll<HTMLElement>('.av-slide-shape')).find(
      (shape) => shape.style.background === '#123456',
    )!;
    // Group at (1000000, 4000000) scaled by 0.5; child at (2000000, 1000000) sized 2000000 × 1000000.
    expect(px(child.style.left)).toBeCloseTo(2000000 / 9525, 0);
    expect(px(child.style.top)).toBeCloseTo(4500000 / 9525, 0);
    expect(px(child.style.width)).toBeCloseTo(1000000 / 9525, 0);
    expect(px(child.style.height)).toBeCloseTo(500000 / 9525, 0);
  });

  it('renders tables and placeholders for charts', async () => {
    const slide = await (await open()).renderSlide(0);
    const table = slide.querySelector('table')!;
    expect(table.querySelectorAll('col')).toHaveLength(2);
    const rows = table.querySelectorAll('tr');
    expect(rows[0]!.children).toHaveLength(1);
    expect((rows[0]!.children[0] as HTMLTableCellElement).colSpan).toBe(2);
    expect((rows[0]!.children[0] as HTMLElement).style.background).toBe('#cccccc');
    expect(rows[1]!.textContent).toBe('Cell ACell B');
    expect(shapeWith(slide, 'Chart').style.border).toContain('dashed');
  });

  it('prefers the fallback of alternate content and skips unplaced shapes', async () => {
    const slide = await (await open()).renderSlide(0);
    expect(slide.textContent).toContain('Fallback content');
    expect(slide.textContent).not.toContain('Choice content');
    expect(slide.textContent).not.toContain('Unplaced shape');
  });

  it('uses an image as the slide background', async () => {
    const slide = await (await open()).renderSlide(1);
    expect(slide.style.backgroundImage).toContain('blob:test/');
    expect(slide.style.backgroundSize).toBe('100% 100%');
  });

  it('falls back to the master background and handles single-stop gradients', async () => {
    const slide = await (await open()).renderSlide(2);
    expect(slide.style.background).toBe('#ffffff');
    const box = slide.querySelector<HTMLElement>('.av-slide-shape:last-child')!;
    expect(box.style.background).toBe('#ff0000');
  });

  it('shows slides in a scrollable view with page and zoom controls', async () => {
    withoutIntersectionObserver();
    const { context, container, notifications } = renderContext(pptxFixture(), 'deck.pptx', 'pptx');
    const instance = instanceOf(await pptxRenderer.render(context));
    const frames = container.querySelectorAll<HTMLElement>('.av-page');
    expect(frames).toHaveLength(4);
    expect(instance.pages!.count()).toBe(4);
    expect(instance.pages!.current()).toBe(1);
    await waitFor(() => container.querySelectorAll('.av-slide').length === 3);
    // The slide with a missing part shows a message instead of breaking the deck.
    await waitFor(() => frames[3]!.textContent?.includes('missing'));

    const slide = container.querySelector<HTMLElement>('.av-slide')!;
    expect(slide.style.width).toBe('960px');
    expect(slide.style.transform).toBe('scale(1)');
    instance.zoom!.set(2);
    expect(instance.zoom!.get()).toBe(2);
    expect(slide.style.transform).toBe('scale(2)');
    expect(frames[0]!.style.width).toBe('1920px');

    instance.pages!.goTo(3);
    expect(instance.pages!.current()).toBe(3);
    instance.pages!.goTo(99);
    expect(instance.pages!.current()).toBe(3);

    const scroller = container.querySelector('.av-pages')!;
    scroller.dispatchEvent(new Event('scroll'));
    await waitFor(() => notifications.includes('page'));
    expect(instance.pages!.current()).toBe(1);
    expect(instance.printable!()).toBe(scroller);
    instance.destroy!();
  });

  it('builds slides lazily as they scroll into view', async () => {
    const observed: Element[] = [];
    let trigger: (entries: Partial<IntersectionObserverEntry>[]) => void = () => undefined;
    vi.stubGlobal(
      'IntersectionObserver',
      class {
        constructor(callback: (entries: Partial<IntersectionObserverEntry>[]) => void) {
          trigger = callback;
        }
        observe(element: Element): void {
          observed.push(element);
        }
        disconnect(): void {
          observed.length = 0;
        }
      },
    );
    const { context, container } = renderContext(pptxFixture(), 'deck.pptx', 'pptx');
    const instance = instanceOf(await pptxRenderer.render(context));
    expect(observed).toHaveLength(4);
    expect(container.querySelectorAll('.av-slide')).toHaveLength(0);
    trigger([
      { isIntersecting: false, target: observed[0]! },
      { isIntersecting: true, target: observed[1]! },
    ]);
    await waitFor(() => container.querySelectorAll('.av-slide').length === 1);
    expect(container.querySelector('.av-slide')?.textContent).toBe('Second slide');
    // A second notification for the same slide does not rebuild it.
    trigger([{ isIntersecting: true, target: observed[1]! }]);
    instance.destroy!();
    expect(observed).toHaveLength(0);
  });

  it('does not touch the DOM after being destroyed mid-render', async () => {
    withoutIntersectionObserver();
    const { context, container } = renderContext(pptxFixture(), 'deck.pptx', 'pptx');
    const instance = instanceOf(await pptxRenderer.render(context));
    instance.destroy!();
    await new Promise((resolve) => setTimeout(resolve, 50));
    expect(container.querySelectorAll('.av-slide')).toHaveLength(0);
  });
});

describe('ODP', () => {
  const open = () => openOdpPresentation(openZip(odpFixture()), host());

  it('reads the page size and count', async () => {
    const deck = await open();
    expect(deck.count).toBe(2);
    expect(deck.width).toBeCloseTo(1058.27, 1);
    expect(deck.height).toBeCloseTo(595.28, 1);
  });

  it('rejects packages without content or slides', async () => {
    await expect(
      openOdpPresentation(openZip(makeZip({ 'a.txt': 'x' })), host()),
    ).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
    const empty = makeZip({
      'content.xml':
        '<office:document-content><office:body><office:presentation/></office:body></office:document-content>',
    });
    await expect(openOdpPresentation(openZip(empty), host())).rejects.toMatchObject({
      code: 'PARSE_FAILED',
    });
    await expect((await open()).renderSlide(9)).rejects.toMatchObject({ code: 'PARSE_FAILED' });
  });

  it('places frames, text, pictures and shapes', async () => {
    const fake = host();
    const deck = await openOdpPresentation(openZip(odpFixture()), fake);
    const slide = await deck.renderSlide(0);
    expect(slide.style.background).toBe('#102030');

    const title = shapeWith(slide, 'Deck title');
    expect(px(title.style.left)).toBeCloseTo(37.8, 1);
    expect(px(title.style.width)).toBeCloseTo(755.9, 1);
    expect((title.firstElementChild as HTMLElement).style.fontSize).toBe('40pt');

    const subtitle = shapeWith(slide, 'Sub');
    expect((subtitle.firstElementChild as HTMLElement).style.fontSize).toBe('28pt');
    expect(subtitle.querySelector<HTMLElement>('span')!.style.fontWeight).toBe('bold');
    expect(subtitle.querySelector('p')!.textContent).toBe('Sub bold x\tyz');
    expect(subtitle.querySelectorAll('br')).toHaveLength(1);
    expect(subtitle.querySelector('ul li')?.textContent).toBe('bullet');

    expect(slide.querySelectorAll('img')).toHaveLength(1);
    expect(fake.blobs).toHaveLength(1);

    const rect = shapeWith(slide, 'In rect');
    expect(rect.style.background).toBe('#ff0000');
    expect(rect.style.border).toContain('1px solid #0000ff');
    const shapes = Array.from(slide.querySelectorAll<HTMLElement>('.av-slide-shape'));
    const ellipse = shapes.find((shape) => shape.style.borderRadius === '50%')!;
    expect(ellipse.style.background).toBe('');
    expect(shapeWith(slide, 'Custom').style.height).toBe('');
    expect(shapeWith(slide, 'Grouped')).toBeDefined();

    const horizontal = shapes.find((shape) => shape.style.height === '1px')!;
    expect(horizontal.style.borderTop).toContain('solid');
    expect(px(horizontal.style.width)).toBeCloseTo(340.2, 0);
    expect(shapes.find((shape) => shape.style.width === '1px')!.style.borderLeft).toContain(
      'solid',
    );
    expect(slide.textContent).not.toContain('Speaker notes');
  });

  it('renders later slides and the slide view', async () => {
    withoutIntersectionObserver();
    const second = await (await open()).renderSlide(1);
    expect(second.textContent).toBe('Second page');
    expect(second.style.background).toBe('');

    const { context, container } = renderContext(odpFixture(), 'deck.odp', 'odp');
    const instance = instanceOf(await odpRenderer.render(context));
    expect(instance.pages!.count()).toBe(2);
    await waitFor(() => container.querySelectorAll('.av-slide').length === 2);
  });
});
