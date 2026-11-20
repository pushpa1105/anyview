import { makeZip, PNG_1X1 } from './zip';

const XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const REL_NS = 'http://schemas.openxmlformats.org/package/2006/relationships';
const OFFICE_REL = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

function relationships(
  entries: [id: string, type: string, target: string, external?: boolean][],
): string {
  const items = entries
    .map(
      ([id, type, target, external]) =>
        `<Relationship Id="${id}" Type="${OFFICE_REL}/${type}" Target="${target}"${external ? ' TargetMode="External"' : ''}/>`,
    )
    .join('');
  return `${XML}<Relationships xmlns="${REL_NS}">${items}</Relationships>`;
}

// ---------------------------------------------------------------------------
// XLSX
// ---------------------------------------------------------------------------

export function xlsxFixture(): Uint8Array {
  const workbook = `${XML}<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="${OFFICE_REL}">
  <workbookPr/>
  <sheets>
    <sheet name="Sales" sheetId="1" r:id="rId1"/>
    <sheet name="Hidden" sheetId="2" state="hidden" r:id="rId2"/>
    <sheet name="Notes" sheetId="3" r:id="rId3"/>
    <sheet name="Broken" sheetId="4" r:id="rId4"/>
  </sheets>
</workbook>`;
  const styles = `${XML}<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts>
  <fonts count="2">
    <font><sz val="11"/><color rgb="FF000000"/></font>
    <font><b/><i/><u/><strike/><color rgb="FFFF0000"/></font>
  </fonts>
  <fills count="3">
    <fill><patternFill patternType="none"/></fill>
    <fill><patternFill patternType="gray125"/></fill>
    <fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill>
  </fills>
  <cellXfs count="6">
    <xf numFmtId="0" fontId="0" fillId="0"/>
    <xf numFmtId="164" fontId="0" fillId="0"/>
    <xf numFmtId="4" fontId="0" fillId="0"/>
    <xf numFmtId="0" fontId="1" fillId="2"><alignment horizontal="center"/></xf>
    <xf numFmtId="9" fontId="0" fillId="0"/>
    <xf numFmtId="0" fontId="0" fillId="2"/>
  </cellXfs>
</styleSheet>`;
  const sharedStrings = `${XML}<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <si><t>Product</t></si>
  <si><t>Price</t></si>
  <si><r><t>Rich </t></r><r><rPr><b/></rPr><t>text</t></r><rPh><t>ignored</t></rPh></si>
  <si><t xml:space="preserve"> padded </t></si>
</sst>`;
  const sheet1 = `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
  <cols><col min="1" max="1" width="20" customWidth="1"/><col min="2" max="2" width="5" hidden="1"/></cols>
  <sheetData>
    <row r="1">
      <c r="A1" t="s" s="3"><v>0</v></c>
      <c r="B1" t="s"><v>1</v></c>
      <c r="C1" t="s"><v>2</v></c>
      <c r="D1" t="s"><v>3</v></c>
    </row>
    <row r="2">
      <c r="A2" t="inlineStr"><is><t>Inline</t></is></c>
      <c r="B2" s="2"><v>1234.5</v></c>
      <c r="C2" s="1"><v>45000</v></c>
      <c r="D2" s="4"><v>0.256</v></c>
    </row>
    <row r="3">
      <c r="A3" t="b"><v>1</v></c>
      <c r="B3" t="b"><v>0</v></c>
      <c r="C3" t="e"><v>#DIV/0!</v></c>
      <c r="D3" t="str"><f>A1&amp;B1</f><v>formula text</v></c>
    </row>
    <row r="4">
      <c r="A4" t="d"><v>2024-01-31T10:00:00Z</v></c>
      <c r="B4"><v>42</v></c>
      <c r="C4"/>
      <c r="D4" s="5"/>
      <c r="E4"><v>not-a-number</v></c>
    </row>
    <row><c><v>7</v></c><c><v>8</v></c></row>
  </sheetData>
  <mergeCells count="2"><mergeCell ref="A6:C7"/><mergeCell ref="bogus"/></mergeCells>
</worksheet>`;
  const sheet3 = `${XML}<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData/></worksheet>`;

  return makeZip({
    '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'xl/workbook.xml']]),
    'xl/workbook.xml': workbook,
    'xl/_rels/workbook.xml.rels': relationships([
      ['rId1', 'worksheet', 'worksheets/sheet1.xml'],
      ['rId2', 'worksheet', 'worksheets/sheet2.xml'],
      ['rId3', 'worksheet', 'worksheets/sheet3.xml'],
      ['rId4', 'worksheet', 'worksheets/missing.xml'],
      ['rId5', 'styles', 'styles.xml'],
    ]),
    'xl/styles.xml': styles,
    'xl/sharedStrings.xml': sharedStrings,
    'xl/worksheets/sheet1.xml': sheet1,
    'xl/worksheets/sheet2.xml': sheet3,
    'xl/worksheets/sheet3.xml': sheet3,
  });
}

// ---------------------------------------------------------------------------
// DOCX
// ---------------------------------------------------------------------------

const W_NS =
  'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
  `xmlns:r="${OFFICE_REL}" ` +
  'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ' +
  'xmlns:v="urn:schemas-microsoft-com:vml" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" ' +
  'xmlns:wps="http://schemas.microsoft.com/office/word/2010/wordprocessingShape"';

export interface DocxFixtureOptions {
  /** Body XML; defaults to a document exercising most supported features. */
  body?: string;
  /** Leave out styles.xml, numbering.xml and footnotes.xml. */
  minimal?: boolean;
}

function drawing(relationshipId: string): string {
  return `<w:drawing><wp:inline><wp:extent cx="952500" cy="476250"/><a:graphic><a:graphicData><pic:pic><pic:blipFill><a:blip r:embed="${relationshipId}"/></pic:blipFill></pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing>`;
}

export const DOCX_BODY = `
<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Annual Report</w:t></w:r></w:p>
<w:p><w:pPr><w:pStyle w:val="Custom"/><w:jc w:val="center"/></w:pPr>
  <w:r><w:t xml:space="preserve">Plain </w:t></w:r>
  <w:r><w:rPr><w:b/><w:i/><w:u w:val="single"/><w:strike/><w:color w:val="FF0000"/><w:sz w:val="28"/><w:highlight w:val="yellow"/><w:rFonts w:ascii="Arial"/></w:rPr><w:t>styled</w:t></w:r>
  <w:r><w:rPr><w:b w:val="0"/><w:vertAlign w:val="superscript"/></w:rPr><w:t>sup</w:t></w:r>
  <w:r><w:rPr><w:rStyle w:val="Emphasis"/><w:caps/><w:smallCaps/></w:rPr><w:t>char style</w:t></w:r>
  <w:r><w:rPr><w:vanish/></w:rPr><w:t>hidden text</w:t></w:r>
  <w:r><w:tab/><w:br/><w:cr/><w:noBreakHyphen/><w:sym w:font="Wingdings" w:char="F0FC"/><w:instrText>FIELD CODE</w:instrText></w:r>
</w:p>
<w:p><w:pPr><w:ind w:left="720" w:right="360" w:firstLine="240"/><w:spacing w:before="120" w:after="240" w:line="360" w:lineRule="auto"/><w:shd w:fill="EEEEEE"/><w:pageBreakBefore/><w:jc w:val="both"/></w:pPr>
  <w:hyperlink r:id="rIdLink"><w:r><w:t>external link</w:t></w:r></w:hyperlink>
  <w:hyperlink r:id="rIdBad"><w:r><w:t>unsafe link</w:t></w:r></w:hyperlink>
  <w:hyperlink w:anchor="bookmark"><w:r><w:t>internal link</w:t></w:r></w:hyperlink>
  <w:ins><w:r><w:t>inserted</w:t></w:r></w:ins>
  <w:del><w:r><w:delText>deleted</w:delText></w:r></w:del>
  <w:sdt><w:sdtContent><w:r><w:t>content control</w:t></w:r></w:sdtContent></w:sdt>
  <w:fldSimple w:instr="PAGE"><w:r><w:t>7</w:t></w:r></w:fldSimple>
  <w:r><w:footnoteReference w:id="2"/></w:r>
  <w:r><w:footnoteReference w:id="99"/></w:r>
</w:p>
<w:p><w:pPr><w:spacing w:line="280" w:lineRule="exact"/><w:jc w:val="right"/></w:pPr><w:r><w:br w:type="page"/><w:t>after break</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>first</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>nested a</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>nested b</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr><w:r><w:t>second</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="2"/></w:numPr><w:ind w:left="1440" w:hanging="720"/></w:pPr><w:r><w:t>bullet</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="3"/></w:numPr></w:pPr><w:r><w:t>restarted</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>not numbered</w:t></w:r></w:p>
<w:p><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="77"/></w:numPr></w:pPr><w:r><w:t>unknown list</w:t></w:r></w:p>
<w:p><w:pPr><w:outlineLvl w:val="2"/></w:pPr><w:r><w:t>Outline heading</w:t></w:r></w:p>
<w:tbl>
  <w:tblPr><w:tblStyle w:val="TableGrid"/><w:jc w:val="center"/></w:tblPr>
  <w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
  <w:tr>
    <w:tc><w:tcPr><w:gridSpan w:val="2"/><w:shd w:fill="DDEEFF"/><w:vAlign w:val="center"/></w:tcPr><w:p><w:r><w:t>wide</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:vMerge w:val="restart"/><w:vAlign w:val="bottom"/></w:tcPr><w:p><w:r><w:t>tall</w:t></w:r></w:p></w:tc>
  </w:tr>
  <w:tr>
    <w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:tcBorders><w:top w:val="nil"/></w:tcBorders></w:tcPr><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc>
    <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p/></w:tc>
  </w:tr>
</w:tbl>
<w:tbl><w:tblPr><w:tblBorders><w:top w:val="none"/></w:tblBorders><w:jc w:val="right"/></w:tblPr>
  <w:tr><w:tc><w:p><w:r><w:t>borderless</w:t></w:r></w:p></w:tc></w:tr>
</w:tbl>
<w:p><w:r>${drawing('rIdImage')}</w:r><w:r>${drawing('rIdEmf')}</w:r><w:r>${drawing('rIdMissing')}</w:r>
  <w:r><w:pict><v:shape style="width:10pt"><v:imagedata r:id="rIdImage"/></v:shape></w:pict></w:r>
  <w:r><mc:AlternateContent><mc:Choice Requires="wps"><w:drawing><wp:anchor><wp:extent cx="1905000" cy="952500"/><a:graphic><a:graphicData><wps:wsp><wps:txbx><w:txbxContent><w:p><w:r><w:t>text box</w:t></w:r></w:p></w:txbxContent></wps:txbx></wps:wsp></a:graphicData></a:graphic></wp:anchor></w:drawing></mc:Choice><mc:Fallback><w:pict/></mc:Fallback></mc:AlternateContent></w:r>
</w:p>
<w:sdt><w:sdtContent><w:p><w:r><w:t>block control</w:t></w:r></w:p></w:sdtContent></w:sdt>
<w:sectPr><w:pgSz w:w="11906" w:h="16838"/><w:pgMar w:top="1440" w:right="1080" w:bottom="1440" w:left="1080"/></w:sectPr>`;

export function docxFixture(options: DocxFixtureOptions = {}): Uint8Array {
  const document = `${XML}<w:document ${W_NS}><w:body>${options.body ?? DOCX_BODY}</w:body></w:document>`;
  const styles = `${XML}<w:styles ${W_NS}>
  <w:docDefaults>
    <w:rPrDefault><w:rPr><w:rFonts w:ascii="Calibri"/><w:sz w:val="22"/></w:rPr></w:rPrDefault>
    <w:pPrDefault><w:pPr><w:spacing w:after="160"/></w:pPr></w:pPrDefault>
  </w:docDefaults>
  <w:style w:type="paragraph" w:default="1" w:styleId="Normal"><w:name w:val="Normal"/></w:style>
  <w:style w:type="paragraph" w:styleId="Heading1"><w:name w:val="heading 1"/><w:basedOn w:val="Normal"/><w:rPr><w:b/><w:sz w:val="32"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Custom"><w:name w:val="Custom"/><w:basedOn w:val="Loop"/><w:rPr><w:color w:val="auto"/></w:rPr></w:style>
  <w:style w:type="paragraph" w:styleId="Loop"><w:name w:val="Loop"/><w:basedOn w:val="Custom"/></w:style>
  <w:style w:type="character" w:styleId="Emphasis"><w:name w:val="Emphasis"/><w:rPr><w:i/></w:rPr></w:style>
  <w:style w:type="table" w:styleId="TableGrid"><w:name w:val="Table Grid"/><w:tblPr><w:tblBorders><w:top w:val="single"/></w:tblBorders></w:tblPr></w:style>
  <w:style w:type="paragraph"><w:name w:val="No id"/></w:style>
</w:styles>`;
  const numbering = `${XML}<w:numbering ${W_NS}>
  <w:abstractNum w:abstractNumId="0">
    <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1."/><w:pPr><w:ind w:left="720" w:hanging="360"/></w:pPr></w:lvl>
    <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="lowerLetter"/><w:lvlText w:val="%1.%2)"/><w:pPr><w:ind w:left="1440" w:hanging="360"/></w:pPr></w:lvl>
  </w:abstractNum>
  <w:abstractNum w:abstractNumId="1">
    <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val="&#xF0B7;"/><w:rPr><w:rFonts w:ascii="Symbol"/></w:rPr></w:lvl>
  </w:abstractNum>
  <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
  <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
  <w:num w:numId="3"><w:abstractNumId w:val="0"/><w:lvlOverride w:ilvl="0"><w:startOverride w:val="5"/></w:lvlOverride></w:num>
  <w:num><w:abstractNumId w:val="0"/></w:num>
</w:numbering>`;
  const footnotes = `${XML}<w:footnotes ${W_NS}>
  <w:footnote w:type="separator" w:id="0"><w:p/></w:footnote>
  <w:footnote w:id="2"><w:p><w:r><w:t>The footnote text.</w:t></w:r></w:p></w:footnote>
</w:footnotes>`;

  const documentRelationships: [string, string, string, boolean?][] = [
    ['rIdLink', 'hyperlink', 'https://example.com/', true],
    ['rIdBad', 'hyperlink', 'javascript:alert(1)', true],
    ['rIdImage', 'image', 'media/image1.png'],
    ['rIdEmf', 'image', 'media/image2.emf'],
  ];
  if (!options.minimal) {
    documentRelationships.push(
      ['rIdStyles', 'styles', 'styles.xml'],
      ['rIdNumbering', 'numbering', 'numbering.xml'],
      ['rIdFootnotes', 'footnotes', 'footnotes.xml'],
    );
  }
  const files: Record<string, string | Uint8Array> = {
    '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'word/document.xml']]),
    'word/document.xml': document,
    'word/_rels/document.xml.rels': relationships(documentRelationships),
    'word/media/image1.png': PNG_1X1,
    'word/media/image2.emf': new Uint8Array([1, 2, 3]),
  };
  if (!options.minimal) {
    files['word/styles.xml'] = styles;
    files['word/numbering.xml'] = numbering;
    files['word/footnotes.xml'] = footnotes;
  }
  return makeZip(files);
}

// ---------------------------------------------------------------------------
// PPTX
// ---------------------------------------------------------------------------

const P_NS =
  'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
  `xmlns:r="${OFFICE_REL}" ` +
  'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
  'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"';

const xfrm = (x: number, y: number, cx: number, cy: number, extra = ''): string =>
  `<a:xfrm${extra}><a:off x="${x}" y="${y}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>`;

export function pptxFixture(): Uint8Array {
  const presentation = `${XML}<p:presentation ${P_NS}>
  <p:sldIdLst><p:sldId id="256" r:id="rId2"/><p:sldId id="257" r:id="rId3"/><p:sldId id="258" r:id="rId4"/><p:sldId id="259" r:id="rIdGone"/></p:sldIdLst>
  <p:sldSz cx="9144000" cy="6858000"/>
  <p:defaultTextStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:defaultTextStyle>
</p:presentation>`;
  const theme = `${XML}<a:theme ${P_NS}><a:themeElements><a:clrScheme name="x">
  <a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1>
  <a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2>
  <a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2>
</a:clrScheme></a:themeElements></a:theme>`;
  const master = `${XML}<p:sldMaster ${P_NS}>
  <p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr>${xfrm(457200, 274638, 8229600, 1143000)}</p:spPr>
      <p:txBody><a:bodyPr anchor="ctr"/><a:p><a:r><a:t>Master title prompt</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Body"/><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr>${xfrm(457200, 1600200, 8229600, 4525963)}</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="9" name="Logo"/><p:nvPr/></p:nvSpPr>
      <p:spPr>${xfrm(8000000, 6000000, 900000, 600000)}<a:solidFill><a:srgbClr val="00AA00"/></a:solidFill></p:spPr>
      <p:txBody><a:bodyPr/><a:p><a:r><a:t>Master logo</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld>
  <p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2"/>
  <p:txStyles>
    <p:titleStyle><a:lvl1pPr algn="ctr"><a:defRPr sz="4400" b="1"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="Arial"/></a:defRPr></a:lvl1pPr></p:titleStyle>
    <p:bodyStyle>
      <a:lvl1pPr marL="342900" indent="-342900"><a:buChar char="&#x2022;"/><a:defRPr sz="3200"/></a:lvl1pPr>
      <a:lvl2pPr marL="742950" indent="-285750"><a:buAutoNum type="arabicPeriod"/><a:defRPr sz="2800"/></a:lvl2pPr>
    </p:bodyStyle>
    <p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"/></a:lvl1pPr></p:otherStyle>
  </p:txStyles>
</p:sldMaster>`;
  const layout = `${XML}<p:sldLayout ${P_NS}><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr>
      <p:spPr>${xfrm(685800, 400000, 7772400, 1470025)}</p:spPr><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:t>Layout title prompt</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Content"/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr>
      <p:spPr>${xfrm(685800, 2000000, 7772400, 4000000)}</p:spPr><p:txBody><a:bodyPr/><a:p/></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="8" name="Stripe"/><p:nvPr/></p:nvSpPr>
      <p:spPr>${xfrm(0, 0, 9144000, 100000)}<a:solidFill><a:schemeClr val="accent2"/></a:solidFill></p:spPr></p:sp>
  </p:spTree></p:cSld></p:sldLayout>`;
  const slide1 = `${XML}<p:sld ${P_NS}><p:cSld>
  <p:bg><p:bgPr><a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:schemeClr val="accent1"><a:lumMod val="50000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000"/></a:gradFill></p:bgPr></p:bg>
  <p:spTree>
    <p:nvGrpSpPr><p:cNvPr id="1" name=""/></p:nvGrpSpPr><p:grpSpPr/>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/>
      <p:txBody><a:bodyPr/><a:p><a:r><a:rPr lang="en-US"/><a:t>Quarterly Review</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="3" name="Content 2"/><p:nvPr><p:ph idx="1"/></p:nvPr></p:nvSpPr><p:spPr/>
      <p:txBody><a:bodyPr><a:normAutofit fontScale="50000"/></a:bodyPr>
        <a:p><a:r><a:t>First point</a:t></a:r></a:p>
        <a:p><a:pPr lvl="1"/><a:r><a:t>Sub one</a:t></a:r></a:p>
        <a:p><a:pPr lvl="1"/><a:r><a:t>Sub two</a:t></a:r></a:p>
        <a:p><a:pPr><a:buNone/></a:pPr><a:r><a:rPr b="1" i="1" u="sng" strike="sngStrike" sz="2000" baseline="30000" cap="all"><a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill><a:latin typeface="Georgia"/><a:hlinkClick r:id="rIdLink"/></a:rPr><a:t>Styled link</a:t></a:r><a:br/><a:fld type="slidenum"><a:t>‹#›</a:t></a:fld></a:p>
        <a:p><a:endParaRPr sz="1200"/></a:p>
      </p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="4" name="Oval"/><p:nvPr/></p:nvSpPr>
      <p:spPr>${xfrm(1000000, 1000000, 2000000, 1000000, ' rot="2700000" flipH="1"')}<a:prstGeom prst="ellipse"/><a:solidFill><a:schemeClr val="accent1"><a:tint val="50000"/></a:schemeClr></a:solidFill><a:ln w="38100"><a:solidFill><a:prstClr val="red"/></a:solidFill></a:ln></p:spPr>
      <p:txBody><a:bodyPr anchor="b" wrap="none" lIns="0" tIns="0"/><a:p><a:pPr algn="r"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr><a:r><a:t>In oval</a:t></a:r></a:p></p:txBody></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="5" name="Rounded"/><p:nvPr/></p:nvSpPr>
      <p:spPr>${xfrm(4000000, 1000000, 1000000, 500000)}<a:prstGeom prst="roundRect"/><a:noFill/><a:ln><a:noFill/></a:ln></p:spPr>
      <p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent2"/></a:fillRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style></p:sp>
    <p:sp><p:nvSpPr><p:cNvPr id="6" name="Styled"/><p:nvPr/></p:nvSpPr>
      <p:spPr>${xfrm(5200000, 1000000, 1000000, 500000)}</p:spPr>
      <p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent2"><a:lumOff val="20000"/></a:schemeClr></a:fillRef><a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>
      <p:txBody><a:bodyPr anchor="ctr"/><a:p><a:pPr algn="ctr"/><a:r><a:t>Themed</a:t></a:r></a:p></p:txBody></p:sp>
    <p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="Line"/></p:nvCxnSpPr>
      <p:spPr>${xfrm(100000, 5000000, 3000000, 0)}<a:prstGeom prst="line"/><a:ln w="12700"><a:solidFill><a:srgbClr val="333333"/></a:solidFill></a:ln></p:spPr></p:cxnSp>
    <p:pic><p:nvPicPr><p:cNvPr id="8" name="Picture" descr="A logo"/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill><p:spPr>${xfrm(6000000, 3000000, 952500, 952500)}</p:spPr></p:pic>
    <p:pic><p:nvPicPr><p:cNvPr id="9" name="Unsupported"/><p:nvPr/></p:nvPicPr>
      <p:blipFill><a:blip r:embed="rIdWmf"/></p:blipFill><p:spPr>${xfrm(0, 0, 100, 100)}</p:spPr></p:pic>
    <p:pic><p:nvPicPr><p:cNvPr id="10" name="No position"/><p:nvPr/></p:nvPicPr><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill><p:spPr/></p:pic>
    <p:grpSp><p:nvGrpSpPr><p:cNvPr id="11" name="Group"/></p:nvGrpSpPr>
      <p:grpSpPr><a:xfrm><a:off x="1000000" y="4000000"/><a:ext cx="2000000" cy="1000000"/><a:chOff x="0" y="0"/><a:chExt cx="4000000" cy="2000000"/></a:xfrm></p:grpSpPr>
      <p:sp><p:nvSpPr><p:cNvPr id="12" name="Child"/><p:nvPr/></p:nvSpPr>
        <p:spPr>${xfrm(2000000, 1000000, 2000000, 1000000)}<a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr></p:sp>
    </p:grpSp>
    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="13" name="Table"/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="500000" y="5200000"/><a:ext cx="4000000" cy="800000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/table"><a:tbl>
        <a:tblGrid><a:gridCol w="2000000"/><a:gridCol w="2000000"/></a:tblGrid>
        <a:tr h="400000"><a:tc gridSpan="2"><a:txBody><a:bodyPr/><a:p><a:r><a:t>Header</a:t></a:r></a:p></a:txBody><a:tcPr><a:solidFill><a:srgbClr val="CCCCCC"/></a:solidFill></a:tcPr></a:tc><a:tc hMerge="1"><a:txBody><a:bodyPr/><a:p/></a:txBody></a:tc></a:tr>
        <a:tr h="400000"><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell A</a:t></a:r></a:p></a:txBody></a:tc><a:tc><a:txBody><a:bodyPr/><a:p><a:r><a:t>Cell B</a:t></a:r></a:p></a:txBody></a:tc></a:tr>
      </a:tbl></a:graphicData></a:graphic></p:graphicFrame>
    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="14" name="Chart"/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="5000000" y="5200000"/><a:ext cx="2000000" cy="800000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/chart"/></a:graphic></p:graphicFrame>
    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="15" name="Ole"/><p:nvPr/></p:nvGraphicFramePr>
      <p:xfrm><a:off x="7200000" y="5200000"/><a:ext cx="1000000" cy="800000"/></p:xfrm>
      <a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/presentationml/2006/ole"><p:oleObj><p:pic><p:blipFill><a:blip r:embed="rIdImage"/></p:blipFill></p:pic></p:oleObj></a:graphicData></a:graphic></p:graphicFrame>
    <p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="16" name="No box"/><p:nvPr/></p:nvGraphicFramePr><a:graphic><a:graphicData uri="x"/></a:graphic></p:graphicFrame>
    <mc:AlternateContent><mc:Choice Requires="x"><p:sp><p:nvSpPr><p:cNvPr id="17" name="Modern"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(0, 0, 10, 10)}</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Choice content</a:t></a:r></a:p></p:txBody></p:sp></mc:Choice>
      <mc:Fallback><p:sp><p:nvSpPr><p:cNvPr id="18" name="Compatible"/><p:nvPr/></p:nvSpPr><p:spPr>${xfrm(0, 0, 10, 10)}</p:spPr><p:txBody><a:bodyPr/><a:p><a:r><a:t>Fallback content</a:t></a:r></a:p></p:txBody></p:sp></mc:Fallback></mc:AlternateContent>
    <p:sp><p:nvSpPr><p:cNvPr id="19" name="No box"/><p:nvPr/></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:p><a:r><a:t>Unplaced shape</a:t></a:r></a:p></p:txBody></p:sp>
  </p:spTree></p:cSld></p:sld>`;
  const slide2 = `${XML}<p:sld ${P_NS} showMasterSp="0"><p:cSld>
  <p:bg><p:bgPr><a:blipFill><a:blip r:embed="rIdImage"/><a:stretch/></a:blipFill></p:bgPr></p:bg>
  <p:spTree><p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:nvPr/></p:nvSpPr>
    <p:spPr>${xfrm(0, 0, 1000000, 1000000)}<a:solidFill><a:srgbClr val="ABCDEF"/></a:solidFill></p:spPr>
    <p:txBody><a:bodyPr/><a:p><a:r><a:t>Second slide</a:t></a:r></a:p></p:txBody></p:sp></p:spTree></p:cSld></p:sld>`;
  const slide3 = `${XML}<p:sld ${P_NS}><p:cSld><p:spTree>
    <p:sp><p:nvSpPr><p:cNvPr id="2" name="Box"/><p:nvPr/></p:nvSpPr>
    <p:spPr>${xfrm(0, 0, 1000000, 1000000)}<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs></a:gsLst></a:gradFill></p:spPr></p:sp>
  </p:spTree></p:cSld></p:sld>`;

  const slideRelationships = (extra: [string, string, string, boolean?][] = []): string =>
    relationships([['rIdLayout', 'slideLayout', '../slideLayouts/slideLayout1.xml'], ...extra]);

  return makeZip({
    '[Content_Types].xml': `${XML}<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"/>`,
    '_rels/.rels': relationships([['rId1', 'officeDocument', 'ppt/presentation.xml']]),
    'ppt/presentation.xml': presentation,
    'ppt/_rels/presentation.xml.rels': relationships([
      ['rId1', 'slideMaster', 'slideMasters/slideMaster1.xml'],
      ['rId2', 'slide', 'slides/slide1.xml'],
      ['rId3', 'slide', 'slides/slide2.xml'],
      ['rId4', 'slide', 'slides/slide3.xml'],
      ['rIdGone', 'slide', 'slides/missing.xml'],
    ]),
    'ppt/theme/theme1.xml': theme,
    'ppt/slideMasters/slideMaster1.xml': master,
    'ppt/slideMasters/_rels/slideMaster1.xml.rels': relationships([
      ['rId1', 'theme', '../theme/theme1.xml'],
    ]),
    'ppt/slideLayouts/slideLayout1.xml': layout,
    'ppt/slideLayouts/_rels/slideLayout1.xml.rels': relationships([
      ['rId1', 'slideMaster', '../slideMasters/slideMaster1.xml'],
    ]),
    'ppt/slides/slide1.xml': slide1,
    'ppt/slides/_rels/slide1.xml.rels': slideRelationships([
      ['rIdImage', 'image', '../media/image1.png'],
      ['rIdWmf', 'image', '../media/image2.wmf'],
      ['rIdLink', 'hyperlink', 'https://example.com/', true],
    ]),
    'ppt/slides/slide2.xml': slide2,
    'ppt/slides/_rels/slide2.xml.rels': slideRelationships([
      ['rIdImage', 'image', '../media/image1.png'],
    ]),
    'ppt/slides/slide3.xml': slide3,
    'ppt/slides/_rels/slide3.xml.rels': slideRelationships(),
    'ppt/media/image1.png': PNG_1X1,
    'ppt/media/image2.wmf': new Uint8Array([1, 2, 3]),
  });
}

// ---------------------------------------------------------------------------
// OpenDocument
// ---------------------------------------------------------------------------

const ODF_NS =
  'xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" ' +
  'xmlns:style="urn:oasis:names:tc:opendocument:xmlns:style:1.0" ' +
  'xmlns:text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" ' +
  'xmlns:table="urn:oasis:names:tc:opendocument:xmlns:table:1.0" ' +
  'xmlns:draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" ' +
  'xmlns:fo="urn:oasis:names:tc:opendocument:xmlns:xsl-fo-compatible:1.0" ' +
  'xmlns:xlink="http://www.w3.org/1999/xlink" ' +
  'xmlns:svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" ' +
  'xmlns:presentation="urn:oasis:names:tc:opendocument:xmlns:presentation:1.0"';

function odfPackage(mimetype: string, content: string, styles?: string): Uint8Array {
  const files: Record<string, string | Uint8Array> = {
    mimetype,
    'content.xml': content,
    'Pictures/image1.png': PNG_1X1,
    'Pictures/vector.wmf': new Uint8Array([1, 2, 3]),
  };
  if (styles) files['styles.xml'] = styles;
  return makeZip(files, { stored: ['mimetype'] });
}

export function odtFixture(): Uint8Array {
  const styles = `${XML}<office:document-styles ${ODF_NS}>
  <office:styles>
    <style:style style:name="Standard" style:family="paragraph"><style:text-properties fo:font-size="12pt"/></style:style>
    <style:style style:name="Heading" style:family="paragraph" style:parent-style-name="Standard"><style:text-properties fo:font-weight="bold"/></style:style>
  </office:styles>
  <office:automatic-styles><style:page-layout style:name="pm1"><style:page-layout-properties fo:page-width="21cm" fo:page-height="29.7cm" fo:margin-top="2cm" fo:margin-bottom="2cm" fo:margin-left="1in" fo:margin-right="72pt"/></style:page-layout></office:automatic-styles>
</office:document-styles>`;
  const content = `${XML}<office:document-content ${ODF_NS}>
  <office:automatic-styles>
    <style:style style:name="P1" style:family="paragraph" style:parent-style-name="Heading"><style:paragraph-properties fo:text-align="center" fo:margin-left="1cm" fo:text-indent="5mm"/><style:text-properties fo:color="#336699"/></style:style>
    <style:style style:name="P2" style:family="paragraph"><style:paragraph-properties fo:text-align="end"/></style:style>
    <style:style style:name="P3" style:family="paragraph"><style:paragraph-properties fo:text-align="justify"/></style:style>
    <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold" fo:font-style="italic" style:text-underline-style="solid" style:text-line-through-style="solid" fo:background-color="#ffff00" fo:font-size="14pt" style:text-position="super 58%"/></style:style>
    <style:style style:name="T2" style:family="text"><style:text-properties style:text-position="sub" fo:color="red; background: url(x)"/></style:style>
    <text:list-style style:name="L1"><text:list-level-style-number text:level="1" style:num-format="a"/><text:list-level-style-bullet text:level="2" text:bullet-char="-"/></text:list-style>
  </office:automatic-styles>
  <office:body><office:text>
    <text:tracked-changes><text:p>removed change</text:p></text:tracked-changes>
    <text:h text:outline-level="2" text:style-name="P1">Chapter</text:h>
    <text:p text:style-name="P2">Plain <text:span text:style-name="T1">styled</text:span><text:span text:style-name="T2">sub</text:span><text:s text:c="3"/>spaced<text:tab/>tabbed<text:line-break/>broken <text:a xlink:href="https://example.com/">link</text:a> <text:a xlink:href="javascript:alert(1)">unsafe</text:a><text:note><text:note-citation>1</text:note-citation><text:note-body><text:p>note body</text:p></text:note-body></text:note><text:page-number>5</text:page-number></text:p>
    <text:p text:style-name="P3"><draw:frame svg:width="2cm" svg:height="1cm"><draw:image xlink:href="Pictures/image1.png"/></draw:frame><draw:frame svg:width="2cm"><draw:image xlink:href="Pictures/vector.wmf"/></draw:frame><draw:frame><draw:image xlink:href="https://example.com/remote.png"/></draw:frame><draw:frame svg:width="3cm"><draw:text-box><text:p>boxed text</text:p></draw:text-box></draw:frame></text:p>
    <text:list text:style-name="L1">
      <text:list-item><text:p>one</text:p>
        <text:list><text:list-item><text:p>nested</text:p></text:list-item></text:list>
      </text:list-item>
      <text:list-header><text:p>two</text:p></text:list-header>
    </text:list>
    <text:list><text:list-item><text:p>plain bullet</text:p></text:list-item></text:list>
    <text:section><text:p>in section</text:p></text:section>
    <table:table>
      <table:table-header-rows><table:table-row><table:table-cell table:number-columns-spanned="2"><text:p>head</text:p></table:table-cell><table:covered-table-cell/></table:table-row></table:table-header-rows>
      <table:table-row><table:table-cell table:number-rows-spanned="2"><text:p>a</text:p></table:table-cell><table:table-cell><text:p>b</text:p></table:table-cell></table:table-row>
    </table:table>
  </office:text></office:body>
</office:document-content>`;
  return odfPackage('application/vnd.oasis.opendocument.text', content, styles);
}

export function odsFixture(): Uint8Array {
  const content = `${XML}<office:document-content ${ODF_NS}><office:body><office:spreadsheet>
  <table:table table:name="Data">
    <table:table-header-rows><table:table-row><table:table-cell office:value-type="string"><text:p>Name</text:p></table:table-cell><table:table-cell><text:p>Amount</text:p></table:table-cell></table:table-row></table:table-header-rows>
    <table:table-row><table:table-cell office:value-type="string"><text:p>Widget</text:p></table:table-cell><table:table-cell office:value-type="float" office:value="12.5"><text:p>12.50</text:p></table:table-cell><table:table-cell table:number-columns-repeated="1020"/></table:table-row>
    <table:table-row table:number-rows-repeated="2"><table:table-cell table:number-columns-repeated="2"><text:p>same</text:p></table:table-cell></table:table-row>
    <table:table-row><table:table-cell table:number-columns-spanned="2" table:number-rows-spanned="2"><text:p>merged<text:s/>cell</text:p><text:p>line two</text:p></table:table-cell><table:covered-table-cell/></table:table-row>
    <table:table-row table:number-rows-repeated="1048570"><table:table-cell table:number-columns-repeated="1024"/></table:table-row>
  </table:table>
  <table:table><table:table-row><table:table-cell/></table:table-row></table:table>
</office:spreadsheet></office:body></office:document-content>`;
  return odfPackage('application/vnd.oasis.opendocument.spreadsheet', content);
}

export function odpFixture(): Uint8Array {
  const styles = `${XML}<office:document-styles ${ODF_NS}><office:automatic-styles><style:page-layout style:name="PM1"><style:page-layout-properties fo:page-width="28cm" fo:page-height="15.75cm"/></style:page-layout></office:automatic-styles></office:document-styles>`;
  const content = `${XML}<office:document-content ${ODF_NS}>
  <office:automatic-styles>
    <style:style style:name="dp1" style:family="drawing-page"><style:drawing-page-properties draw:fill="solid" draw:fill-color="#102030"/></style:style>
    <style:style style:name="gr1" style:family="graphic"><style:graphic-properties draw:fill="solid" draw:fill-color="#ff0000" draw:stroke="solid" svg:stroke-color="#0000ff"/></style:style>
    <style:style style:name="gr2" style:family="graphic"><style:graphic-properties draw:fill="none" draw:stroke="none"/></style:style>
    <style:style style:name="T1" style:family="text"><style:text-properties fo:font-weight="bold"/></style:style>
  </office:automatic-styles>
  <office:body><office:presentation>
    <draw:page draw:name="page1" draw:style-name="dp1">
      <draw:frame presentation:class="title" svg:x="1cm" svg:y="1cm" svg:width="20cm" svg:height="3cm"><draw:text-box><text:p>Deck title</text:p></draw:text-box></draw:frame>
      <draw:frame presentation:class="subtitle" svg:x="1cm" svg:y="5cm" svg:width="20cm" svg:height="2cm"><draw:text-box><text:p>Sub <text:span text:style-name="T1">bold</text:span><text:s/>x<text:tab/>y<text:line-break/>z</text:p><text:list><text:list-item><text:p>bullet</text:p></text:list-item></text:list><text:p/></draw:text-box></draw:frame>
      <draw:frame svg:x="2cm" svg:y="8cm" svg:width="4cm" svg:height="4cm"><draw:image xlink:href="Pictures/image1.png"/></draw:frame>
      <draw:frame svg:x="2cm" svg:y="8cm" svg:width="4cm" svg:height="4cm"><draw:image xlink:href="Pictures/vector.wmf"/></draw:frame>
      <draw:rect draw:style-name="gr1" svg:x="10cm" svg:y="8cm" svg:width="3cm" svg:height="2cm"><text:p>In rect</text:p></draw:rect>
      <draw:ellipse draw:style-name="gr2" svg:x="14cm" svg:y="8cm" svg:width="3cm" svg:height="2cm"/>
      <draw:custom-shape svg:x="18cm" svg:y="8cm" svg:width="3cm"><text:p>Custom</text:p></draw:custom-shape>
      <draw:line svg:x1="1cm" svg:y1="14cm" svg:x2="10cm" svg:y2="14cm"/>
      <draw:connector svg:x1="12cm" svg:y1="10cm" svg:x2="12cm" svg:y2="14cm"/>
      <draw:g><draw:rect svg:x="20cm" svg:y="12cm" svg:width="1cm" svg:height="1cm"><text:p>Grouped</text:p></draw:rect></draw:g>
      <draw:frame presentation:class="notes" svg:x="0cm" svg:y="0cm" svg:width="1cm" svg:height="1cm"><draw:text-box><text:p>Speaker notes</text:p></draw:text-box></draw:frame>
    </draw:page>
    <draw:page draw:name="page2"><draw:frame svg:x="1cm" svg:y="1cm" svg:width="5cm" svg:height="2cm"><draw:text-box><text:p>Second page</text:p></draw:text-box></draw:frame></draw:page>
  </office:presentation></office:body>
</office:document-content>`;
  return odfPackage('application/vnd.oasis.opendocument.presentation', content, styles);
}

// ---------------------------------------------------------------------------
// EPUB
// ---------------------------------------------------------------------------

export function epubFixture(): Uint8Array {
  const container = `${XML}<container version="1.0" xmlns="urn:oasis:names:tc:opendocument:xmlns:container"><rootfiles><rootfile full-path="OEBPS/content.opf" media-type="application/oebps-package+xml"/></rootfiles></container>`;
  const opf = `${XML}<package xmlns="http://www.idpf.org/2007/opf" version="3.0"><metadata xmlns:dc="http://purl.org/dc/elements/1.1/"><dc:title>The Test Book</dc:title></metadata>
  <manifest>
    <item id="c1" href="text/chapter%201.xhtml" media-type="application/xhtml+xml"/>
    <item id="c2" href="text/chapter2.xhtml" media-type="application/xhtml+xml"/>
    <item id="gone" href="text/gone.xhtml" media-type="application/xhtml+xml"/>
    <item id="css" href="styles/book.css" media-type="text/css"/>
    <item id="img" href="images/cover.png" media-type="image/png"/>
    <item href="no-id.xhtml"/>
  </manifest>
  <spine><itemref idref="c1"/><itemref idref="c2"/><itemref idref="gone"/><itemref idref="unknown"/></spine>
</package>`;
  const chapter1 = `<?xml version="1.0" encoding="UTF-8"?><html xmlns="http://www.w3.org/1999/xhtml" xmlns:xlink="http://www.w3.org/1999/xlink"><head><title/>
  <link rel="stylesheet" href="../styles/book.css"/><link rel="stylesheet" href="../styles/missing.css"/><link rel="icon" href="x.ico"/>
  <meta http-equiv="refresh" content="0;url=https://evil.example"/><script src="evil.js"/><base href="https://evil.example/"/></head>
  <body onload="alert(1)"><h1 id="top">Chapter One</h1><a id="anchor"/>
  <p onclick="alert(2)">Hello <a href="chapter2.xhtml">next</a> <a href="https://example.com">out</a></p>
  <img src="../images/cover.png" srcset="x 2x" alt="cover"/><img src="../images/missing.png"/><img src="https://example.com/remote.png"/>
  <svg xmlns="http://www.w3.org/2000/svg"><image xlink:href="../images/cover.png"/><image href="../images/missing.png"/></svg>
  <iframe src="https://evil.example"/><form action="/steal"><input name="x"/></form><script>alert(3)</script></body></html>`;
  const chapter2 = `<html xmlns="http://www.w3.org/1999/xhtml"><body><p>Second chapter</p></body></html>`;
  return makeZip(
    {
      mimetype: 'application/epub+zip',
      'META-INF/container.xml': container,
      'OEBPS/content.opf': opf,
      'OEBPS/text/chapter 1.xhtml': chapter1,
      'OEBPS/text/chapter2.xhtml': chapter2,
      'OEBPS/styles/book.css': '@import url("evil.css"); h1 { color: navy }',
      'OEBPS/images/cover.png': PNG_1X1,
    },
    { stored: ['mimetype'] },
  );
}

// ---------------------------------------------------------------------------
// RTF, email, notebook
// ---------------------------------------------------------------------------

const PNG_HEX = Array.from(PNG_1X1, (byte) => byte.toString(16).padStart(2, '0')).join('');

export const RTF_FIXTURE = String.raw`{\rtf1\ansi\ansicpg1252\deff0\uc1
{\fonttbl{\f0\fswiss Arial;}}
{\colortbl;\red255\green0\blue0;\red0\green0\blue255;}
{\stylesheet{\s0 Normal;}}
{\info{\title Hidden title}{\author Nobody}}
{\*\generator Test;}
{\*\unknowndest ignore me}
\pard\qc\b\fs32 Heading\b0\fs24\par
\pard\li720\fi-360\sb120\sa240 Plain {\i italic} {\ul under\ulnone  none} {\strike struck} {\cf1 red} {\highlight2 marked} {\super up}{\sub down}{\nosupersub flat}\par
\pard\qr Caf\'e9 \u8364?uro \u-10179?\u-8704? \{braces\} \\ back\~nb\_hy\emdash\endash\lquote\rquote\ldblquote\rdblquote\bullet\tab tabbed\line broken\par
\pard\qj {\v hidden}{\plain\b bold after plain} justified\par
\pard\ql{\field{\*\fldinst HYPERLINK "https://example.com"}{\fldrslt link text}}\par
\trowd\cellx3000\cellx6000\pard\intbl Cell A\cell\pard\intbl Cell B\par second line\cell\row
\trowd\cellx3000\cellx6000\pard\intbl Cell C\cell\pard\intbl Cell D\cell\row
\pard After table\par
{\pict\pngblip\picwgoal1500\pichgoal750 ${PNG_HEX}}
{\pict\wmetafile8 010203}
{\*\shppict{\pict\jpegblip ffd8ffe0}}{\nonshppict{\pict\pngblip 00}}
\pard\sect Section\
Line after backslash-newline\bin3 abcDone
}`;

export const EML_FIXTURE = [
  'From: =?UTF-8?B?SsO8cmdlbg==?= <jurgen@example.com>',
  'To: =?iso-8859-1?Q?Andr=E9_Du_Pont?= <andre@example.com>,',
  ' second@example.com',
  'Cc: copy@example.com',
  'Subject: =?UTF-8?Q?Quarterly_report?= =?UTF-8?Q?_=E2=9C=93?=',
  'Date: Mon, 1 Jan 2024 10:00:00 +0000',
  'MIME-Version: 1.0',
  'Content-Type: multipart/mixed; boundary="outer"',
  '',
  'This is the preamble.',
  '--outer',
  'Content-Type: multipart/alternative; boundary="inner"',
  '',
  '--inner',
  'Content-Type: text/plain; charset="iso-8859-1"',
  'Content-Transfer-Encoding: quoted-printable',
  '',
  'Plain caf=E9 body with a soft =',
  'line break.',
  '--inner',
  'Content-Type: text/html; charset=utf-8',
  'Content-Transfer-Encoding: base64',
  '',
  btoa('<p>HTML <b>body</b><script>alert(1)</script></p>'),
  '--inner--',
  '--outer',
  'Content-Type: text/csv; name="data.csv"',
  'Content-Disposition: attachment; filename="data.csv"',
  'Content-Transfer-Encoding: base64',
  '',
  btoa('a,b\n1,2'),
  '--outer',
  'Content-Type: application/octet-stream',
  "Content-Disposition: attachment; filename*=UTF-8''r%C3%A9sum%C3%A9.txt",
  '',
  'resume text',
  '--outer',
  'Content-Type: image/png',
  'Content-Transfer-Encoding: base64',
  '',
  btoa(String.fromCharCode(...PNG_1X1)),
  '--outer',
  'Content-Type: message/rfc822',
  '',
  'Subject: nested',
  '',
  'inner message',
  '--outer',
  'Content-Type: text/plain; name*0="long-"; name*1="name.txt"',
  '',
  'named part',
  '--outer--',
  'This is the epilogue.',
].join('\r\n');

export const NOTEBOOK_FIXTURE = JSON.stringify({
  nbformat: 4,
  nbformat_minor: 5,
  metadata: { kernelspec: { language: 'python' }, language_info: { name: 'python' } },
  cells: [
    { cell_type: 'markdown', source: ['# Analysis\n', '\n', 'Some *notes*.'] },
    {
      cell_type: 'code',
      source: 'import math\nprint(math.pi)',
      outputs: [
        { output_type: 'stream', name: 'stdout', text: ['3.14159\n'] },
        { output_type: 'stream', name: 'stderr', text: '\u001b[31mwarning\u001b[0m' },
        {
          output_type: 'execute_result',
          data: { 'text/plain': ['42'], 'text/html': '<script>alert(1)</script><b>42</b>' },
        },
        {
          output_type: 'display_data',
          data: {
            'image/png':
              'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg==\n',
            'text/plain': '<Figure>',
          },
        },
        { output_type: 'display_data', data: { 'text/markdown': '**rendered**' } },
        {
          output_type: 'error',
          ename: 'ValueError',
          evalue: 'bad',
          traceback: ['\u001b[0;31mTraceback\u001b[0m', 'line 2'],
        },
        { output_type: 'error', ename: 'KeyError', evalue: 'k' },
        { output_type: 'display_data', data: 'not an object' },
        { output_type: 'display_data', data: {} },
        'not an output',
      ],
    },
    { cell_type: 'raw', source: 'raw cell text' },
    { cell_type: 'raw', source: '' },
    'not a cell',
  ],
});
