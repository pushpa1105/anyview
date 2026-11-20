import {
  docxFixture,
  EML_FIXTURE,
  epubFixture,
  NOTEBOOK_FIXTURE,
  odpFixture,
  odsFixture,
  odtFixture,
  pptxFixture,
  RTF_FIXTURE,
  xlsxFixture,
} from '../helpers/fixtures';
import { makeZip, PNG_1X1 } from '../helpers/zip';

/** A minimal but fully valid one-page PDF with correct cross-reference offsets. */
export function minimalPdf(): string {
  const stream = 'BT /F1 28 Tf 72 720 Td (Hello from a PDF) Tj ET';
  const objects = [
    '<< /Type /Catalog /Pages 2 0 R >>',
    '<< /Type /Pages /Kids [3 0 R] /Count 1 >>',
    '<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Contents 4 0 R /Resources << /Font << /F1 5 0 R >> >> >>',
    `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
    '<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>',
  ];
  let pdf = '%PDF-1.4\n';
  const offsets: number[] = [];
  objects.forEach((body, index) => {
    offsets.push(pdf.length);
    pdf += `${index + 1} 0 obj\n${body}\nendobj\n`;
  });
  const xref = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets) pdf += `${String(offset).padStart(10, '0')} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return pdf;
}

/** One second of silence as an 8 kHz mono WAV file. */
export function silentWav(): Uint8Array {
  const samples = 8000;
  const wav = new Uint8Array(44 + samples);
  const view = new DataView(wav.buffer);
  const tag = (offset: number, text: string): void => {
    for (let i = 0; i < text.length; i++) wav[offset + i] = text.charCodeAt(i);
  };
  tag(0, 'RIFF');
  view.setUint32(4, 36 + samples, true);
  tag(8, 'WAVEfmt ');
  view.setUint32(16, 16, true);
  view.setUint16(20, 1, true);
  view.setUint16(22, 1, true);
  view.setUint32(24, 8000, true);
  view.setUint32(28, 8000, true);
  view.setUint16(32, 1, true);
  view.setUint16(34, 8, true);
  tag(36, 'data');
  view.setUint32(40, samples, true);
  wav.fill(128, 44);
  return wav;
}

const MARKDOWN = [
  '# Markdown sample',
  '',
  'Some **bold**, *italic* and `code`.',
  '',
  '- [x] done',
  '- [ ] todo',
  '',
  '| a | b |',
  '|---|---|',
  '| 1 | 2 |',
  '',
  '```js',
  'const answer = 42;',
  '```',
  '',
].join('\n');

/** Sample files for the real-browser check (scripts/browser-check.mjs) and the examples. */
export function sampleFiles(): Record<string, string | Uint8Array> {
  return {
    'sample.xlsx': xlsxFixture(),
    'sample.docx': docxFixture(),
    'sample.pptx': pptxFixture(),
    'sample.odt': odtFixture(),
    'sample.ods': odsFixture(),
    'sample.odp': odpFixture(),
    'sample.epub': epubFixture(),
    'sample.rtf': RTF_FIXTURE,
    'sample.eml': EML_FIXTURE,
    'sample.ipynb': NOTEBOOK_FIXTURE,
    'sample.zip': makeZip({ 'docs/readme.md': '# Inside the archive', 'data.csv': 'a,b\n1,2' }),
    'sample.pdf': minimalPdf(),
    'sample.wav': silentWav(),
    'sample.png': PNG_1X1,
    // The script must never run: SVG is only ever shown through <img>.
    'sample.svg':
      '<svg xmlns="http://www.w3.org/2000/svg" width="200" height="120"><script>parent.postMessage("svg-script-ran","*")</script><rect width="200" height="120" fill="#0969da"/><text x="20" y="70" fill="#fff" font-size="28">SVG</text></svg>',
    'sample.md': MARKDOWN,
    'sample.csv': 'name,amount,date\nWidget,12.50,2024-01-31\n"Gadget, large",1200,2024-02-01\n',
    'sample.json':
      '{"name":"anyview","formats":["pdf","docx","xlsx"],"nested":{"ok":true,"count":3}}',
    'sample.xml':
      '<?xml version="1.0"?>\n<catalog>\n  <book id="1"><title>XML sample</title></book>\n</catalog>\n',
    'sample.ts':
      'export function greet(name: string): string {\n  return "Hello, " + name; // greeting\n}\n',
    'sample.txt': 'Plain text sample.\nSecond line.\n',
    // The script must never run: HTML is shown in a sandboxed frame.
    'sample.html':
      '<!DOCTYPE html><html><head><title>HTML sample</title></head><body><h1>HTML sample</h1><p>Rendered in a sandbox.</p><script>parent.postMessage("html-script-ran","*")</script></body></html>',
    'legacy.doc': new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xa1, 0xb1, 0x1a, 0xe1, 0, 0, 0, 0]),
  };
}
