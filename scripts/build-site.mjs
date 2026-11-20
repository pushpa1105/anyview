// Assembles the documentation site with its playground into dist-site/, ready
// for GitHub Pages or any static host.
//
//   npm run site
//
// It needs the built package (npm run build) and the sample files that the
// test suite generates (npm test).
import { cpSync, existsSync, mkdirSync, readdirSync, rmSync, writeFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { fileURLToPath } from 'node:url';

const root = resolve(fileURLToPath(import.meta.url), '../..');
const out = join(root, 'dist-site');
const dist = join(root, 'dist');
const generated = join(root, 'tests/.generated');

// Office and binary samples come from the test fixtures: published name, source name.
const GENERATED_SAMPLES = [
  ['report.docx', 'real.docx'],
  ['budget.xlsx', 'real.xlsx'],
  ['slides.pptx', 'real.pptx'],
  ['document.pdf', 'sample.pdf'],
  ['book.epub', 'sample.epub'],
  ['archive.zip', 'sample.zip'],
];

if (!existsSync(join(dist, 'index.js'))) {
  console.error('dist/ is missing. Run "npm run build" first.');
  process.exit(1);
}
const missing = GENERATED_SAMPLES.filter(([, source]) => !existsSync(join(generated, source)));
if (missing.length > 0) {
  console.error('Sample files are missing. Run "npm test" once to generate them.');
  process.exit(1);
}

rmSync(out, { recursive: true, force: true });
cpSync(join(root, 'site'), out, { recursive: true });

// The browser build of the package: the ES modules and the chunks they load on demand.
mkdirSync(join(out, 'lib'), { recursive: true });
let modules = 0;
for (const file of readdirSync(dist)) {
  if (!file.endsWith('.js')) continue;
  cpSync(join(dist, file), join(out, 'lib', file));
  modules++;
}

for (const [name, source] of GENERATED_SAMPLES) {
  cpSync(join(generated, source), join(out, 'samples', name));
}

// GitHub Pages must serve the files as they are, without Jekyll.
writeFileSync(join(out, '.nojekyll'), '');

console.log(
  `Site written to dist-site/ (${modules} modules, ${readdirSync(join(out, 'samples')).length} samples).`,
);
