import { createViewer, FORMATS, SUPPORTED_FORMAT_IDS, VERSION } from './lib/index.js';

const SAMPLES = [
  ['Word', 'report.docx'],
  ['Excel', 'budget.xlsx'],
  ['PowerPoint', 'slides.pptx'],
  ['PDF', 'document.pdf'],
  ['Markdown', 'guide.md'],
  ['CSV', 'sales.csv'],
  ['Code', 'example.ts'],
  ['JSON', 'data.json'],
  ['Image', 'logo.svg'],
  ['E-book', 'book.epub'],
  ['Archive', 'archive.zip'],
];

const CATEGORIES = {
  document: 'Documents',
  spreadsheet: 'Spreadsheets',
  presentation: 'Presentations',
  image: 'Images',
  video: 'Video',
  audio: 'Audio',
  text: 'Text and code',
  archive: 'Archives',
  font: 'Fonts',
  other: 'Other',
};

const $ = (id) => document.getElementById(id);
const status = $('status');
const themeSelect = $('theme');
const toolbarToggle = $('toolbar');
const stage = $('stage');

let viewer;
/** What is open now, so it can be reopened when the viewer is recreated. */
let current = null;

function formatSize(bytes) {
  if (!Number.isFinite(bytes)) return '';
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} kB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function createPlaygroundViewer() {
  viewer?.destroy();
  viewer = createViewer('#viewer', { theme: themeSelect.value, toolbar: toolbarToggle.checked });
  viewer.on('loadstart', () => {
    status.textContent = 'Opening…';
  });
  viewer.on('load', ({ format, file }) => {
    const page = viewer.getPage();
    status.textContent = [
      file.name,
      format.label,
      formatSize(file.size),
      page ? `${page.pageCount} ${page.pageCount === 1 ? 'page' : 'pages'}` : '',
    ]
      .filter(Boolean)
      .join(' · ');
  });
  viewer.on('error', (error) => {
    status.textContent = `${error.code}: ${error.message}`;
  });
}

function open(source, name) {
  current = { source, name };
  // Load errors are shown by the viewer itself and reported through the error event.
  viewer.load(source, name ? { name } : undefined).catch(() => {});
}

function selectSample(button) {
  for (const other of $('samples').children) other.setAttribute('aria-pressed', 'false');
  button?.setAttribute('aria-pressed', 'true');
}

// Samples.
for (const [label, file] of SAMPLES) {
  const button = document.createElement('button');
  button.type = 'button';
  button.textContent = label;
  button.title = file;
  button.setAttribute('aria-pressed', 'false');
  button.addEventListener('click', () => {
    selectSample(button);
    open(`samples/${file}`);
  });
  $('samples').appendChild(button);
}

// Own files: picker and drag and drop.
$('open-file').addEventListener('click', () => $('file-input').click());
$('file-input').addEventListener('change', (event) => {
  const file = event.target.files?.[0];
  if (file) {
    selectSample(null);
    open(file);
  }
  event.target.value = '';
});
let dragDepth = 0;
const hasFiles = (event) => Array.from(event.dataTransfer?.types ?? []).includes('Files');
stage.addEventListener('dragenter', (event) => {
  if (!hasFiles(event)) return;
  event.preventDefault();
  dragDepth++;
  stage.classList.add('dragging');
});
stage.addEventListener('dragover', (event) => {
  if (hasFiles(event)) event.preventDefault();
});
stage.addEventListener('dragleave', () => {
  dragDepth = Math.max(0, dragDepth - 1);
  if (dragDepth === 0) stage.classList.remove('dragging');
});
stage.addEventListener('drop', (event) => {
  if (!hasFiles(event)) return;
  event.preventDefault();
  dragDepth = 0;
  stage.classList.remove('dragging');
  const file = event.dataTransfer.files[0];
  if (file) {
    selectSample(null);
    open(file);
  }
});

// Options.
themeSelect.addEventListener('change', () => viewer.setTheme(themeSelect.value));
toolbarToggle.addEventListener('change', () => {
  // The toolbar is chosen when a viewer is created, so build a new one.
  createPlaygroundViewer();
  if (current) open(current.source, current.name);
});

// Install command.
$('copy-install').addEventListener('click', async () => {
  const button = $('copy-install');
  try {
    await navigator.clipboard.writeText($('install-command').textContent);
    button.textContent = 'Copied';
  } catch {
    button.textContent = 'Select and copy';
  }
  setTimeout(() => (button.textContent = 'Copy'), 1500);
});

// Quick start tabs.
const tabs = Array.from(document.querySelectorAll('[role="tab"]'));
for (const tab of tabs) {
  tab.addEventListener('click', () => {
    for (const other of tabs) other.setAttribute('aria-selected', String(other === tab));
    for (const panel of document.querySelectorAll('[data-panel]')) {
      panel.hidden = panel.getAttribute('data-panel') !== tab.getAttribute('data-tab');
    }
  });
}

// Format list, straight from the package.
$('version').textContent = VERSION;
const groups = new Map();
for (const format of FORMATS) {
  if (!SUPPORTED_FORMAT_IDS.includes(format.id)) continue;
  if (!groups.has(format.category)) groups.set(format.category, []);
  groups.get(format.category).push(format);
}
for (const [category, title] of Object.entries(CATEGORIES)) {
  const formats = groups.get(category);
  if (!formats) continue;
  const block = document.createElement('div');
  const heading = document.createElement('h3');
  heading.textContent = title;
  const list = document.createElement('ul');
  for (const format of formats) {
    const item = document.createElement('li');
    const name = document.createElement('span');
    name.textContent = format.label;
    const extensions = document.createElement('code');
    extensions.textContent = format.extensions.map((extension) => `.${extension}`).join(' ');
    item.append(name, extensions);
    list.appendChild(item);
  }
  block.append(heading, list);
  $('format-list').appendChild(block);
}

createPlaygroundViewer();
const first = $('samples').firstElementChild;
selectSample(first);
open(`samples/${SAMPLES[0][1]}`);
