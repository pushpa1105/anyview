import { describe, expect, it } from 'vitest';
import { highlightToFragment, isHighlightable, tokenize } from '../../src/util/highlight';
import { renderMarkdown } from '../../src/util/markdown';

const types = (code: string, language: string): string[] =>
  tokenize(code, language)
    .filter((token) => token.type !== 'plain')
    .map((token) => `${token.type}:${token.text}`);

/** Renders Markdown into a detached element so it can be queried. */
function md(source: string, baseUrl?: string): HTMLElement {
  const host = document.createElement('div');
  host.appendChild(renderMarkdown(source, { baseUrl: baseUrl ?? null }));
  return host;
}

describe('tokenize', () => {
  it('never loses or duplicates characters', () => {
    const samples: [string, string][] = [
      ['const a = `tpl ${x}`; // c\n/* b */ f(1.5e3, "s\\"q")', 'javascript'],
      ['{"a": [1, true, null], "b": "x"}', 'json'],
      ['def f(x):\n    """doc"""\n    return x # c', 'python'],
      ['<a href="x" b=\'y\'><!-- c --><![CDATA[z]]></a>&amp;', 'xml'],
      ['key: value # c\n- item: 1', 'yaml'],
      ["SELECT * FROM t WHERE a = 'x' -- c", 'sql'],
      ['+added\n-removed\n@@ hunk @@\n context', 'diff'],
      ['plain text', 'unknown-language'],
      ['"unterminated', 'javascript'],
      ['/* unterminated', 'javascript'],
    ];
    for (const [code, language] of samples) {
      expect(
        tokenize(code, language)
          .map((token) => token.text)
          .join(''),
      ).toBe(code);
    }
  });

  it('classifies JavaScript tokens', () => {
    expect(types('const x = fn(42, "s", true); // note', 'javascript')).toEqual([
      'keyword:const',
      'function:fn',
      'number:42',
      'string:"s"',
      'literal:true',
      'comment:// note',
    ]);
  });

  it('treats TypeScript and other aliases like their base language', () => {
    expect(types('interface A {}', 'typescript')).toEqual(['keyword:interface']);
    expect(types('fn main() {}', 'rust')).toContain('keyword:fn');
    expect(types('echo "hi" # c', 'bash')).toEqual(['keyword:echo', 'string:"hi"', 'comment:# c']);
  });

  it('marks JSON property names separately from string values', () => {
    expect(types('{"key": "value", "n": 1.5}', 'json')).toEqual([
      'attr:"key"',
      'string:"value"',
      'attr:"n"',
      'number:1.5',
    ]);
  });

  it('handles multi-line strings, triple quotes and escapes', () => {
    expect(types('`a\nb`', 'javascript')).toEqual(['string:`a\nb`']);
    expect(types('"""a\nb"""', 'python')).toEqual(['string:"""a\nb"""']);
    expect(types('"a\\"b" c', 'javascript')).toEqual(['string:"a\\"b"']);
    // A single-line string stops at the end of the line.
    expect(types('"open\nnext', 'javascript')).toEqual(['string:"open']);
  });

  it('recognises hexadecimal, binary and suffixed numbers', () => {
    expect(types('0xFF 0b101 1_000 12px 3.5e-2', 'javascript')).toEqual([
      'number:0xFF',
      'number:0b101',
      'number:1_000',
      'number:12px',
      'number:3.5e-2',
    ]);
  });

  it('is case-insensitive for SQL keywords', () => {
    expect(types('select a FROM t', 'sql')).toEqual(['keyword:select', 'keyword:FROM']);
  });

  it('highlights keys in YAML, TOML and INI', () => {
    expect(types('name: app\n  port: 80', 'yaml')).toEqual(['attr:name', 'attr:port', 'number:80']);
    expect(types('title = "x"', 'toml')).toEqual(['attr:title', 'string:"x"']);
    expect(types('; note\nkey=value', 'ini')).toEqual(['comment:; note', 'attr:key']);
  });

  it('tokenizes markup tags, attributes and comments', () => {
    expect(types('<a href="x">t</a><!-- c -->', 'html')).toEqual([
      'tag:<a',
      'attr:href',
      'string:"x"',
      'tag:>',
      'tag:</a',
      'tag:>',
      'comment:<!-- c -->',
    ]);
    expect(types('<br/>', 'xml')).toEqual(['tag:<br', 'tag:/>']);
  });

  it('classifies diff lines', () => {
    expect(tokenize('+a\n-b\n@@ x\n c', 'diff').map((token) => token.type)).toEqual([
      'tag',
      'keyword',
      'function',
      'plain',
    ]);
  });

  it('returns no tokens for empty unknown input', () => {
    expect(tokenize('', null)).toEqual([]);
    expect(tokenize('x', undefined)).toEqual([{ type: 'plain', text: 'x' }]);
  });

  it('reports which languages it knows', () => {
    expect(isHighlightable('typescript')).toBe(true);
    expect(isHighlightable('JSON')).toBe(true);
    expect(isHighlightable('klingon')).toBe(false);
    expect(isHighlightable(null)).toBe(false);
  });
});

describe('highlightToFragment', () => {
  it('builds spans with token classes and plain text nodes', () => {
    const host = document.createElement('pre');
    host.appendChild(highlightToFragment('let a = 1;', 'javascript'));
    expect(host.textContent).toBe('let a = 1;');
    expect(host.querySelector('.av-tok-keyword')?.textContent).toBe('let');
    expect(host.querySelector('.av-tok-number')?.textContent).toBe('1');
  });

  it('never interprets code as markup', () => {
    const host = document.createElement('pre');
    host.appendChild(highlightToFragment('<img src=x onerror=alert(1)>', 'html'));
    expect(host.querySelector('img')).toBeNull();
    expect(host.textContent).toBe('<img src=x onerror=alert(1)>');
  });
});

describe('renderMarkdown: blocks', () => {
  it('renders ATX and setext headings', () => {
    const host = md('# One\n\n## Two ##\n\nThree\n=====\n\nFour\n----\n\n####### not a heading');
    expect(host.querySelector('h1')?.textContent).toBe('One');
    expect(host.querySelectorAll('h2')[0]?.textContent).toBe('Two');
    expect(host.querySelectorAll('h1')[1]?.textContent).toBe('Three');
    expect(host.querySelectorAll('h2')[1]?.textContent).toBe('Four');
    expect(host.querySelector('p')?.textContent).toBe('####### not a heading');
  });

  it('joins paragraph lines and separates paragraphs on blank lines', () => {
    const host = md('line one\nline two\n\nsecond');
    const paragraphs = host.querySelectorAll('p');
    expect(paragraphs).toHaveLength(2);
    expect(paragraphs[0]!.textContent).toBe('line one\nline two');
  });

  it('renders fenced code with a language and keeps its content literal', () => {
    const host = md('```js\nconst a = "<b>";\n```\n\n~~~\nplain *text*\n~~~');
    const blocks = host.querySelectorAll('pre > code');
    expect(blocks[0]!.getAttribute('data-language')).toBe('js');
    expect(blocks[0]!.textContent).toBe('const a = "<b>";');
    expect(blocks[0]!.querySelector('.av-tok-keyword')).not.toBeNull();
    expect(blocks[1]!.textContent).toBe('plain *text*');
    expect(host.querySelector('em')).toBeNull();
  });

  it('closes an unterminated fence at the end of the document', () => {
    expect(md('```\ncode').querySelector('pre')?.textContent).toBe('code');
  });

  it('renders indented code blocks', () => {
    const host = md('text\n\n    indented\n      more\n\nafter');
    expect(host.querySelector('pre')?.textContent).toBe('indented\n  more');
  });

  it('renders horizontal rules', () => {
    expect(md('a\n\n---\n\n***\n\n_ _ _').querySelectorAll('hr')).toHaveLength(3);
  });

  it('renders nested block quotes with lazy continuation', () => {
    const host = md('> quote\ncontinued\n> > nested\n\nafter');
    const quote = host.querySelector('blockquote')!;
    expect(quote.querySelector('p')?.textContent).toBe('quote\ncontinued');
    expect(quote.querySelector('blockquote')?.textContent).toBe('nested');
    expect(host.lastElementChild?.textContent).toBe('after');
  });

  it('renders unordered, ordered and nested lists', () => {
    const host = md('- a\n- b\n  - b1\n  - b2\n- c\n\n3. three\n4. four');
    const top = host.querySelector('ul')!;
    expect(top.children).toHaveLength(3);
    expect(top.children[1]!.querySelector('ul')?.children).toHaveLength(2);
    const ordered = host.querySelector('ol')!;
    expect(ordered.getAttribute('start')).toBe('3');
    expect(ordered.children).toHaveLength(2);
  });

  it('starts a new list when the bullet character changes', () => {
    expect(md('- a\n* b').querySelectorAll('ul')).toHaveLength(2);
  });

  it('supports multi-paragraph list items and lazy continuation', () => {
    const host = md('1. first\n\n   more\n2. second\nlazy');
    const items = host.querySelectorAll('ol > li');
    expect(items).toHaveLength(2);
    expect(items[0]!.querySelectorAll('p')).toHaveLength(2);
    expect(items[1]!.textContent).toBe('second\nlazy');
  });

  it('renders task lists as disabled checkboxes', () => {
    const boxes =
      md('- [x] done\n- [ ] todo').querySelectorAll<HTMLInputElement>('li.av-task input');
    expect(boxes).toHaveLength(2);
    expect(boxes[0]!.hasAttribute('checked')).toBe(true);
    expect(boxes[1]!.hasAttribute('checked')).toBe(false);
    expect(boxes[0]!.hasAttribute('disabled')).toBe(true);
  });

  it('renders GFM tables with alignment', () => {
    const host = md('| a | b | c |\n|:--|:-:|--:|\n| 1 | 2 | 3 |\n| x \\| y | z |');
    const headers = host.querySelectorAll('th');
    expect(Array.from(headers, (cell) => cell.textContent)).toEqual(['a', 'b', 'c']);
    expect(headers[1]!.style.textAlign).toBe('center');
    expect(headers[2]!.style.textAlign).toBe('right');
    const rows = host.querySelectorAll('tbody tr');
    expect(rows).toHaveLength(2);
    expect(rows[1]!.children[0]!.textContent).toBe('x | y');
    // Short rows are padded to the header width.
    expect(rows[1]!.children).toHaveLength(3);
  });

  it('does not treat a pipe line without a divider as a table', () => {
    expect(md('a | b\nc | d').querySelector('table')).toBeNull();
  });

  it('normalises CRLF and strips a byte-order mark', () => {
    expect(md('﻿# Title\r\n\r\ntext').querySelector('h1')?.textContent).toBe('Title');
  });
});

describe('renderMarkdown: inline', () => {
  it('renders emphasis, strong, strikethrough and code', () => {
    const host = md('*em* _em2_ **strong** __strong2__ ~~del~~ `code`');
    expect(Array.from(host.querySelectorAll('em'), (node) => node.textContent)).toEqual([
      'em',
      'em2',
    ]);
    expect(Array.from(host.querySelectorAll('strong'), (node) => node.textContent)).toEqual([
      'strong',
      'strong2',
    ]);
    expect(host.querySelector('del')?.textContent).toBe('del');
    expect(host.querySelector('code')?.textContent).toBe('code');
  });

  it('nests emphasis', () => {
    const host = md('**bold and *italic* text**');
    expect(host.querySelector('strong > em')?.textContent).toBe('italic');
  });

  it('leaves snake_case, lone markers and spaced asterisks alone', () => {
    expect(md('snake_case_name and 2 * 3 * 4 and a ~ b').querySelector('em')).toBeNull();
    expect(md('unclosed *emphasis').textContent).toBe('unclosed *emphasis');
  });

  it('honours backslash escapes and entities', () => {
    expect(md('\\*not em\\* &amp; &copy; &#65;').textContent).toBe('*not em* & &copy; A');
  });

  it('handles code spans with backticks inside', () => {
    expect(md('`` a ` b ``').querySelector('code')?.textContent).toBe('a ` b');
    expect(md('lone ` tick').querySelector('code')).toBeNull();
  });

  it('renders links that open in a new tab without opener access', () => {
    const link = md('[site](https://example.com "Title")').querySelector('a')!;
    expect(link.getAttribute('href')).toBe('https://example.com');
    expect(link.title).toBe('Title');
    expect(link.target).toBe('_blank');
    expect(link.rel).toContain('noopener');
  });

  it('supports angle-bracket URLs, nested parentheses and inline formatting in link text', () => {
    expect(md('[a](<https://example.com/a b>)').querySelector('a')?.getAttribute('href')).toBe(
      'https://example.com/a b',
    );
    expect(md('[a](https://example.com/x_(y))').querySelector('a')?.getAttribute('href')).toBe(
      'https://example.com/x_(y)',
    );
    expect(md('[**bold**](https://example.com)').querySelector('a > strong')).not.toBeNull();
  });

  it('resolves reference-style links', () => {
    const host = md(
      '[one][ref] [two][] [Two]\n\n[ref]: https://example.com/1\n[two]: <https://example.com/2> "T"',
    );
    const links = host.querySelectorAll('a');
    expect(Array.from(links, (link) => link.getAttribute('href'))).toEqual([
      'https://example.com/1',
      'https://example.com/2',
      'https://example.com/2',
    ]);
    expect(host.textContent).not.toContain('[ref]:');
  });

  it('renders text for links with unsafe or missing targets', () => {
    const host = md('[click](javascript:alert(1)) and [nothing] and [broken](');
    expect(host.querySelector('a')).toBeNull();
    expect(host.textContent).toContain('click');
    expect(host.textContent).toContain('[nothing]');
  });

  it('renders images with safe sources only', () => {
    const image = md('![alt *text*](https://example.com/a.png "T")').querySelector('img')!;
    expect(image.getAttribute('src')).toBe('https://example.com/a.png');
    expect(image.alt).toBe('alt text');
    expect(image.title).toBe('T');
    const unsafe = md('![x](javascript:alert(1))');
    expect(unsafe.querySelector('img')).toBeNull();
    expect(unsafe.textContent).toBe('x');
  });

  it('resolves relative links and images against the base URL', () => {
    const host = md('[doc](other.md) ![pic](img/a.png)', 'https://example.com/docs/readme.md');
    expect(host.querySelector('a')?.getAttribute('href')).toBe('https://example.com/docs/other.md');
    expect(host.querySelector('img')?.getAttribute('src')).toBe(
      'https://example.com/docs/img/a.png',
    );
  });

  it('autolinks angle-bracket and bare URLs', () => {
    const host = md('<https://a.example> <mailto:me@a.example> see https://b.example/path. Done');
    const links = host.querySelectorAll('a');
    expect(Array.from(links, (link) => link.getAttribute('href'))).toEqual([
      'https://a.example',
      'mailto:me@a.example',
      'https://b.example/path',
    ]);
    expect(links[1]!.textContent).toBe('me@a.example');
  });

  it('creates hard line breaks', () => {
    expect(md('a  \nb').querySelectorAll('br')).toHaveLength(1);
    expect(md('a\\\nb').querySelectorAll('br')).toHaveLength(1);
    expect(md('a<br>b<br/>c').querySelectorAll('br')).toHaveLength(2);
  });

  it('shows raw HTML as text instead of interpreting it', () => {
    const host = md(
      '<script>alert(1)</script>\n\n<img src=x onerror=alert(1)>\n\ninline <b onclick="x()">tag</b>',
    );
    expect(host.querySelector('script')).toBeNull();
    expect(host.querySelector('img')).toBeNull();
    expect(host.querySelector('b')).toBeNull();
    expect(host.textContent).toContain('<script>alert(1)</script>');
  });
});
