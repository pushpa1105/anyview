/**
 * A compact syntax highlighter. It is a lexer, not a parser: it recognises
 * comments, strings, numbers, keywords and call-like identifiers for a set of
 * common languages, which is what makes code readable at a glance.
 */

export type TokenType =
  'plain' | 'comment' | 'string' | 'number' | 'keyword' | 'literal' | 'function' | 'tag' | 'attr';

export interface Token {
  type: TokenType;
  text: string;
}

interface Language {
  keywords?: ReadonlySet<string>;
  literals?: ReadonlySet<string>;
  lineComments?: readonly string[];
  blockComments?: readonly (readonly [string, string])[];
  /** Quote characters that delimit single-line strings. */
  quotes?: string;
  /** Quote characters whose strings may span lines (template literals…). */
  multilineQuotes?: string;
  tripleQuotes?: boolean;
  caseInsensitive?: boolean;
  /** JSON: a string followed by a colon is a property name. */
  keysAreStrings?: boolean;
  /** `key: value` / `key = value` at line start (YAML, TOML, INI). */
  lineKeys?: RegExp;
}

const words = (list: string): ReadonlySet<string> => new Set(list.split(' '));
const C_LITERALS = words(
  'true false null undefined nullptr NULL nil None True False NaN Infinity this self super',
);
const C_COMMENTS = { lineComments: ['//'], blockComments: [['/*', '*/']] } as const;

const JS_KEYWORDS =
  'as async await break case catch class const continue debugger default delete do else enum export extends finally for from function get if implements import in instanceof interface let new of package private protected public return set static switch throw try typeof var void while with yield type namespace declare readonly abstract keyof infer is satisfies';

const LANGUAGES: Readonly<Record<string, Language>> = {
  javascript: {
    ...C_COMMENTS,
    keywords: words(JS_KEYWORDS),
    literals: C_LITERALS,
    quotes: '"\'',
    multilineQuotes: '`',
  },
  json: { ...C_COMMENTS, literals: words('true false null'), quotes: '"', keysAreStrings: true },
  python: {
    lineComments: ['#'],
    keywords: words(
      'and as assert async await break class continue def del elif else except finally for from global if import in is lambda nonlocal not or pass raise return try while with yield match case',
    ),
    literals: C_LITERALS,
    quotes: '"\'',
    tripleQuotes: true,
  },
  clike: {
    ...C_COMMENTS,
    keywords: words(
      'abstract auto bool boolean break byte case catch char class const constexpr continue default delete do double dynamic else enum explicit extends extern final finally float fn for foreach func fun go goto if impl implements import in inline instanceof int interface internal let long match mut namespace new operator override package private protected pub public register return short signed sizeof static struct super switch template throw throws trait try type typedef typename union unsigned use using val var virtual void volatile when where while defer chan map range select fallthrough func guard init protocol extension object companion data sealed suspend echo function elseif endif endwhile endfor require include',
    ),
    literals: C_LITERALS,
    quotes: '"\'',
    multilineQuotes: '`',
  },
  ruby: {
    lineComments: ['#'],
    keywords: words(
      'alias and begin break case class def defined? do else elsif end ensure for if in module next not or redo rescue retry return then undef unless until when while yield require include attr_accessor attr_reader puts',
    ),
    literals: C_LITERALS,
    quotes: '"\'',
  },
  shell: {
    lineComments: ['#'],
    keywords: words(
      'if then else elif fi for while until do done case esac in function return exit export local readonly declare set unset source alias echo cd',
    ),
    quotes: '"\'',
    multilineQuotes: '`',
  },
  powershell: {
    lineComments: ['#'],
    blockComments: [['<#', '#>']],
    keywords: words(
      'if else elseif switch foreach for while do until function param return break continue try catch finally throw begin process end in',
    ),
    quotes: '"\'',
    caseInsensitive: true,
  },
  batch: {
    lineComments: ['::', 'REM ', 'rem '],
    keywords: words(
      'if else for in do goto call echo set setlocal endlocal exit not exist defined pause',
    ),
    quotes: '"',
    caseInsensitive: true,
  },
  sql: {
    lineComments: ['--'],
    blockComments: [['/*', '*/']],
    keywords: words(
      'select from where and or not insert into values update set delete create table alter drop index view join inner left right outer full on as group by order having limit offset union all distinct case when then else end primary key foreign references constraint default unique check begin commit rollback with in is like between exists asc desc',
    ),
    literals: words('null true false'),
    quotes: '"\'',
    multilineQuotes: '`',
    caseInsensitive: true,
  },
  css: {
    blockComments: [['/*', '*/']],
    lineComments: ['//'],
    quotes: '"\'',
    lineKeys: /^[ \t]*(--[\w-]+|[a-z-]+)(?=\s*:[^:])/,
  },
  yaml: {
    lineComments: ['#'],
    literals: words('true false null yes no on off ~'),
    quotes: '"\'',
    lineKeys: /^[ \t]*(?:- +)?([\w.$/-]+)(?=\s*:(\s|$))/,
  },
  toml: {
    lineComments: ['#'],
    literals: words('true false'),
    quotes: '"\'',
    tripleQuotes: true,
    lineKeys: /^[ \t]*([\w.-]+)(?=\s*=)/,
  },
  ini: { lineComments: ['#', ';'], quotes: '"\'', lineKeys: /^[ \t]*([\w.-]+)(?=\s*[=:])/ },
  lua: {
    lineComments: ['--'],
    blockComments: [['--[[', ']]']],
    keywords: words(
      'and break do else elseif end for function goto if in local not or repeat return then until while',
    ),
    literals: words('true false nil'),
    quotes: '"\'',
  },
  tex: { lineComments: ['%'] },
  makefile: { lineComments: ['#'], quotes: '"\'', lineKeys: /^([\w.-]+)(?=\s*[:?+]?=)/ },
  dockerfile: {
    lineComments: ['#'],
    keywords: words(
      'FROM RUN CMD LABEL EXPOSE ENV ADD COPY ENTRYPOINT VOLUME USER WORKDIR ARG ONBUILD STOPSIGNAL HEALTHCHECK SHELL AS',
    ),
    quotes: '"\'',
  },
  graphql: {
    lineComments: ['#'],
    keywords: words(
      'query mutation subscription fragment on type interface union enum input scalar schema extend directive implements',
    ),
    literals: words('true false null'),
    quotes: '"',
    tripleQuotes: true,
  },
};

const ALIASES: Readonly<Record<string, string>> = {
  typescript: 'javascript',
  js: 'javascript',
  ts: 'javascript',
  jsx: 'javascript',
  tsx: 'javascript',
  py: 'python',
  java: 'clike',
  c: 'clike',
  cpp: 'clike',
  csharp: 'clike',
  cs: 'clike',
  go: 'clike',
  rust: 'clike',
  php: 'clike',
  swift: 'clike',
  kotlin: 'clike',
  scala: 'clike',
  dart: 'clike',
  r: 'python',
  perl: 'shell',
  bash: 'shell',
  sh: 'shell',
  zsh: 'shell',
  yml: 'yaml',
  rb: 'ruby',
  scss: 'css',
  less: 'css',
  xml: 'markup',
  html: 'markup',
  svg: 'markup',
  markup: 'markup',
  diff: 'diff',
};

/** True when `language` (or an alias of it) has highlighting rules. */
export function isHighlightable(language: string | null | undefined): boolean {
  if (!language) return false;
  const key = language.toLowerCase();
  return key in LANGUAGES || key in ALIASES;
}

const IDENTIFIER = /[A-Za-z_$][\w$]*/y;
const NUMBER =
  /(?:0[xX][\da-fA-F_]+|0[bB][01_]+|\d[\d_]*(?:\.\d[\d_]*)?(?:[eE][+-]?\d+)?)[a-zA-Z%]*/y;

function tokenizeMarkup(code: string): Token[] {
  const tokens: Token[] = [];
  const pattern = /<!--[\s\S]*?(?:-->|$)|<!\[CDATA\[[\s\S]*?(?:\]\]>|$)|<[!?/]?[A-Za-z][^<>]*>?/g;
  let last = 0;
  let match: RegExpExecArray | null;
  while ((match = pattern.exec(code)) !== null) {
    if (match.index > last) tokens.push({ type: 'plain', text: code.slice(last, match.index) });
    const text = match[0];
    if (text.startsWith('<!--')) {
      tokens.push({ type: 'comment', text });
    } else if (text.startsWith('<![CDATA[')) {
      tokens.push({ type: 'string', text });
    } else {
      // Split a tag into its name, attribute names and attribute values.
      const part = /^<[!?/]?[\w:.-]+|[\w:@.-]+(?==)|"[^"]*"?|'[^']*'?|[^\w"'<]+|\w+/g;
      let piece: RegExpExecArray | null;
      while ((piece = part.exec(text)) !== null) {
        const value = piece[0];
        const type: TokenType = value.startsWith('<')
          ? 'tag'
          : value.startsWith('"') || value.startsWith("'")
            ? 'string'
            : text[piece.index + value.length] === '='
              ? 'attr'
              : /^[/?]?>$/.test(value.trim())
                ? 'tag'
                : 'plain';
        tokens.push({ type, text: value });
      }
    }
    last = match.index + text.length;
  }
  if (last < code.length) tokens.push({ type: 'plain', text: code.slice(last) });
  return tokens;
}

function tokenizeDiff(code: string): Token[] {
  return code.split(/(?<=\n)/).map((line): Token => {
    const type: TokenType = line.startsWith('+')
      ? 'tag'
      : line.startsWith('-')
        ? 'keyword'
        : line.startsWith('@@')
          ? 'function'
          : 'plain';
    return { type, text: line };
  });
}

/** Splits source code into typed tokens. Unknown languages yield one plain token. */
export function tokenize(code: string, language: string | null | undefined): Token[] {
  const key = (language ?? '').toLowerCase();
  const name = key in LANGUAGES ? key : ALIASES[key];
  if (name === 'markup') return tokenizeMarkup(code);
  if (name === 'diff') return tokenizeDiff(code);
  const rules = name ? LANGUAGES[name] : undefined;
  if (!rules) return code ? [{ type: 'plain', text: code }] : [];

  const tokens: Token[] = [];
  const length = code.length;
  let plainStart = 0;
  let i = 0;
  let lineStart = true;

  const flush = (end: number): void => {
    if (end > plainStart) tokens.push({ type: 'plain', text: code.slice(plainStart, end) });
  };
  const push = (type: TokenType, end: number): void => {
    flush(i);
    tokens.push({ type, text: code.slice(i, end) });
    i = end;
    plainStart = end;
  };

  outer: while (i < length) {
    const char = code[i]!;

    if (lineStart && rules.lineKeys) {
      rules.lineKeys.lastIndex = 0;
      const lineEnd = code.indexOf('\n', i);
      const key = rules.lineKeys.exec(code.slice(i, lineEnd === -1 ? length : lineEnd));
      if (key?.[1]) {
        const keyStart = i + key[0].length - key[1].length;
        i = keyStart;
        push('attr', keyStart + key[1].length);
        lineStart = false;
        continue;
      }
    }
    lineStart = char === '\n';

    for (const [open, close] of rules.blockComments ?? []) {
      if (code.startsWith(open, i)) {
        const end = code.indexOf(close, i + open.length);
        push('comment', end === -1 ? length : end + close.length);
        continue outer;
      }
    }
    for (const marker of rules.lineComments ?? []) {
      if (code.startsWith(marker, i)) {
        const end = code.indexOf('\n', i);
        push('comment', end === -1 ? length : end);
        continue outer;
      }
    }

    const isQuote = rules.quotes?.includes(char) ?? false;
    const isMultiline = rules.multilineQuotes?.includes(char) ?? false;
    if (isQuote || isMultiline) {
      const triple = rules.tripleQuotes === true && code.startsWith(char.repeat(3), i);
      const delimiter = triple ? char.repeat(3) : char;
      let end = i + delimiter.length;
      for (; end < length; end++) {
        const current = code[end]!;
        if (current === '\\') end++;
        else if (code.startsWith(delimiter, end)) break;
        else if (current === '\n' && !triple && !isMultiline) {
          end--;
          break;
        }
      }
      const stop = Math.min(length, end + delimiter.length);
      let type: TokenType = 'string';
      if (rules.keysAreStrings && /^\s*:/.test(code.slice(stop, stop + 16))) type = 'attr';
      push(type, stop);
      continue;
    }

    if (char >= '0' && char <= '9') {
      NUMBER.lastIndex = i;
      const match = NUMBER.exec(code);
      if (match) {
        push('number', i + match[0].length);
        continue;
      }
    }

    IDENTIFIER.lastIndex = i;
    const identifier = IDENTIFIER.exec(code);
    if (identifier) {
      const word = identifier[0];
      const end = i + word.length;
      const lookup = rules.caseInsensitive ? word.toLowerCase() : word;
      if (rules.keywords?.has(lookup)) push('keyword', end);
      else if (rules.literals?.has(lookup)) push('literal', end);
      else if (code[end] === '(') push('function', end);
      else i = end;
      continue;
    }
    i++;
  }
  flush(length);
  return tokens;
}

/** Renders code as text nodes and `<span class="av-tok-…">` elements. */
export function highlightToFragment(
  code: string,
  language: string | null | undefined,
): DocumentFragment {
  const fragment = document.createDocumentFragment();
  for (const token of tokenize(code, language)) {
    if (token.type === 'plain') {
      fragment.appendChild(document.createTextNode(token.text));
    } else {
      const span = document.createElement('span');
      span.className = `av-tok-${token.type}`;
      span.textContent = token.text;
      fragment.appendChild(span);
    }
  }
  return fragment;
}
