import js from '@eslint/js';
import tseslint from 'typescript-eslint';

const UNSAFE_HTML_MESSAGE =
  'Never inject markup strings: build DOM with createElement/textContent, or use a sandboxed iframe for untrusted HTML.';

export default tseslint.config(
  {
    ignores: [
      'dist',
      'dist-*',
      'coverage',
      'node_modules',
      'examples',
      'scripts',
      'tests/.generated',
    ],
  },
  js.configs.recommended,
  ...tseslint.configs.recommendedTypeChecked,
  {
    languageOptions: {
      parserOptions: { projectService: true, tsconfigRootDir: import.meta.dirname },
    },
    rules: {
      '@typescript-eslint/consistent-type-imports': 'error',
      '@typescript-eslint/no-unused-vars': ['error', { argsIgnorePattern: '^_' }],
      // Security rule for the whole package: file contents are untrusted input.
      'no-restricted-syntax': [
        'error',
        {
          selector:
            "AssignmentExpression[left.type='MemberExpression'][left.property.name=/^(innerHTML|outerHTML)$/]",
          message: UNSAFE_HTML_MESSAGE,
        },
        {
          selector: "CallExpression[callee.property.name='insertAdjacentHTML']",
          message: UNSAFE_HTML_MESSAGE,
        },
        {
          // document.write / contentDocument.write, but not stream writers.
          selector:
            'CallExpression[callee.property.name=/^(write|writeln)$/]:matches([callee.object.name=/[dD]ocument$/], [callee.object.property.name=/[dD]ocument$/])',
          message: UNSAFE_HTML_MESSAGE,
        },
      ],
    },
  },
  {
    // Tests pass mocks and destructured helpers around freely.
    files: ['tests/**/*.ts'],
    rules: { '@typescript-eslint/unbound-method': 'off' },
  },
  {
    files: ['*.config.js', '*.config.ts'],
    ...tseslint.configs.disableTypeChecked,
  },
);
