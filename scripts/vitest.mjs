// Runs Vitest from a path with an upper-case Windows drive letter.
//
// Some editors (VS Code among them) open terminals in `c:\...` rather than
// `C:\...`. Vitest then gets loaded through two differently-cased paths, ends
// up with two copies of its runtime, and every suite fails with
// "Cannot read properties of undefined (reading 'config')".
// On other platforms this script is a plain pass-through.
import { resolve } from 'node:path';
import { pathToFileURL } from 'node:url';

const normaliseDrive = (path) => path.replace(/^[a-z]:/, (drive) => drive.toUpperCase());

const root = normaliseDrive(process.cwd());
process.chdir(root);

await import(pathToFileURL(resolve(root, 'node_modules/vitest/vitest.mjs')).href);
