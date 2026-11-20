// Removes build and coverage output.
import { rmSync } from 'node:fs';

for (const directory of ['dist', 'coverage']) {
  rmSync(directory, { recursive: true, force: true });
}
