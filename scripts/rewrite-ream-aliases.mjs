#!/usr/bin/env node
// Rewrite reamkit's `@/…` path-alias imports as relative imports.
//
// Upstream reamkit imports its own modules through a `@/*` → `src/*` alias.
// An application that consumes document-kit's built package has no such
// alias, and `tsc` does not rewrite aliases in the `.d.ts` files it emits, so
// every alias import in src/ream (and in its tests) is turned into a relative
// one here. Run it after bringing in a newer upstream (docs/ream-upstream.md):
//
//   node scripts/rewrite-ream-aliases.mjs            # src/ream and tests/ream
//   node scripts/rewrite-ream-aliases.mjs --check    # exit 1 if any alias is left
//
// Only module specifiers are touched — `from '@/…'`, `import '@/…'`,
// `import('@/…')` and `vi.mock('@/…')` — never prose in comments. The alias
// root is always src/ream, whichever tree the importing file is in. The script
// is idempotent: a second run changes nothing.
//
// Second rule, tests only: upstream's tests open fixtures by paths relative to
// the repository root (`'tests/fixtures/fonts/Roboto-Regular.ttf'`), and here
// they live one folder down. A string that starts with `tests/fixtures/` becomes
// `tests/ream/fixtures/`.

import { readFileSync, readdirSync, statSync, writeFileSync } from 'node:fs';
import { dirname, join, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';

const repo = resolve(dirname(fileURLToPath(import.meta.url)), '..');
const aliasRoot = join(repo, 'src', 'ream');
const roots = [join(repo, 'src', 'ream'), join(repo, 'tests', 'ream')];
const check = process.argv.includes('--check');

const SPECIFIER =
  /(\bfrom\s*|\bimport\s*\(\s*|\bimport\s+|\bvi\.mock\(\s*|\bvi\.importActual(?:<[^>]*>)?\(\s*)(['"])@\/([^'"]+)\2/g;

function* walk(dir) {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    const st = statSync(path);
    if (st.isDirectory()) yield* walk(path);
    else if (/\.(?:ts|mts|tsx)$/.test(name)) yield path;
  }
}

function relativeSpecifier(fromFile, target) {
  let rel = relative(dirname(fromFile), join(aliasRoot, target)).split(sep).join('/');
  if (rel === '') rel = '.';
  if (!rel.startsWith('.')) rel = `./${rel}`;
  // A module importing its own directory ('@/excel' from inside excel/) would
  // become '.' or '..': name the index file so every resolver agrees.
  if (rel === '.' || rel === '..' || rel.endsWith('/..')) rel = `${rel}/index`;
  return rel;
}

const FIXTURE_PATH = /(['"`])tests\/fixtures\//g;

let files = 0;
let rewrites = 0;
for (const root of roots) {
  for (const file of walk(root)) {
    const source = readFileSync(file, 'utf8');
    let count = 0;
    const out = source.replace(SPECIFIER, (_m, pre, quote, target) => {
      count += 1;
      return `${pre}${quote}${relativeSpecifier(file, target)}${quote}`;
    });
    const final = file.startsWith(join(repo, 'tests', 'ream'))
      ? out.replace(FIXTURE_PATH, (_m, quote) => {
          count += 1;
          return `${quote}tests/ream/fixtures/`;
        })
      : out;
    if (count === 0) continue;
    files += 1;
    rewrites += count;
    if (!check) writeFileSync(file, final);
    else console.error(`alias import or root-relative fixture path left in ${relative(repo, file)} (${count})`);
  }
}

if (check) {
  if (rewrites > 0) process.exit(1);
  console.log('no @/ alias imports or root-relative fixture paths left in src/ream or tests/ream');
} else {
  console.log(`rewrote ${rewrites} alias imports and fixture paths in ${files} files`);
}
