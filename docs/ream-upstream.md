# The layout engine: reamkit upstream and our changes

`src/ream/` is the reamkit engine, brought in as document-kit's own source. Its public entry
is `src/ream/index.ts` (`document-kit/ream`), upstream's own `src/index.ts` unchanged.
`tests/ream/` is upstream's test suite with the fixtures it reads. Licence: MIT,
Alexandr Krassavin (`NOTICE-REAMKIT.md`).

## Where it came from

- reamkit **1.35.0** (`package.json` version; `CHANGELOG.md` heads with 1.35.0), the `main`
  branch of <https://github.com/alex-krassavin/reamkit> as a source archive (no Git history),
  unpacked at `~/Downloads/document/reamkit-main` on the machine that brought it in.
- Taken: `src/**` → `src/ream/**`, `tests/**` → `tests/ream/**` (fixtures and snapshot
  included).
- Left out: `corpus/`, `stand/`, `docs/`, `scripts/` (corpus, publishing and table builders),
  the ESLint, Prettier, publint and `@tanstack/config` set-up, and upstream's `vite.config.ts`
  (its library build is replaced by document-kit's own, see below).

## Changes from upstream

Each change is marked in the source with a `document-kit:` comment unless noted.

1. **Zip timestamp in local time** (`src/ream/core/opc/opc-writer.ts`, `FIXED_MTIME`).
   Upstream built the fixed ZIP time with `Date.UTC(1980, 0, 1)`. fflate writes DOS time from
   the *local* clock fields, so west of Greenwich that instant is 1979-12-31, before DOS time
   begins, and fflate threw: 136 upstream tests failed in Pacific time. Now
   `new Date(1980, 0, 1).getTime()`: 1980-01-01 00:00 in every zone, the same bytes everywhere.
   (Commented in place, without the `document-kit:` prefix.)
2. **A whole complex field in one run** (`src/ream/word/document-parser.ts`, `fieldRunParts`).
   The `docx` npm library writes "Page {PAGE} of {NUMPAGES}" as ONE `w:r` holding
   `fldChar begin`, `instrText`, `fldChar separate`, `fldChar end` and the text around them.
   Word reads that; upstream read the run as one piece and the page numbers never printed. The
   run is now split into one run per part (each with the run's properties), as Word itself
   writes a field, before the field state machine sees it. Pinned by
   `tests/layout/layout.test.ts` ("Page N of M" on every page) and the browser case.
   (Commented in place, without the `document-kit:` prefix.)
3. **Alias imports made relative** (all of `src/ream` and `tests/ream`). Upstream imports its
   own modules as `@/…` through a path alias. An application consuming the built package has no
   such alias, and `tsc` does not rewrite aliases in the `.d.ts` files it emits, so every alias
   import is rewritten as a relative one by `scripts/rewrite-ream-aliases.mjs` (module
   specifiers only: `from`, `import`, `import()`, `vi.mock`; comments untouched). The same
   script points the tests' root-relative fixture paths (`'tests/fixtures/…'`) at
   `'tests/ream/fixtures/…'`. Specifiers stay extensionless, which bundler resolution (Vite,
   esbuild, webpack, TypeScript `moduleResolution: "bundler"`) reads; a consumer on
   `moduleResolution: "nodenext"` would need extensions, which upstream added in a build hook
   and we do not.
4. **No network by default** (`src/ream/core/fonts/remote-fonts.ts`, `NO_NETWORK`; doc comments
   in `core/converter/ream.ts`, `word/docx-to-pdf.ts`, `excel/xlsx-to-pdf.ts`,
   `core/fonts/provider.ts`). Upstream's font auto-download fell back to the global `fetch` and
   downloaded substitutes from jsDelivr. document-kit's rule is that nothing in `src/` reaches
   the network on its own, so the fallback is now `NO_NETWORK`, which refuses every request: a
   required face fails as it would offline (`Failed to download font from …`, with the refusal
   as its `cause`), a best-effort face is missing. A caller who injects a `fetch`
   (`fontFetch`, `remoteFontProvider({ fetch })`) still gets the upstream behaviour; that is the
   caller's choice, and `document-kit/layout` never does it. Upstream's own tests all inject
   their `fetch`, so none changed. `tests/layout/no-network.test.ts` pins it.
5. **`familyNamesInFlow`** (`src/ream/core/fonts/families.ts`). `familiesInFlow` is split in two:
   `familyNamesInFlow` returns the family names the text is drawn in, and `familiesInFlow` maps
   them to substitute keys exactly as before (same keys, same order). `document-kit/layout` uses
   the names to report which named family went to which stand-in.

Nothing else differs. `diff -r` against the upstream tree, after running the rewrite script
over a fresh copy, shows only the five changes above.

## Bringing in a newer upstream

1. Get the new release's source and diff its `src/` and `tests/` against the last one taken
   (1.35.0) to see what moved.
2. Copy its `src/**` over `src/ream/**` and its `tests/**` over `tests/ream/**` (fixtures
   included; skip `corpus/`, `stand/`, `docs/`, `scripts/`). Delete files upstream deleted.
3. Run `node scripts/rewrite-ream-aliases.mjs`, then `node scripts/rewrite-ream-aliases.mjs
   --check` (exits 1 if an alias import or root-relative fixture path is left).
4. Re-apply changes 1, 2, 4 and 5 above unless upstream has made them itself (check
   `FIXED_MTIME`, a `w:r` split for complex fields, every fallback to the global `fetch` in
   `core/fonts`, `core/converter` and `excel`, and `familyNamesInFlow`). Search for new uses of
   `fetch`, `XMLHttpRequest` and `node:` in `src/ream` and treat any as change 4.
5. Check `src/layout` still compiles against the internals it imports (`core/converter/project`,
   `core/fonts/families`, `core/fonts/remote-fonts`, `core/fonts/scripts`,
   `layout/styled-layout`, `pdf/styled-page-emitter`, `pdf/writer`, `svg/svg-writer`,
   `html/html-writer`, `word/docx-reader`, `core/crypto/*`).
6. `npx tsc -p tsconfig.json --noEmit`, then `npx vitest run tests/ream tests/layout` in the
   machine's own zone and once with `TZ=UTC` (and a zone east of Greenwich), then
   `npm run build && npm run test:browser`.
7. Update the version and date at the top of this file, and the change list.
