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

6. **A section's vertical alignment** (`core/document-model/types.ts`
   `SectionProperties.verticalAlign`; `word/document-parser.ts` `parseSectPrNode`;
   `word/docx-writer.ts` `sectPrXml`; `layout/styled-layout.ts`
   `PageAssembler.alignPageVertically`, `lowerItem`). Upstream read no `w:vAlign` (§17.6.23),
   so a cover page centred in Word was set at the top of the sheet. `center` and `bottom` now
   lower everything the flow placed on each of the section's pages by half, or all, of the room
   left under it; bookmark destinations move with their text. The header, footer, notes band
   and floating drawings stay where they are (a picture anchored to a paragraph does not move
   with it). `both` (spread over the page) is read but laid out as `top`. A page shared by two
   sections, a page of several columns and a section whose lines run down the sheet are left
   as laid out. Pinned by `tests/layout/vertical-align.test.ts`.
7. **Page-number fields in the body** (`layout/body-fields.ts`, new; `layout/styled-layout.ts`
   `layoutStyledDocument`, now a loop over the upstream body renamed `layoutStyledOnce`, and
   `MAX_FIELD_PASSES`; `word/document-parser.ts` `parseFieldInstr`, `synthesizeFieldRun`,
   `applyFieldFsm`; `Run.field` gains `'PAGEREF'` and `Run.fieldTarget`;
   `LaidOutPage.pageLabel`; `substitutePageFields` leaves PAGEREF alone). Upstream printed the
   cached text of every field in the body, so a contents list the `docx` library wrote (whose
   `PAGEREF _TocN \h` fields cache nothing) had no page numbers, and PAGE or NUMPAGES in body
   text printed whatever Word last cached. Now PAGEREF (§17.16.5.45) is read with its
   bookmark (a `\p` PAGEREF, "above/below", keeps its cached text), and a field opened inside
   another's result keeps its own opening run's properties and link. A document with such a
   field in its body is laid out with the cached text, the numbers read from the pages (each in
   its page's numbering format, `pageLabel`), and laid out again with them until a layout prints
   the numbers it finds, at most four times; if it does not settle, the last layout is kept and
   a `degraded` `text` loss says the numbers may be a page out. A PAGE field is placed by a
   bookmark the pass gives its paragraph (removed from the result's bookmarks); a PAGE field in
   a table cell, and a PAGEREF to a bookmark the layout does not place (in a cell, a header,
   a text box), keep their cached text. A PAGEREF in a header or footer is filled the same
   way. A document with no such field is laid out once, exactly as before. Pinned by
   `tests/layout/fields.test.ts`.
8. **Table border conflicts** (`core/border-conflict.ts`, new; `layout/styled-layout.ts`
   `resolveCellBorders`, `heavierBorder`; `layout/page-doc.ts` `paintPlan`;
   `svg/svg-writer.ts`). Upstream weighed two facing cell borders by width alone, kept the
   lower or right cell's border on a tie, and gave a neighbour that states no `w:tcBorders` no
   claim at all. Now §17.4.66's rules decide: a cell's own border before the table's; a side
   that is none gives way; then width in eighths times the style's number; then the style
   earlier in the list; then the darker colour (R + B + 2G, then B + 2G, then G). A neighbour
   with no borders of its own claims the table's inside rule. The paint plan now paints
   borders lightest first (a stable sort by the same weight), so a heavy border is not cut at
   its corners by the projecting ends of the thin rules that run into it; and the SVG writer
   strokes a solid rule with a square cap as the PDF does (ISO 32000-1 §8.4.3.3), so a boxed
   cell's corners close in the page pictures too. Output changes only where a page has rules of
   different weights (the byte gate's fixtures are unchanged). Pinned by
   `tests/layout/table-borders.test.ts`.
9. **Rows kept with the next** (`layout/styled-layout.ts` `rowKeepsWithNext`,
   `keptRowsHeight`, `RowLayout.keepNext`, the table pass in `paginateSections`,
   `keptTogetherHeight`). Upstream ignored `w:keepNext` inside a table. A row now stays on the
   page of the row after it when every paragraph of its first cell keeps with the next (style
   included), which is how Word reads it: a run of kept rows that will not fit in the room left
   but will fit a fresh page, under the repeated header rows, moves to the next page whole. A
   run taller than a page breaks as before, and a table's last row keeps with nothing. A
   paragraph kept with a table (a caption) now counts the rows kept with the first, when they
   fit a page. Pinned by `tests/layout/keep-rows.test.ts`.
10. **A line a float pushes past the text area** (`layout/styled-layout.ts`, the line loop in
    the paragraph pass, marked `document-kit:`). A float that leaves no room beside a line
    pushes it under itself; upstream checked the page's foot before that push, so a frame
    standing at the foot of the page (a `w:framePr` anchored to the bottom margin, the model's
    `pinned` block) pushed the next line into the bottom margin, drawn there. The pushed place
    is now checked too, and a line that would end below the text area starts the next page, as
    Word sets it. Pinned by `tests/layout/report-features.test.ts` ("pushes text that would run
    into it on to the next page").
11. **A frame at the foot of a full page** (`layout/styled-layout.ts`, the `shape` float branch
    in the block pass, marked `document-kit:`). A text frame aligned to the bottom of the page or
    margin (`w:yAlign="bottom"`, `w:vAnchor` page or margin, wrapping) whose top stands above
    text already set on its page was drawn over that text. It now starts the next page and
    stands at its foot, where its anchor then falls. Frames placed by offset, anchored to the
    paragraph, aligned to the top or centre, or not wrapping are placed as before. Pinned by
    `tests/layout/report-features.test.ts` ("goes to the foot of the next page …").

12. **A no-break hyphen drawn and read as a hyphen** (`core/font/ttf-parser.ts`, `parseTtf`'s
    `glyphForCodepoint`; `layout/styled-layout.ts`, `countSharedGlyphs`; both marked
    `document-kit:`). Word's `w:noBreakHyphen` is read as U+2011, which keeps a word from breaking
    there. A face is now asked for U+2011 as its hyphen-minus (U+002D) glyph, falling back to its
    own U+2011 only when it has no hyphen, as Word draws it. Tinos has no U+2011, so a phone number
    kept on one line ("Phone: 250-555-0100") printed its hyphens in Arimo's shape inside a Times
    line; a face with a U+2011 glyph of its own wrote the PDF's text back as U+2011, so a number
    copied out of the PDF did not dial. The shared-glyph count that names a glyph in `/ToUnicode`
    skips U+2011, so the hyphen glyph is always named U+002D. (`src/layout`'s `pageText` and
    `pageLines` read the character as `-` too.) Pinned by `tests/layout/report-features.test.ts`
    ("reads and prints the unbreakable hyphens as the hyphen a person types").

Nothing else differs. `diff -r` against the upstream tree, after running the rewrite script
over a fresh copy, shows only the twelve changes above.

## Bringing in a newer upstream

1. Get the new release's source and diff its `src/` and `tests/` against the last one taken
   (1.35.0) to see what moved.
2. Copy its `src/**` over `src/ream/**` and its `tests/**` over `tests/ream/**` (fixtures
   included; skip `corpus/`, `stand/`, `docs/`, `scripts/`). Delete files upstream deleted.
3. Run `node scripts/rewrite-ream-aliases.mjs`, then `node scripts/rewrite-ream-aliases.mjs
   --check` (exits 1 if an alias import or root-relative fixture path is left).
4. Re-apply changes 1, 2 and 4 to 12 above unless upstream has made them itself (check
   `FIXED_MTIME`, a `w:r` split for complex fields, every fallback to the global `fetch` in
   `core/fonts`, `core/converter` and `excel`, `familyNamesInFlow`, `w:vAlign`, PAGEREF and
   body fields, §17.4.66 border conflicts, `w:keepNext` in table rows, and the two foot-frame
   checks of changes 10 and 11, the no-break hyphen of change 12). Search for
   `document-kit:` in the old tree to find each place. Search for new uses of
   `fetch`, `XMLHttpRequest` and `node:` in `src/ream` and treat any as change 4.
5. Check `src/layout` still compiles against the internals it imports (`core/converter/project`,
   `core/font/ttf-build` (`buildTrueType`, for `src/layout/tickBoxes.ts`),
   `core/fonts/families`, `core/fonts/remote-fonts`, `core/fonts/scripts` (`scriptsInFlow`,
   `companionsInFlow`, `forEachTextInFlow`, `scriptForCodepoint`), `core/font` (`ParsedTtf`),
   `layout/styled-layout`, `pdf/styled-page-emitter`, `pdf/writer`, `svg/svg-writer`,
   `html/html-writer`, `word/docx-reader`, `core/crypto/*`).
6. `npx tsc -p tsconfig.json --noEmit`, then `npx vitest run tests/ream tests/layout` in the
   machine's own zone and once with `TZ=UTC` (and a zone east of Greenwich), then
   `npm run build && npm run test:browser`.
7. Update the version and date at the top of this file, and the change list.
