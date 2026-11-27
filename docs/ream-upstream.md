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

13. **A percentage width written with its sign** (`word/table-parser.ts`, `pctFraction`, marked
    `document-kit:`). §17.18.107 lets a `w:type="pct"` width be fiftieths of a percent (`4300`) or
    a percentage with its sign (`86%`); the `docx` library up to 9.7 writes the second. Read as a
    number, `86%` was nothing, so the table fell back to its grid, which that library writes as a
    100-twip stand-in: a placeholder box meant to fill most of the line was drawn five points wide,
    one word a line. `w:tblW` and `w:tcW` now read both forms. Pinned by
    `tests/layout/pictures-and-cells.test.ts` ("reads \"86%\" as 86 per cent …").

14. **A stand-in grid gives way to the cells' widths** (`layout/styled-layout.ts`, `gridIsStandIn`,
    marked `document-kit:`). The `docx` library, given no column widths, writes a 100-twip
    `w:gridCol` per column beside the table's width (`w:tblW`) and each cell's own (`w:tcW`); Word
    sizes such a table from the cells. Upstream trusts the grid of a table not laid out `fixed`,
    so the grid, scaled to the table's width, drew every column the same: a results table's
    parameter column a quarter of its width, every name wrapping. A grid less than half the width
    the table states is now taken as that stand-in and the cells' declared widths decide the
    columns, as shares of the width the table states (an autofit table's own width wins over its
    cells': a two-column table whose cells ask for 56% of a full-width table spans it); a table
    whose cells declare none keeps the grid. Pinned by
    `tests/layout/pictures-and-cells.test.ts` ("a table whose grid is a stand-in").

15. **A border held to the widths Word draws** (`word/table-parser.ts`, `borderEighths`;
    `word/paragraph-properties.ts`, `parseParagraphBorders`; both marked `document-kit:`).
    §17.3.4 gives a border's `w:sz` the range two to 96 eighths of a point, and Word draws a
    smaller one as two. The `docx` library writes a hairline as `w:sz="1"`; read as an eighth,
    every row of a table ruled that way stood an eighth of a point shorter than Word sets it, a
    line and a half over a forty-row page. A size of zero is kept. Pinned by
    `tests/layout/pictures-and-cells.test.ts` ("a border thinner than Word draws").

16. **Compact PDFs** (`layout/styled-layout.ts`, `StyledRenderOptions.compact` and
    `collectImageResources`; `pdf/styled-page-emitter.ts`, the page content stream and the
    PDF/A ICC profile; `core/images.ts`, `EmbedImageOptions.keepPalette`,
    `PreparedImage.palette`, `preparePng`, `splitChannels` and `knockOutColor`;
    `pdf/image-xobject.ts`, `addImage`; all marked `document-kit:`). Upstream writes each
    page's drawing as a plain stream and expands a palette PNG to RGB, so a long report's PDF
    was about five times the size it needs to be, most of it operators in plain text. With the
    new `compact` option the page content (§7.4.4) and the ICC profile are deflated, and a
    palette PNG is kept as one index a pixel under `[/Indexed /DeviceRGB hival <lookup>]`
    (§8.6.6.3), its `tRNS` still the soft mask; a palette image whose alpha PDF/A-1 flattens
    goes as RGB, as before. Off by default, so the engine's own output and its byte gate are
    unchanged; `src/layout` (`layoutDocx`, hence `docxToPdf`) turns it on. On a 51-page
    report the PDF went from 1.90 MB to 0.43 MB, and on 20 pages with 18 palette charts from
    1.42 MB to 0.78 MB, every page rendering the same. Pinned by `tests/ream/image.test.ts`
    ("kept as a palette") and `tests/layout/compact-pdf.test.ts`.

17. **Formulas kept on parsed cells** (`core/spreadsheet-model/types.ts`, `WorksheetCell.formula`
    and `WorksheetCell.sharedFormula`; `excel/worksheet-parser.ts`, `cellFormula`; both marked
    `document-kit:`). Upstream reads a cell's cached `<v>` and drops its `<f>`. A workbook
    written by a library that stores no cached values (openpyxl) then has an empty cell where
    every total stands. The formula text (without `=`) and a shared formula's group index are
    now kept on the cell, so `document-kit/docx`'s spreadsheet block can evaluate what the
    file did not cache. Nothing upstream reads the new fields; rendering is unchanged. Pinned by
    `tests/docx/spreadsheet.test.ts`.

18. **Bold and italic made for an embedded family that lacks them** (`layout/styled-layout.ts`,
    `runFontKeyAndParsed`, marked `document-kit:`). Upstream drew a bold run in a family the
    document embeds (or, through `document-kit/layout`'s `namedFonts`, the caller names) with
    whichever face the family has, so a family embedded with its regular face alone printed bold
    text regular. It now takes the same synthetic face the stand-ins take (`syntheticFace`):
    thickened for bold, slanted for italic, as Word draws a family it has no such face for. A
    family with its own bold face is drawn as before. Pinned by
    `tests/layout/named-fonts.test.ts`.
19. **OpenType faces with CFF outlines in the PDF** (`pdf/cid-font.ts`, `embedTtfFont` and the
    new `embedCffFont`, marked `document-kit:`). Upstream embedded every face as TrueType
    (`/FontFile2`, subset by `glyf`); a CFF face (an `.otf` brand face) has no `glyf`, and
    subsetting it threw. Such a face is now embedded whole as `/FontFile3` of subtype `/OpenType`
    under a `/CIDFontType0` descendant: its CFF is not CID-keyed (the layout refuses one that is,
    with a loss), so a CID is the glyph id and the page's codes are unchanged. Not subset: the
    PDF carries the whole face. Measuring, the page pictures (`svg/svg-text.ts` already drew CFF
    outlines) and `/ToUnicode` are unchanged. Pinned by `tests/layout/named-fonts.test.ts`.
20. **A band's pictures behind the text painted behind the page** (`layout/styled-layout.ts`,
    `drawBlocksSequentially`, marked `document-kit:`). Upstream drew a header's or footer's
    `behindDoc` picture first within the band but pushed it with the band's other items, so it
    rode the page's image pass, which paints after every fill: a full-page ground anchored in
    the header covered the body's cell and paragraph shading (§17.4.32, §17.3.1.31) in the PDF
    and the page pictures, while Word paints it under them (§20.4.2.3). Such a picture (or
    shape) is now marked `behind` with its `relativeHeight` as `z`, so the paint plan sets it
    down before everything else on the page, the band's in z order and ahead of the body's own
    behind-text floats. Pinned by `tests/layout/placed-blocks.test.ts` ("a page ground behind
    shading").
21. **A picture in a text frame stays in the frame** (`word/document-parser.ts`,
    `tryExtractDrawingFromParagraph`, marked `document-kit:`). Upstream collapses a paragraph
    that holds nothing but an inline picture into a standalone picture block, which drops the
    paragraph's `w:framePr` (§17.3.1.11): the picture fell out of the frame group and was drawn
    in the flow where the frame was anchored (the top of a cover page), and the frame's later
    paragraphs became a second frame laid over the first. A framed paragraph whose drawings
    are all inline is now kept as a paragraph, the picture an inline run in it, so the frame
    sets it at its own x and y with the rest of its paragraphs. A re-read keeps it the same way,
    so the round trip is unchanged. Pinned by `tests/layout/placed-blocks.test.ts` ("sets a
    picture in a frame inside the frame's box").
22. **Kerning behind an extension lookup, and the legacy `kern` table** (`core/font/opentype-layout.ts`,
    `parseGposKerning` and the new `parseKernTable`; `core/font/ttf-parser.ts`, `parseTtf`; marked
    `document-kit:`). Upstream read GPOS pair adjustments (LookupType 2) only where the `kern`
    feature names them directly; a face that reaches them through an Extension Positioning
    subtable (LookupType 9, format 1, a 32-bit offset to the wrapped subtable) had no kerning
    at all. Carlito, the Calibri stand-in, is such a face, and so is many a brand face. Such a
    lookup is now followed to its pair subtable. A face whose GPOS kerns nothing is read for
    its legacy OpenType `kern` table (version 0, horizontal format 0 subtables). Under the
    engine's default profile (kerned measuring) a Carlito run is now kerned where it was not;
    no fixture of the byte gate is set in such a face. Pinned by `tests/layout/kerning.test.ts`.
23. **Kerning as the run asks** (`layout/styled-layout.ts`, `StyledRenderOptions.kernAsAsked`,
    `kernsByDefault`, `runKerns`, `withKerning`, `kernedFaces`, in `tokenizeParagraph` and the
    table natural-width measure; `svg/svg-text.ts`, `SvgGlyphs` keyed by face; marked
    `document-kit:`). Upstream read `w:kern` (§17.3.2.19) into `kerningMinPt` and never used
    it: every run was measured kerned under the default profile and kern-free under `'word'`.
    With the new `kernAsAsked` option a run is kerned only when it states `w:kern` and its size
    is at least the size named (0 turns it off), as Word kerns; a run that states none is set
    at its glyphs' own widths. A kerned run takes a view of its face's resource (the same name,
    face and glyph subset, a kerned measure), so it is measured, broken and drawn (`TJ`
    adjustments in the PDF, glyph advances in the page pictures) kerned. The SVG writer's glyph
    outlines are keyed by the face rather than the resource, so a view shares its face's
    outlines. Off by default, so the engine's own output and the byte gate are unchanged;
    `src/layout` (`layoutDocx`, hence `docxToPdf` and `docxToPages`) turns it on. Pinned by
    `tests/layout/kerning.test.ts`.
24. **A run's character scale read from a .docx** (`word/run-properties.ts`, marked
    `document-kit:`). Upstream laid out, drew and wrote `widthScale` (§17.3.2.43 `w:w`) but
    never read it, so a run Word sets at 98% of its width came out at full width and broke its
    lines early. `w:w` is now read into `widthScale` (its percent over 100, kept within 1 to
    600; 100 kept too, as it undoes a style's share). Pinned by `tests/ream/run-properties.test.ts`,
    `tests/ream/run-width-scale.test.ts` and `tests/layout/size-scale.test.ts`.
25. **Text state after a picture or a drawn box** (`pdf/styled-page-emitter.ts`, marked
    `document-kit:`). The emitter set its record of the text state back to the defaults (`Tc` 0,
    `Tz` 100) after a picture, a math box or a rule, on the reasoning that ET resets them. It does
    not: ET ends the text object, and `Tc` and `Tz` stand until set again (ISO 32000-1 §9.3.1).
    A plain run after a picture therefore skipped its own `0 Tc` and `100 Tz` and was drawn with
    the tracking and scale of the run before it, wider than it was laid out. The record is now
    unknown there (NaN), so the next run states both. The byte gate's `xlsx: tracks` case changed
    with it (a sheet's tracked text after a rule now states its spacing). Pinned by
    `tests/layout/text-state-after-pictures.test.ts`.

Nothing else differs. `diff -r` against the upstream tree, after running the rewrite script
over a fresh copy, shows only the twenty-five changes above.

## Bringing in a newer upstream

1. Get the new release's source and diff its `src/` and `tests/` against the last one taken
   (1.35.0) to see what moved.
2. Copy its `src/**` over `src/ream/**` and its `tests/**` over `tests/ream/**` (fixtures
   included; skip `corpus/`, `stand/`, `docs/`, `scripts/`). Delete files upstream deleted.
3. Run `node scripts/rewrite-ream-aliases.mjs`, then `node scripts/rewrite-ream-aliases.mjs
   --check` (exits 1 if an alias import or root-relative fixture path is left).
4. Re-apply changes 1, 2 and 4 to 25 above unless upstream has made them itself (check
   `FIXED_MTIME`, a `w:r` split for complex fields, every fallback to the global `fetch` in
   `core/fonts`, `core/converter` and `excel`, `familyNamesInFlow`, `w:vAlign`, PAGEREF and
   body fields, §17.4.66 border conflicts, `w:keepNext` in table rows, and the two foot-frame
   checks of changes 10 and 11, the no-break hyphen of change 12, percentage widths with their sign of change 13, the stand-in grid of change 14, border widths of change 15, the `compact` option of change 16, synthetic faces for embedded families of change 18, CFF faces in the PDF of change 19, band pictures behind the page of change 20, framed pictures of change 21, extension and legacy kerning of change 22, `kernAsAsked` of change 23, `w:w` read of change 24, the text state after a picture of change 25). Search for
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
