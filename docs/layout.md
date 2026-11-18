# document-kit/layout

A Word file (`.docx`) to a PDF, to one SVG picture per page, or to flowed HTML, in the browser,
with fonts the caller hands in. It is a small front over the reamkit engine (`src/ream`,
`document-kit/ream`): the engine's reader, layout and writers do the work.

## API

```ts
import { createFontProvider, docxToPdf, docxToPages, docxToHtml, layoutDocx } from 'document-kit/layout';

const fonts = createFontProvider(async ({ file }) =>
  new Uint8Array(await (await fetch(`${base}/fonts/${file}`)).arrayBuffer()),
);

const { bytes, pageCount, losses } = await docxToPdf(docx, { fonts });
const { pages, pageSizes, pageText } = await docxToPages(docx, { fonts });
const { html } = await docxToHtml(docx);

// Preview and export from one layout:
const laid = await layoutDocx(docx, { fonts });
const preview = laid.toPages();      // same as docxToPages
const pdf = laid.toPdf({ info });    // same as docxToPdf, without laying out again
```

- `layoutDocx(docx, { fonts?, password?, pdfA?, tagged? })` → `Promise<LaidOutDocx>`: the file
  laid out once. `LaidOutDocx` has `pageCount`, `pageSizes`, `pageText`, `pageLines`, `losses` (what reading
  and laying out recorded), and two writers that reuse the layout, each callable any number of
  times: `toPdf({ info? })` → `PdfResult` and `toPages({ sharedGlyphs? })` → `PagesResult`.
  `pdfA` and `tagged` are layout options because the accessible PDF/A levels and a tagged PDF
  need the structure tree the layout builds; they change nothing on the pages. `docxToPdf` and
  `docxToPages` are `layoutDocx` followed by one writer, and write the same bytes as before.

- `docxToPdf(docx, { fonts?, password?, info?, pdfA?, tagged? })` →
  `Promise<{ bytes, pageCount, losses }>`. `info` (title, author, …) is laid over the Word
  file's own properties; `pdfA` is a level such as `'PDF/A-2b'`; `tagged` writes a tagged PDF.
- `docxToPages(docx, { fonts?, password?, sharedGlyphs? })` →
  `Promise<{ pages, pageSizes, pageText, pageLines, losses, glyphDefs? }>`. The same layout the PDF gets.
  - `pages[i]` is a standalone SVG document for page *i*, drawn by the engine's SVG writer one
    page at a time. Text is drawn from the faces' own glyph outlines (`<path>` definitions and
    `<use>`), so a page shows the same in any viewer with no fonts installed, and there is no
    selectable or searchable text in it. Each page carries its own glyph definitions, so pages
    do not share anything and some outlines repeat between pages. Every id in a page is
    prefixed `dk-p<page>-`, so pages can be placed inline in one HTML document without taking
    each other's glyphs (the engine numbers them from zero on every page). The page sits on a white
    rectangle with a light grey outline (the engine's preview style). Size is in points:
    `width`/`height`/`viewBox` are the page in pt.
  - `sharedGlyphs: true` is for an application that puts every page inline in ONE HTML
    document: the glyph outlines are taken out of the pages and put once in `glyphDefs`, a
    hidden `<svg>` of `<defs>` (ids `dk-g1`, `dk-g2`, …) that the pages refer to. Place
    `glyphDefs` anywhere in the same document (it takes no room); a page without it draws no
    text. Each page keeps its own clips and pictures. An outline is known by its shape, so the
    same glyph on two pages is defined once. The saving is the outlines only: each drawn glyph
    is still one `<use>` on its page, which is most of a page's size. On the test sample it is
    about a sixth at two pages and a quarter at twelve. Off by default: each page stays a
    standalone picture.
  - `pageSizes[i]` is `{ widthPt, heightPt }`.
  - `pageText[i]` is the text of the page's laid-out lines, one per `\n`, in layout order (use it
    for search and checks). Text inside charts and drawn shapes is not included. A no-break hyphen
    (Word's `w:noBreakHyphen`) reads as `-`, as it does in the PDF's text.
  - `pageLines[i]` is the same lines with their places: `{ text, x, y }`, `x` where the line's
    first character starts and `y` its baseline, in points from the page's top-left corner,
    rounded to a hundredth. For checks that compare two layouts line by line.
- `docxToHtml(docx, { password? })` → `Promise<{ html, losses }>`. Flowed, self-contained HTML
  (pictures inlined). No layout and no fonts. Headers and footers have no fixed page in a flowed
  document and are reported as a loss, so page numbers do not appear.
- `createFontProvider(load)` → `FontProvider`. `load({ family, style, file })` returns the bytes
  of one bundled font file (`Uint8Array` or `ArrayBuffer`, sync or async), or `undefined`/`null`
  for a file not supplied. A family is loaded only when a document needs it, each family once; a
  loader that throws is reported as a loss and asked again next time.
- `standInFor(name)` → the bundled family that stands in for a font name (`'carlito'` for
  Calibri, …). `STAND_IN_FILES` lists the twenty bundled files; `STAND_IN_NAMES` the families'
  published names.
- `WrongPasswordError` is thrown for an encrypted Word file opened without its password or with
  the wrong one. A file that is not a `.docx` throws an `Error`.

Types exported: `FontProvider`, `FontLoader`, `FontFileBytes`, `StandInFamily`, `StandInFile`,
`FaceStyle`, `ProvidedFamily`, `LayoutOptions`, `PagedOptions`, `PdfOptions`, `PdfResult`,
`PageLine`, `PagesResult`, `PagesWriteOptions`, `PageSize`, `HtmlResult`, `LayoutLoss`, `LayoutDocxOptions`,
`LaidOutDocx`, `PdfWriteOptions`, `DocumentInfo`, `PdfALevel`.

## The font rule

The kit never reaches the network. For each family the document draws text in:

1. A font **embedded in the document** (Word's *Embed fonts*) is used as it is.
2. Otherwise the **caller's stand-in** for that name, chosen by the engine's own substitution
   table (`standInFor`): Calibri → Carlito, Cambria → Caladea, Arial/Helvetica → Arimo, Times New
   Roman → Tinos, Courier New → Cousine; other serif names → Tinos, other monospaced names →
   Cousine, anything else → Arimo. Arimo is always asked for: it draws text with no named font,
   math and chart labels.
3. A stand-in nobody supplied is **not fetched**. Its text is drawn with the first family that
   was supplied (Arimo when it was), and the loss says so.

Without a provider, only embedded fonts are used; a document that embeds none cannot be laid
out and `docxToPdf`/`docxToPages` reject with "No fonts…". `docxToHtml` needs no fonts.

The open files are in `fonts/` (`fonts/README.md` lists each family's licence and source) and
are exported as `document-kit/fonts/<file>`, for an application to serve with its own assets.

The engine underneath (`document-kit/ream`) can still download fonts, but only if a caller hands
it a `fetch` (`fontFetch`, `remoteFontProvider({ fetch })`); its default refuses every download
(`docs/ream-upstream.md`, change 4). `tests/layout/no-network.test.ts` fails if `fetch` is called
during a conversion, and the browser test fails on any request off its own server.

## Losses

Every result carries `losses`: an ordered list of the engine's `Loss` records,
`{ severity, feature, detail, where? }`.

- `severity`: `'substituted'` (replaced by an equivalent, e.g. a stand-in font),
  `'degraded'` (kept with less fidelity) or `'dropped'` (missing from the output).
- `feature`: what was affected, e.g. `fonts.substitution`, `text`, `headersFooters`.
- `detail`: a sentence naming what happened; `where`: a page, when the engine knows it.

The engine's own read and write losses come through unchanged (e.g. a character no face has a
glyph for is `dropped` `text`, named by code point, `where: 'page N'`). The kit adds font
losses of its own:

- `substituted`: `"calibri" drawn with its stand-in Carlito` — every named family not embedded
  (names are as the engine reads them, lower-cased). Expect one for every face a document
  names: no Office face is drawn as itself unless the document embeds it. A document that names
  no face (the `docx` library's default) has none.
- `substituted`: `"calibri" drawn with Arimo: Carlito Regular was not available (not supplied)`.
- `degraded`: `Carlito Bold was not available (…)` — bold or italic text is then drawn from
  another of the family's faces, thickened or slanted by the layout.
- `degraded`: Greek or Cyrillic in Cambria text wants Noto Serif, which is not bundled.
- `degraded`: the document has Japanese/Korean/Chinese/Arabic/Hebrew/Thai/Indic or symbol text
  that no supplied face draws, named by code point (`… symbol text that no supplied face draws
  (U+2610) …`). Only characters missing from every face supplied (and every embedded face) are
  named: Arimo draws ▼ ► ■ ↑, so a document with only those has no such loss. The missing ones
  come out as empty boxes (and a `dropped` `text` loss names them per page).
- `degraded` `text`: page numbers in body fields did not settle (see below).

Tick boxes ☐ ☑ ☒ (U+2610–2612), which a Word check box prints (the block model's `checkbox` run and `choices`
block), are drawn from a small face the kit builds from outlines (`src/layout/tickBoxes.ts`, through the engine's
TrueType builder): a square, ticked or crossed. It is added only when the document has such a character that no
supplied or embedded face draws, so they are no longer named as missing symbols; the characters stay in `pageText`.

## What the layout does with a Word file's page set-up and fields

- **Vertical alignment of a section** (`w:vAlign`, the model's `page.verticalAlign`): `center`
  sets the text of each of the section's pages in the middle of the room between the margins,
  `bottom` at its foot. The header, the footer and floating drawings stay where they are.
  `both` (spread over the page) is laid out as `top`. A page shared by two sections, a page of
  several columns and a section turned on its side are left at the top.
- **Page numbers in the body.** A contents list's `PAGEREF` fields print the page their heading
  lands on, in that page's numbering (`iv` on a page numbered in roman). `PAGE` in body text
  prints the page its paragraph starts on, `NUMPAGES` the page count. These numbers only
  pagination knows, so a document with such a field is laid out, the numbers read from the
  pages, and laid out again with them, until a layout prints what it finds: two layouts for an
  ordinary contents list (one if the file's cached numbers are already right), at most four.
  If filling the numbers keeps moving the text (a "9" becoming "10" pushing a line over), the
  fourth layout is kept, its numbers may be a page out, and a `degraded` loss says so. Not
  filled: a `PAGE` field in a table cell, a `PAGEREF` to a bookmark in a table cell, a text box
  or a header, and a `PAGEREF \p` ("above"/"below"); they print the file's cached text.
- **Table borders.** Where two cells share an edge, the border shown follows ECMA-376
  §17.4.66: a cell's own border before the table's, a side with none gives way, then the
  heavier (width × the style's number), then the earlier style, then the darker colour. Heavier
  rules are painted over lighter ones, so a thick border keeps its corners.
- **Frames at the foot of the page** (`w:framePr` with `w:yAlign="bottom"` against the page or margin, the model's
  `pinned` block): set at the foot of the page their anchor falls on, the text flowing above them. A line the frame
  would push below the text area starts the next page, and a frame that would cover text already set on its page
  (the page is too full for it under its text) stands at the foot of the next page instead
  (`docs/ream-upstream.md`, changes 10 and 11).
- **Column widths.** A table's grid (`w:tblGrid`) decides its columns, as Word keeps it, unless the
  table is laid out `fixed` (the cells' widths decide) or the grid is far narrower than the width
  the table states: the `docx` library, given no column widths, writes 100-twip stand-ins beside
  the cells' real widths, and those cells' widths then decide, as Word draws them
  (`docs/ream-upstream.md`, change 14). A percentage width may be written `4300` or `86%`
  (change 13).
- **Rows kept together.** A row whose first cell's paragraphs keep with the next stays on the
  page of the next row (the model's `keepWithNext`, on by default for header rows, and
  `keepTogether`); a run of such rows that fits a page moves to the next page whole rather than
  leave a header row at the foot of a page. A run taller than a page still breaks.

## Known gaps

- **Word files only.** The engine also reads `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt` and PDF;
  this API takes `.docx`. Use `document-kit/ream` for the rest (and pass fonts yourself).
- **No caller-named fonts.** A caller cannot yet supply their own bytes for a named family
  (say, a licensed Calibri); only the five stand-ins. The engine supports it (`embeddedFonts`);
  the API does not expose it yet.
- **No scripts beyond Latin, Greek and Cyrillic** in the bundled set (see above).
- **Each call lays out.** `docxToPdf` and `docxToPages` each run the layout; use
  `layoutDocx` to write both from one. A document with page-number fields in its body is laid
  out two to four times inside that one call (see above).
- **Bundle size.** The engine is large: `dist/layout.js` is small but imports a shared chunk of
  about 1.5 MB unminified (about 440 kB gzipped) that also holds readers this API does not use
  (the build cannot split them out because `document-kit/ream` shares it).
- **SVG pages have no text layer**; `pageText` stands in for search. Pages repeat glyph
  outlines unless `sharedGlyphs` is asked for.
- **Tick boxes only.** Of the symbols the bundled faces lack, only ☐ ☑ ☒ are drawn (see Losses);
  other symbols still print as empty boxes with a loss.
- **HTML has no headers, footers or page numbers.**
- **Fidelity is the engine's.** Stand-ins match widths, not shapes. Layout differences from
  Word in particular documents are the engine's to fix (and record in `docs/ream-upstream.md`).
