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

// A brand's own faces, by family name, drawn before anything else (TrueType or OpenType/CFF):
const branded = await layoutDocx(docx, {
  fonts,
  namedFonts: [
    { family: 'Brand Sans', bytes: regular },
    { family: 'Brand Sans', weight: 700, bytes: bold },
  ],
});
const preview = laid.toPages();      // same as docxToPages
const pdf = laid.toPdf({ info });    // same as docxToPdf, without laying out again
```

- `layoutDocx(docx, { fonts?, namedFonts?, password?, pdfA?, tagged? })` → `Promise<LaidOutDocx>`: the file
  laid out once. `LaidOutDocx` has `pageCount`, `pageSizes`, `pageText`, `pageLines`, `losses` (what reading
  and laying out recorded), and two writers that reuse the layout, each callable any number of
  times: `toPdf({ info? })` → `PdfResult` and `toPages({ sharedGlyphs? })` → `PagesResult`.
  `pdfA` and `tagged` are layout options because the accessible PDF/A levels and a tagged PDF
  need the structure tree the layout builds; they change nothing on the pages. `docxToPdf` and
  `docxToPages` are `layoutDocx` followed by one writer, and write the same bytes as before.

- `namedFonts` (every paged call takes it: `layoutDocx`, `docxToPdf`, `docxToPages`) is a list of `NamedFont`
  (`{ family, bytes, weight?, style? }`, the type `document-kit` exports and `writeDocx`'s `embedFonts` takes): the
  caller's own faces by family name, used for every run naming the family (the name matched without regard to
  case), **before** a face the document embeds and before the bundled stand-ins. A weight of 600 or more is the
  family's bold face, `style: 'italic'` its italic; a family given no bold or italic face has bold and italic text
  thickened or slanted by the layout, as Word does. TrueType and OpenType (CFF) faces both work, in Node and in the
  browser: a CFF face is drawn in the page pictures from its own outlines and embedded in the PDF whole, as
  `/FontFile3` (`/OpenType`). A face the layout cannot draw with (a CID-keyed CFF font, a collection, a file that
  is not a font) is left out with a `substituted` loss naming it, and its family falls back to the document's own
  face or a stand-in.
- `docxToPdf(docx, { fonts?, password?, info?, pdfA?, tagged? })` →
  `Promise<{ bytes, pageCount, pageSizes, losses }>`. `info` (title, author, …) is laid over the Word
  file's own properties; `pdfA` is a level such as `'PDF/A-2b'`; `tagged` writes a tagged PDF.
  The PDF is written compact: each page's drawing is deflated, and a palette (8-bit indexed) PNG
  stays a palette under an `/Indexed` colour space rather than being expanded to RGB. The pages are
  the same; a long report is about a fifth of the size (`docs/ream-upstream.md`, change 16).
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

Types exported: `NamedFont`, `FontProvider`, `FontLoader`, `FontFileBytes`, `StandInFamily`, `StandInFile`,
`FaceStyle`, `ProvidedFamily`, `LayoutOptions`, `PagedOptions`, `PdfOptions`, `PdfResult`,
`PageLine`, `PagesResult`, `PagesWriteOptions`, `PageSize`, `HtmlResult`, `LayoutLoss`, `LayoutDocxOptions`,
`LaidOutDocx`, `PdfWriteOptions`, `DocumentInfo`, `PdfALevel`.

## Previewing the generated PDF

Use `createPdfSource` when preview and download should use the same generated file. The generator
can run in a worker. Requests for the same immutable snapshot and variant share a pending or
completed result. Failed builds can be retried. Pass a new snapshot when the input document changes.

```ts
import { docxToPdf } from 'document-kit/layout';
import { createPdfSource, openPdfPreview } from 'document-kit/pdf-preview';

const pdfFor = createPdfSource((snapshot: Uint8Array) => docxToPdf(snapshot, { fonts }));
const pdf = await pdfFor(docx, undefined);
const task = openPdfPreview(pdf, (data) => pdfjs.getDocument({ data }));
const preview = await task.promise;

// Reuse the original generated file on download.
const download = await pdfFor(docx, undefined);
```

The lightweight preview entry can be used on the main thread while the layout engine runs in a worker.
The application supplies its viewer and worker configuration. `openPdfPreview` gives the viewer
a copy that can be transferred to its worker. The PDF's `pageSizes` can reserve space for pages
before the viewer renders them. SVG previews are also available through `toPages`.

## The font rule

The kit never reaches the network. For each family the document draws text in:

1. A face the caller **names** for the family (`namedFonts`) is used first.
2. Otherwise a font **embedded in the document** (Word's *Embed fonts*, `writeDocx`'s `embedFonts`) is used as it
   is.
3. Otherwise the **caller's stand-in** for that name, chosen by the engine's own substitution
   table (`standInFor`): Calibri → Carlito, Cambria → Caladea, Arial/Helvetica → Arimo, Times New
   Roman → Tinos, Courier New → Cousine; other serif names → Tinos, other monospaced names →
   Cousine, anything else → Arimo. Arimo is always asked for: it draws text with no named font,
   math and chart labels.
4. A stand-in nobody supplied is **not fetched**. Its text is drawn with the first family that
   was supplied (Arimo when it was), and the loss says so.

Without a provider, only named and embedded fonts are used; a document with neither cannot be
laid out and `docxToPdf`/`docxToPages` reject with "No fonts…". `docxToHtml` needs no fonts.

Built-in glyphs cover ballot boxes ☐, ☑ and ☒ and circled numbers ⓪ and ①–⑳. The engine uses
these glyphs when the requested character is missing from the supplied fonts. The original Unicode
characters are preserved in the PDF text layer.

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
- **Frames placed on the page** (`w:framePr` with `w:x`/`w:y` against the page or the margin, the model's `frame`
  block): set at that point of the page their anchor falls on, `w:w` wide. With `w:wrap="around"` the flow's
  lines that pass it are set beside it; a frame that does not wrap (`w:wrap="through"`/`"none"`) takes nothing
  from the flow. One in a header is set on every page the header is on.
- **Pictures behind or in front of the text** (an anchored `wp:anchor` with `behindDoc`, the model's
  `floatingImage`): drawn at its offset against the page or the margin, before the page's text when it is behind
  it, after when it is in front; with no wrapping it moves nothing.
- **Character spacing** (`w:spacing` in a run or a style, the model's `characterSpacing`): added after every
  character, so tracked text is measured, broken, aligned and drawn wider (or narrower).
- **Line heights** (`w:spacing` `w:line` with `w:lineRule`, the model's `line`): a multiple (`auto`) is set on the
  face's own line, read from its metrics, so it differs between faces as it does in Word. An exact height
  (`exact`) makes every line box that tall whatever its text, and an at-least height (`atLeast`) makes it that tall
  or the text's own height, the larger; the baseline sits the line's descent above the foot of its box, the extra
  room above the text. Lines of one size in an exact paragraph are therefore exactly that many points apart in any
  face, in the flow and in a frame, as in Word; two faces' baselines differ only by their descents (under a point
  at body sizes), the same on every line.
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
- **CFF faces go into the PDF whole.** A TrueType face is subset to the glyphs a document uses;
  an OpenType face with CFF outlines (`namedFonts`) is embedded whole, so a PDF set in one is
  larger by the face's size. A CID-keyed CFF face (most CJK `.otf` faces) is not drawn with at
  all: it is reported and its family falls back.
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
- **`docxToHtml` has no headers, footers or page numbers**: it flows a Word file. To print the
  block model as HTML with its pages, headers, footers, frames and fonts, use `writeHtml` from
  `document-kit/html` (`docs/html.md`), which reads the document the Word writer reads.
- **Where the baseline sits in an exact line.** The pitch is exact, but where the text sits inside its line box is
  the engine's rule (the descent above the box's foot), checked against Word on the engine's test documents and not
  on every face: a line of larger text in an exact paragraph keeps its box, its baseline lower in it by its larger
  descent. Line for line agreement with Word holds for the boxes; a baseline may differ from Word's by a fraction of
  the face's descent.
- **Fidelity is the engine's.** Stand-ins match widths, not shapes. Layout differences from
  Word in particular documents are the engine's to fix (and record in `docs/ream-upstream.md`).
