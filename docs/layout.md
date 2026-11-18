# document-kit/layout

A Word file (`.docx`) to a PDF, to one SVG picture per page, or to flowed HTML, in the browser,
with fonts the caller hands in. It is a small front over the reamkit engine (`src/ream`,
`document-kit/ream`): the engine's reader, layout and writers do the work.

## API

```ts
import { createFontProvider, docxToPdf, docxToPages, docxToHtml } from 'document-kit/layout';

const fonts = createFontProvider(async ({ file }) =>
  new Uint8Array(await (await fetch(`${base}/fonts/${file}`)).arrayBuffer()),
);

const { bytes, pageCount, losses } = await docxToPdf(docx, { fonts });
const { pages, pageSizes, pageText } = await docxToPages(docx, { fonts });
const { html } = await docxToHtml(docx);
```

- `docxToPdf(docx, { fonts?, password?, info?, pdfA?, tagged? })` →
  `Promise<{ bytes, pageCount, losses }>`. `info` (title, author, …) is laid over the Word
  file's own properties; `pdfA` is a level such as `'PDF/A-2b'`; `tagged` writes a tagged PDF.
- `docxToPages(docx, { fonts?, password? })` →
  `Promise<{ pages, pageSizes, pageText, losses }>`. The same layout the PDF gets.
  - `pages[i]` is a standalone SVG document for page *i*, drawn by the engine's SVG writer one
    page at a time. Text is drawn from the faces' own glyph outlines (`<path>` definitions and
    `<use>`), so a page shows the same in any viewer with no fonts installed, and there is no
    selectable or searchable text in it. Each page carries its own glyph definitions, so pages
    do not share anything and some outlines repeat between pages. Every id in a page is
    prefixed `dk-p<page>-`, so pages can be placed inline in one HTML document without taking
    each other's glyphs (the engine numbers them from zero on every page). The page sits on a white
    rectangle with a light grey outline (the engine's preview style). Size is in points:
    `width`/`height`/`viewBox` are the page in pt.
  - `pageSizes[i]` is `{ widthPt, heightPt }`.
  - `pageText[i]` is the text of the page's laid-out lines, one per `\n`, in layout order (use it
    for search and checks). Text inside charts and drawn shapes is not included.
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
`PagesResult`, `PageSize`, `HtmlResult`, `LayoutLoss`, `DocumentInfo`, `PdfALevel`.

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
  and no face for that script is bundled. Hebrew, Greek and Cyrillic letters the bundled faces
  happen to carry are still drawn; others come out as empty boxes (and a `dropped` `text` loss).

## Known gaps

- **Word files only.** The engine also reads `.doc`, `.xlsx`, `.xls`, `.pptx`, `.ppt` and PDF;
  this API takes `.docx`. Use `document-kit/ream` for the rest (and pass fonts yourself).
- **No caller-named fonts.** A caller cannot yet supply their own bytes for a named family
  (say, a licensed Calibri); only the five stand-ins. The engine supports it (`embeddedFonts`);
  the API does not expose it yet.
- **No scripts beyond Latin, Greek and Cyrillic** in the bundled set (see above).
- **Two calls lay out twice.** `docxToPdf` and `docxToPages` each run the layout.
- **Bundle size.** The engine is large: `dist/layout.js` is small but imports a shared chunk of
  about 1.5 MB unminified (about 440 kB gzipped) that also holds readers this API does not use
  (the build cannot split them out because `document-kit/ream` shares it).
- **SVG pages have no text layer** and repeat glyph outlines per page; `pageText` stands in for
  search.
- **HTML has no headers, footers or page numbers.**
- **Fidelity is the engine's.** Stand-ins match widths, not shapes. Layout differences from
  Word in particular documents are the engine's to fix (and record in `docs/ream-upstream.md`).
