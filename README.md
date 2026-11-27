# document-kit

The shared document framework for PGMaps and nalsreport. One block model for a report, a Word writer,
client-side layout to PDF and page pictures, and Word template filling. Everything runs in the browser:
no server, no Node at run time, no network.

| Entry | What it does | Guide |
|---|---|---|
| `document-kit` | The block model: a document as data, with provenance, blanks and placeholders to complete, captions, tick boxes, blocks pinned to the foot of the page, text rules and custom blocks | `docs/model.md` |
| `document-kit/docx` | `writeDocx(document, { embedFonts? })` writes the model as a Word file: named styles, tracking, frames and pictures placed on the page, custom properties, fonts embedded by their licence flags | `docs/model.md` |
| `document-kit/html` | `writeHtml(document, { fonts? })` writes the same model as one self-contained, print-styled HTML page, from the same theme and styles | `docs/html.md` |
| `document-kit/layout` | `docxToPdf`, `docxToPages` (one SVG a page), `docxToHtml`, with fonts passed in as bytes: the bundled stand-ins and the caller's own faces by family name (`namedFonts`, TrueType or CFF) | `docs/layout.md` |
| `document-kit/template` | `fillTemplate`, `templateTags`, `checkTemplate` for Word templates with `{{ tags }}`, and the sentence filters (`capitalize`, `plural`, `prose`, `ranges`) shared with recipes | `docs/template.md` |
| `document-kit/recipe` | Reports from JSON: components and recipes compiled into one JSON object (`compileRecipe`), rendered against data into the block model (`renderRecipe`), with the to-complete list, provenance and hand edits; the same `{{ tag }}` language as Word templates | `docs/recipe.md` |
| `document-kit/ream` | The layout engine itself | `docs/ream-upstream.md` |
| `document-kit/fonts/*` | Open fonts that stand in for Times New Roman, Arial, Calibri, Cambria and Courier New | `fonts/README.md` |

## The chain

```ts
import { document, section, heading, paragraph } from 'document-kit'
import { writeDocx } from 'document-kit/docx'
import { createFontProvider, docxToPdf, docxToPages, layoutDocx } from 'document-kit/layout'

const docx = await writeDocx(document([section([heading(1, 'Report'), paragraph('Made in the browser.')])]))

// The application serves the font files itself; the kit only ever receives bytes.
const fonts = createFontProvider(async ({ file }) => new Uint8Array(await (await fetch(`/fonts/${file}`)).arrayBuffer()))
const pdf = await docxToPdf(docx, { fonts })      // { bytes, pageCount, losses }
const pages = await docxToPages(docx, { fonts })  // { pages: string[], pageSizes, pageText, losses }

// Both from one layout (a preview, then the export):
const laid = await layoutDocx(docx, { fonts })     // { pageCount, pageSizes, pageText, losses, toPdf(), toPages() }
```

One set of faces for all three outputs: the Word file embeds them, the layout draws with them, the print view
carries them.

```ts
import { writeHtml } from 'document-kit/html'

const brand = [{ family: 'Brand Sans', bytes: regular }, { family: 'Brand Sans', weight: 700, bytes: bold }]
const docx = await writeDocx(doc, { embedFonts: brand })                  // fontEmbeddingReport(brand, 'docx') says what went in
const pdf = (await layoutDocx(docx, { fonts, namedFonts: brand })).toPdf()
const html = writeHtml(doc, { fonts: brand })
```

A Word file from anywhere goes through the same layout: one the `docx` library wrote, one a person made in
Word, or a template filled with `fillTemplate`.

## Using it in an application

Add it as a submodule at `vendor/document-kit`, depend on it as `"document-kit": "file:vendor/document-kit"`,
and build it (`npm run build` here) before the application's own dev server or build. The package's entries
point at `dist/`, which is not committed.

## Checks

`npm run check` runs the type checks, the unit tests (the two engines' own suites included), the build, and
`tests/browser`, which runs the built package in headless Chromium with every outside request blocked.
Set `DOCUMENT_KIT_CHROMIUM` to a Chromium or Chrome executable when none is found.

## Where the code came from

`src/ream` is reamkit and `src/template` is officeParser's template engine, both MIT, brought in as source so
they can be changed here. See `NOTICE-REAMKIT.md`, `NOTICE-OFFICEPARSER.md` and the two `docs/*-upstream.md`
files, which list every change from upstream. The fonts' licences are in `fonts/`. Public test documents from
other projects under `tests/ream/fixtures/real` keep their own licences, listed in its `NOTICE.md`.

## Diagram renderers

Use the [diagram registry and template gallery](docs/diagrams.md) to render D3, Mermaid or application-supplied diagrams as document figures.
