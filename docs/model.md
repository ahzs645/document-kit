# The block model and the Word writer

`document-kit` (the `src/model` folder) describes a report as plain data before anything lays it out.
`document-kit/docx` (`src/docx`) writes that data as a Word file. Both run in a browser: no Node, no DOM,
no network. Both applications build the same model; each keeps what only it knows (its wording, its rating
tables, its laboratory rules) in its own code and reaches the kit through custom blocks.

## Units

Every length is in **points** (1/72 inch): margins, indents, spacing, picture boxes, font sizes, border widths.
A picture's own size is in **pixels**, read at 96 to the inch. Colours are six hex digits, with or without `#`.
Column widths in a table are **shares** (any numbers; `[3, 1, 1]` gives the first column three fifths).

## The document

```ts
interface Document {
  metadata?: { title?, author?, subject?, description?, keywords?, created?, modified? }
  theme?: ThemeInput      // any part of the theme; the rest is the kit's neutral default
  sections: Section[]     // at least one
}
```

`metadata.created` (ISO 8601) pins the file's creation and modified times; with it, the same document always
writes the same bytes.

## Sections

A section is a run of pages with one page set-up:

- `page.paper`: `letter` (8.5 × 11 in), `legal` (8.5 × 14 in), `a4`, `tabloid` (11 × 17 in), or `{ width, height }`
  in points, given portrait.
- `page.orientation`: `portrait` or `landscape`. The writer turns the paper; you never swap width and height.
- `page.margins`: `top`, `right`, `bottom`, `left`, and `header` / `footer` (edge to header or footer).
- `page.verticalAlign`: `center` for a cover page.
- `header`, `footer`: lists of blocks. **Left out, a section keeps the header of the section before it** (Word's
  own rule); an empty list means none.
- `firstPage: { header?, footer? }`: a different first page. A part left out there is empty on that page.
- `pageNumbers: { start?, format? }`: restart numbering or number in roman or letters.
- `body`: the blocks.

`pageBox(setup, theme)` gives a section's page size, margins and text box in points.

## Blocks

| Block | What it says |
|---|---|
| `heading` | Levels 1 to 6. Written as Word's Heading styles with outline levels, so the navigation pane and a table of contents read them. |
| `paragraph` | Rich text with a style (`normal`, `title`, `subtitle`, `caption`, `note`, `small`), alignment (including justify), spacing before/after/line, indent (left, right, first line, hanging), keep with next, keep lines together, page break before, tab stops (`at: 'end'` is the right edge), borders and shading. |
| `list` | Bulleted or ordered, nested to any depth; an item may change kind for its children. Real Word numbering; each ordered list counts from its own `start`. |
| `fields` | Label and value rows, as lines (`Label: value`, optional tab stop) or as a two-column table. A `null` value prints a blank rule. |
| `table` | Rows with roles: `header` (theme header look, repeated at the top of each page while among the first rows), `group` (shaded band), `note` (one small cell across the table), `spacer` (a short empty row), `body`. Cells span columns and rows (`colSpan`, `rowSpan`, listed as in HTML: a covered cell is not written). Per-cell alignment, vertical alignment, shading, borders, text style, a corner-to-corner rule (`diagonal`, a struck cell), and blocks inside a cell (paragraphs, lists, pictures, nested tables). Column shares, width as a share of the text, table borders, cell padding, text size, a caption above or below, a note under it, rows kept whole. |
| `figure` | A PNG or JPEG with its pixel size (read from its header when not given), fitted into a box (`fit: { width, height, upscale }`, or `'page'` to fill what the section's page leaves under the caption), a caption, a note, alignment. `image: null` prints a placeholder box with `placeholder` text. |
| `figureGrid` | Pictures in `columns` columns (one column stacks them), each with a label above or below and a placeholder when missing, under one caption. |
| `signature` | A label, the signature picture or room to sign by hand, the signing rule with the name under it, and lines under that (`Title`, `Date`), any of which can be blank. |
| `pageBreak` | The next block starts a new page. Every block also takes `pageBreakBefore`. |
| `toc` | A table of contents over heading levels 1 to `levels`, filled from the headings when written (page numbers are filled when Word updates the field). |
| `spacer` | Vertical space of an exact height. |
| `custom` | A block the kit does not know, with a `kind` and the application's `data` (below). |

Every block may carry `id` (the application's own), `source`, and `pending` (see blanks).

## Inline text

`RichText` is a string, one run, or a list of runs. Runs:

- text (`{ type: 'text', text }`) with `bold`, `italic`, `underline`, `strike`, `superscript`, `subscript`,
  `smallCaps`, `allCaps`, `color`, `size`, `font`, `highlight` (Word's sixteen highlight colours), and `link`
  (a URI; written as a real hyperlink). In text, a newline is a line break and a tab character a tab.
- `field`: `page`, `pages` (the document's page count) or `sectionPages`.
- `break`, `tab`.
- `blank`: a rule to fill by hand, with `what` it is for.
- `image`: a picture in the line (a logo in a header) at a stated size in points.
- `sentence`: a template over named slots (below).

## Theme

The theme is how a document looks, so one application can be Times with a green band and the other Arial and grey,
and the kit knows neither: body font, size and colour; title, subtitle, six heading levels, caption, note, small,
header and footer text (font, size, colour, bold, italic, caps, alignment, spacing); link colour; the table look
(header fill and text, group fill, note text, label fill, outer and inner borders, cell padding, spacer height,
placeholder box); the caption words and separator (`Table 3: Results`, or `Figure 2. View`); the blank rule's
width; and the default page. `resolveTheme(partial)` merges a partial theme over `DEFAULT_THEME`, heading level by
heading level.

## Captions

`numberCaptions(documentOrBlocks)` numbers captioned tables and figures in reading order, each kind on its own
count (figures and figure grids share one), across every section. A caption with `numbered: false` takes no number;
`'silent'` takes one but prints its own text alone. `captionLabel` and `captionPrefix` give the printed words. The
Word writer writes the number as a `SEQ` field with its value already in it, so Word's table of figures and
cross-references find it. Custom blocks are numbered after they are expanded (number `expandDocument(doc, registry)`
to ask what the writer will print).

## Provenance and blanks

A value may say where it came from with `source`, a string tag the application defines (`'run'`, `'reviewer'`,
`'lab'`, `'filled'`). Runs, field rows, table cells, sentences and blocks carry it. The kit never interprets it; the
Word writer marks text by it only when told: `writeDocx(doc, { sourceStyles: { reviewer: { highlight: 'yellow' } } })`.

A value nobody gave is `null`: a field row's value, a table cell, a signature's name or line, a figure's image. It
prints as a blank rule (an empty cell in a table, unless `nullCells: 'rule'`). A sentence with a slot that has no
value is not printed at all, and its separating space goes with it.

`toComplete(document, { registry? })` walks the document (bodies, headers, footers, inside table cells) and lists
everything still to complete: each null value, blank rule, left-out sentence (with its missing slots), missing
picture, and each `pending` item a block declares (what an application knows is missing but cannot show). Each item
names the heading it falls under. With a registry, custom blocks are expanded first and walked too.

## Sentences

`sentence('It is {visibility} from {viewpoint}.', values, { label, source })` fills `{slot}`s from `values`; a slot
that is null, undefined or an empty string leaves the sentence out. `fillSlots` is the same rule as a pure function.
The kit does no grammar: an application that tidies its sentences (articles, plurals) fills them itself and passes
text runs.

## Custom blocks and the registry

```ts
const registry = defineCustomBlocks({
  matrix: (data: MatrixData, context) => [table(/* rows built from data, one cell shaded */)],
})
await writeDocx(doc, { registry })
```

An expander receives the block's data, the context (resolved theme, the section's page box, the section's index) and
the block, and returns the kit's blocks, which may themselves hold custom blocks. Writers expand before writing
(`expandDocument`). A kind with no expander throws `UnknownCustomBlockError`, whose message and `kind` name it; a
custom block is never dropped silently. A custom block's `pageBreakBefore` and `pending` carry onto what it becomes.

## One short example

```ts
import { document, section, heading, paragraph, fields, field, table, headerRow, row, caption,
         pageNumber, pageCount, sentence, toComplete } from 'document-kit'
import { writeDocx } from 'document-kit/docx'

const doc = document([
  section([
    heading(1, 'Results'),
    fields([field('Date', '12 March'), field('Checked by', null)]),
    table([headerRow(['Item', 'Value']), row(['pH', '7.1'])], { caption: caption('Readings') }),
    paragraph([sentence('The site faces {aspect}.', { aspect: null }, { label: 'Aspect' })]),
  ], { footer: [paragraph(['Page ', pageNumber(), ' of ', pageCount()], { align: 'center' })] }),
], { metadata: { title: 'Made-up report', created: '2026-03-12T09:00:00Z' } })

const bytes = await writeDocx(doc)            // Uint8Array, browser or Node
const todo = toComplete(doc)                  // [Checked by, Aspect (needs aspect)]
```

## What the Word writer pins

The `docx` library gives hyperlink relationships random ids, and numbers drawings and bookmarks from counters that
run on across every document made in the same page; it stamps the creation time and every zip entry with the moment
of writing. `writeDocx` renames the random relationship ids (`rIdL1`…), numbers drawing ids, bookmark ids and the
contents' `_Toc` names from 1 in each file, writes the zip entries with one fixed time, and, when the document gives
`metadata.created`, writes that time in `docProps/core.xml`. The same document gives the same `word/document.xml`
always, and the same bytes whenever `created` is given.

## From the applications' blocks to the kit's

### The laboratory reports (Word today)

| Their part | Kit block |
|---|---|
| Letterhead image | `inlineImage` in a section `header` paragraph (bytes from the app; recolouring stays in the app) |
| Report title | `paragraph(…, { style: 'title' })`, colour from the theme |
| Attn / Address / Re / Date / Reference lines | `fields` (lines layout); an unfilled value is `null` |
| Text section (`type: 'section'`, `text`) with `## ` sub-headings and `- ` bullets | `heading` / `paragraph` with `keepWithNext` / `list` (the app parses its markup) |
| `BuilderTable` header rows, `headerRows`, `hideColumnHeader` | `table` rows with role `header` (`repeat` decides which repeat) |
| Row roles `groupHeader`, `note`, `spacer`, `data` | roles `group`, `note`, `spacer`, `body` |
| Merged header cells (repeated labels, blanks under a label) | `colSpan` / `rowSpan` (the app decides the merges) |
| Struck cells of a short final block | cell `diagonal: 'up'` |
| Footnote markers in cells and footnote lines | `sup('a')` runs; footnote lines as `note` paragraphs |
| Chemistry (CaCO3) | `sub('3')` runs (the app finds the formulas) |
| `align` (label/value/header), `exportLayoutPreset` density | `columnAlign`, cell `align`, table `size`, `cellPadding`, `columnWidths` |
| Wide tables: `separateTables` / `stackedBlocks` | several `table`s, or one table with `header` rows inside it (they do not repeat); the split itself stays in the app's `table-layout.ts` |
| `tableContinuation` (repeat / none / auto) | row `repeat` and table `repeatHeader: false` |
| `keepTableTitleWithTable` | always on for captions above a table |
| Captions "Table N: …", `captionMode` none / explicit | `caption`, `numbered: false` / `'silent'` |
| Figures (`layout` wide / standard / half) and placeholders | `figure` with a `fit` box; `image: null` and `placeholder` |
| Attachments (image page, companion file, external report) | `figure` with `pageBreakBefore`, or a placeholder `figure` |
| Signature block (label, picture or wet-ink gap, rule, name and title) | `signature` |
| Header code; footer reference + "Page X of Y" | section `header` paragraph; footer paragraph with `tabs: [{ at: 'end', align: 'right' }]`, `pageNumber()`, `pageCount()` |
| Links in prose (URLs, e-mail) | `link` runs (the app finds them) |
| Filled-value and placeholder marks | `source` tags with `sourceStyles`, or `highlight` on runs |
| Letter / A4 | `page.paper` |

### The visual-impact reports (PDF, HTML and Markdown today)

| Their block | Kit block |
|---|---|
| `heading` 1–3 | `heading` |
| `paragraph` of `Sentence`s | `paragraph` of `sentence` runs, or the app's own filled sentences as text runs; unfilled ones are left out and listed by `toComplete` |
| `text` with `Provenance` | `paragraph` with `source` |
| `list` | `list` |
| `fields` (`FieldRow`, `value: null`) | `fields` with `layout: 'table'` |
| `table` with `note` | `table` with a `header` row and `note` |
| `matrix` (marked cells, tints, class grid) | `custom` → `table` (marked cell shaded; tints as cell `shading`) |
| `scale` (bar, arrow axis, guide) | `custom` → a table of coloured band cells, or a picture the app draws (see below) |
| `choices` (tick boxes, boxed) | `custom` → `paragraph` with ☒ / ☐, or a one-row `table` when boxed |
| `figure` with placeholder | `figure` |
| `signature` lines | `signature` with `lines` whose values are `null` |
| `form` (numbered sections with title bars) | `table` whose title rows are `group` rows and whose cells hold blocks |
| `pageBreak` | `pageBreak` |
| `figurePage` (title box, facts, maps, before/after) | its own landscape `section` holding a `table` with picture cells (custom expanders cannot open a section; the app places it) |
| `figurePair` (stacked, `landscape`, `sheet`) | `figureGrid` with `columns: 1`, in a landscape or `tabloid` landscape section |
| `mapFigure` (legend, scale, north arrow) | `table`: picture cell and a legend cell, or the picture with its furniture drawn in |
| `figureGrid` | `figureGrid` |
| `ReportLayout.paper` letter / legal, `cover`, `contents` | `page.paper`; a first section with `verticalAlign: 'center'` and empty header/footer; a `toc` block |
| Wide tables on landscape pages; analysis sheets on 11 × 17 landscape | the app puts them in `landscape` / `tabloid` landscape sections (`isWide` and `isSheet` stay in the app) |
| Running header and "Page X of Y" | section `header` / `footer` with `pageNumber()`, `pageCount()` |
| `toComplete` list "To complete before signing" | `toComplete(doc, { registry })`, printed by the app as a `list` |
| Draft tints by provenance | `sourceStyles` |

## What the model cannot say yet

- Shapes and vector drawing: a scale bar with an arrow, a mark between matrix cells, an X on a grid, a north arrow.
  The app draws them as a PNG or builds them from table cells. SVG pictures are not written.
- A frame pinned to the foot of the page (the laboratory contact block's `pinBottom`).
- A label printed inside a picture's corner (the `inset` placement of a figure pair); labels go above or below.
- A list of figures (the contents' Images list); only headings make a table of contents.
- Internal links and bookmarks the app names; `link` takes URIs only.
- Footnotes and endnotes as Word notes (footnote markers are superscript runs).
- Automatic splitting of a too-wide table into blocks, and automatic placing of wide tables on landscape pages.
- A custom block that opens a section of its own.
- JPEG orientation from EXIF: a rotated photo is drawn as stored.
