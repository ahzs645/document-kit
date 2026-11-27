# The block model and the Word writer

`document-kit` (the `src/model` folder) describes a report as plain data before anything lays it out.
`document-kit/docx` (`src/docx`) writes that data as a Word file. Both run in a browser: no Node, no DOM,
no network. Both applications build the same model; each keeps what only it knows (its wording, its rating
tables, its laboratory rules) in its own code and reaches the kit through custom blocks.

## Units

Every length is in **points** (1/72 inch): margins, indents, spacing, picture boxes, font sizes, border widths.
A picture's own size is in **pixels**, read at 96 to the inch. Colours are hex: `#rrggbb` or the short `#rgb`
(each digit doubled, as in CSS), in any case, with or without the `#`. Anything else throws an error naming the value.
Column widths in a table are **shares** (any numbers; `[3, 1, 1]` gives the first column three fifths).

## The document

```ts
interface Document {
  metadata?: { title?, author?, subject?, description?, keywords?, created?, modified?, custom? }
  theme?: ThemeInput      // any part of the theme; the rest is the kit's neutral default
  styles?: DocumentStyles // named paragraph and character styles beyond the theme's (below)
  textRules?: TextRule[]  // automatic text touches: links, italic terms, raised markers (below)
  sections: Section[]     // at least one
}
```

`metadata.created` (ISO 8601) pins the file's creation and modified times; with it, the same document always
writes the same bytes.

`metadata.custom` holds custom document properties by name (Word's File › Properties › Custom,
`docProps/custom.xml`): text, a number (a whole number within 32 bits is written as `vt:i4`, any other as `vt:r8`),
a yes/no (`vt:bool`) or a `Date` (`vt:filetime`, UTC). An application's own ids, provenance and font requirements
travel in the file this way. A name is 1 to 255 characters; a number that is not finite, or an invalid `Date`, stops
the writer with the property named. `writeHtml` writes each as a `<meta name content>`.

## Sections

A section is a run of pages with one page set-up:

- `page.paper`: `letter` (8.5 × 11 in), `legal` (8.5 × 14 in), `a4`, `tabloid` (11 × 17 in), or `{ width, height }`
  in points, given portrait.
- `page.orientation`: `portrait` or `landscape`. The writer turns the paper; you never swap width and height.
- `page.margins`: `top`, `right`, `bottom`, `left`, and `header` / `footer` (edge to header or footer).
- `page.verticalAlign`: `top` (default), `center` for a cover page, or `bottom`. Written as Word's `w:vAlign`; the
  layout honours it on every page of the section.
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
| `paragraph` | Rich text with a style (`normal`, `title`, `subtitle`, `caption`, `note`, `small`, or the id of a style the document defines), alignment (including justify), spacing before/after/line (a multiple, or exact or at-least points: see Line heights), indent (left, right, first line, hanging), keep with next, keep lines together, page break before, tab stops (`at: 'end'` is the right edge), borders (each with its distance from the text, `space`, as a rule set a little under a letterhead) and shading. |
| `list` | Bulleted or ordered, nested to any depth; an item may change kind for its children. Real Word numbering; each ordered list counts from its own `start`. |
| `fields` | Label and value rows, as lines (`Label: value`, optional tab stop) or as a two-column table. A `null` value prints a blank rule. In lines, `rowSpacing` is the space after each row, and a row with `sameLine: true` goes on the line of the row before it after `sameLineGap` points (24) — "E-mail: …   Phone: …" — when it fits; when it does not, it moves whole to the next line, since its spaces and hyphens are written unbreakable (the hyphen as Word's no-break hyphen, which Word and `document-kit/layout` draw as the font's own hyphen and read back as `-`). It then shares the paragraph of the row before it, so it lacks that row's `rowSpacing` above it. |
| `table` | Rows with roles: `header` (theme header look, repeated at the top of each page while among the first rows), `group` (shaded band), `note` (one small cell across the table), `spacer` (a short empty row), `body`. Cells span columns and rows (`colSpan`, `rowSpan`, listed as in HTML: a covered cell is not written). Per-cell alignment, vertical alignment, shading, borders, padding (`padding`, over the table's `cellPadding`), text style, a corner-to-corner rule (`diagonal`, a struck cell), text turned a quarter (`textDirection: 'up'` or `'down'`), and blocks inside a cell (paragraphs, lists, pictures, nested tables). Column shares, width as a share of the text, table borders, cell padding, text size, a caption above or below, a note under it, rows kept whole (`cantSplit`), a row kept with the next (`keepWithNext`, on by default for a `header` row with a row after it), the whole table on one page (`keepTogether`), its own line height (`line`, a multiple or points), and `autoFit` (Word's AutoFit to contents, the shares only a start; fixed by default). A `header` row further down the table (a stacked block's own header) is drawn in the header look and not repeated. |
| `spreadsheet` | One sheet of an Excel workbook (`.xlsx` bytes in `workbook`), written as a native table: details below. |
| `figure` | A PNG or JPEG with its pixel size (read from its header when not given), fitted into a box (`fit: { width, height, upscale }`, or `'page'` to fill what the section's page leaves under the caption), a caption, a note, alignment. `crop: { top, right, bottom, left }` shows a band of it (the share, 0 to 1, cut away at each edge; the part shown is what is fitted), and `border: { width, color }` draws a solid frame round it. `image: null` prints a placeholder box with `placeholder` text. |
| `figureGrid` | Pictures in `columns` columns (one column stacks them), each with a label above or below and a placeholder when missing, under one caption. |
| `signature` | A label, the signature picture or room to sign by hand, the signing rule with the name (and title) under it, and lines under that (`Title`, `Date`), any of which can be blank. Kept on one page. Options below. |
| `choices` | A tick-box question: `label`, `options` (strings), `chosen` (one of them, a list for several, or null: unanswered, listed), `answer` (a free-text answer line; null is a blank to complete; absent, no line), `layout` `inline` (default) or `list`, `boxed` (a ruled box, the boxes at the right). |
| `pinned` | Blocks pinned to the foot of the page they fall on (below). |
| `frame` | Blocks in a text frame placed anywhere on the page: `x`, `y`, `width` (and `height`) in points, against the page or the margins (below). |
| `floatingImage` | A picture placed out of the flow, behind the text by default: a full-page cover background (`box: 'page'`), a watermark, a mark in a corner (below). |
| `pageBreak` | The next block starts a new page. Every block also takes `pageBreakBefore`. |
| `toc` | A table of contents over heading levels 1 to `levels`, filled from the headings when written. Word fills the page numbers when it updates the field; `document-kit/layout` fills them itself, each in its page's numbering. |
| `spacer` | Vertical space of an exact height. |
| `custom` | A block the kit does not know, with a `kind` and the application's `data` (below). |

Every block may carry `id` (the application's own), `source`, and `pending` (see blanks).

## Inline text

`RichText` is a string, one run, or a list of runs. Runs:

- text (`{ type: 'text', text }`) with `bold`, `italic`, `underline`, `strike`, `superscript`, `subscript`,
  `smallCaps`, `allCaps`, `color`, `size`, `font`, `highlight` (Word's sixteen highlight colours),
  `characterSpacing` (tracking in points, added after every character, negative to condense; Word's `w:spacing` on
  the run: tracking given in em is `em × size`), `characterStyle` (the id of a character style the document defines,
  its look under the run's own), and `link` (a URI; written as a real hyperlink). In text, a newline is a line break
  and a tab character a tab.
- `field`: `page`, `pages` (the document's page count) or `sectionPages`.
- `break`, `tab`.
- `blank`: a rule to fill by hand, with `what` it is for. `show: 'placeholder'` prints `[what]` instead, in the
  theme's `placeholderMark` look (a yellow highlight by default): a value nobody gave, set apart in its sentence.
- `checkbox`: a tick box, `checked` or not (below).
- `image`: a picture in the line (a logo in a header) at a stated size in points.
- `sentence`: a template over named slots (below).

## Tick boxes

A `checkbox` run, and each option of a `choices` block, is written as Word's own check box (a `w14:checkbox`
content control showing ☒ or ☐ from MS Gothic), which a reader can tick in Word. The bundled fonts have no ☐ ☑ ☒,
so `document-kit/layout` draws them from a small face of boxes it builds for them (`src/layout/tickBoxes.ts`): a
square, crossed when ticked, in the PDF and the page pictures alike. The characters stay in the text, so
`pageText` reads a box as `☐` or `☒`, and `plainText` the same.

## Pinned blocks

`{ type: 'pinned', blocks, spaceBelow? }` takes its blocks out of the text's flow and sets them at the foot of the
page the block falls on: a disclaimer at the bottom of the first page, a contact block at the bottom of the last.
The Word writer gives each paragraph the blocks make one text frame anchored to the bottom margin and centred across
the text width (`w:framePr`, `w:vAnchor="margin"`, `w:yAlign="bottom"`, wrapping around); Word sets consecutive
paragraphs that share a frame as one box. `spaceBelow` keeps empty room under the blocks inside the frame, so the
last line clears the margin.

- It holds headings, paragraphs, lists, `lines` fields, a signature, tick boxes (not boxed), spacers and pictures
  without a caption. A frame cannot hold a table, so a table, a boxed question, `table` fields, a placeholder
  picture or a captioned figure there is refused with an error naming it.
- The text of the page flows above it. A line that would run into it goes on to the next page.
- When the text above it already reaches lower than its top — the page is too full for it under its text — it goes
  to the foot of the next page instead (with nothing else on that page if nothing follows). `document-kit/layout`
  does both (`docs/ream-upstream.md`, changes 10 and 11); Word does the same.
- Its blocks are walked like any others: to-complete, caption numbers, custom blocks and hand edits reach into it.

## Named styles

`document.styles` defines paragraph and character styles beyond the theme's, written as named Word styles (they show
in Word's style gallery and keep their names when the file is edited) and as CSS classes by `writeHtml`. A paragraph
names one with `style`, a run with `characterStyle`.

```ts
styles: {
  paragraph: [
    { id: 'AffiliationDepartment', name: 'Affiliation Department', text: { size: 12, bold: true, characterSpacing: 0.6, after: 2 } },
    { id: 'AffiliationSubdepartment', basedOn: 'AffiliationDepartment', next: 'normal', text: { size: 10.5, bold: false } },
  ],
  character: [{ id: 'Accent', text: { color: 'b45309', characterSpacing: 1 } }],
}
paragraph('Faculty of Examples', { style: 'AffiliationDepartment' })
text('tracked', { characterStyle: 'Accent' })
```

- A paragraph style: `id` (a letter, then letters and digits), `name` (what Word shows; the id when absent),
  `basedOn` (a theme style, `normal`, `title`, `subtitle`, `caption`, `note`, `small`, `heading1`–`heading6`, or
  another defined id; `normal` when absent), `next`, `text` (a `ThemeText`: font, size, colour, bold, italic, caps,
  small caps, `characterSpacing`, alignment, spacing before and after, line height (a multiple or points),
  `keepWithNext`), `indent`,
  `borders` (each with its `space`), `shading` and `keepLinesTogether`.
- A character style: `id`, `name`, `basedOn` (another character style) and `text` (a `TextStyle`).
- The Word file holds each style's own settings over its base, as Word keeps them; `styleTable(theme, styles)`
  (from `document-kit`) gives each style resolved through its bases, which is what `writeHtml` writes.
- Refused with a `StyleError` naming it: an id that is malformed, defined twice, or one of the theme's own (`Title`,
  `Heading1`, `note`…, in any case); a `basedOn` or `next` naming no style; styles based on each other in a circle;
  a paragraph or run naming a style the document does not define.

The theme's text roles take `characterSpacing` too (`theme.headings[0].characterSpacing`), written on their Word
styles.

## Line heights

A line height (`LineHeight`) is a multiple of single spacing or a height in points, wherever the model takes one: a
paragraph's `spacing.line`, a theme text role's `line` (`theme.body.line`, `theme.headings[0].line`…), a named
paragraph style's `text.line`, a table's `line` and the theme's `table.line`.

```ts
paragraph('Lines 14pt apart in any face', { spacing: { line: { exact: 14 } } })
paragraph('At least 14pt; a larger line grows', { spacing: { line: { atLeast: 14 } } })
paragraph('1.15 of the face\'s own line', { spacing: { line: 1.15 } })
{ theme: { body: { line: { exact: 13 } } }, styles: { paragraph: [{ id: 'CoverLabel', text: { line: { exact: 12 } } }] } }
```

- A **multiple** (`1.15`) is Word's "Multiple": it is measured on the face's own line, which Word and the layout read
  from the face's metrics, so the same multiple stands taller in one face than in another (11pt Carlito's single
  line is 13.4pt, Arimo's 12.6pt) and lines land differently between faces. Written `w:line` in 240ths.
- `{ exact: n }` is Word's "Exactly": every line is `n` points tall whatever its text, so lines land in the same
  places in Word, the PDF and the HTML whatever the face. Written `w:line` in twips with `w:lineRule="exact"`. Text
  taller than the line is not given more room (in Word its tops are cut off).
- `{ atLeast: n }` is Word's "At least": a line is `n` points tall unless its text needs more. `w:lineRule="atLeast"`.
- Points from more than 0 to 1584; a multiple from more than 0 to 132. Anything else (both keys, another key, 0, a
  negative) throws an error naming the value. `readLineHeight` (from `document-kit`) reads one as the writers do.
- A multiple written over a style whose line height is in points (a paragraph's `line: 1` in a body set exactly, a
  style based on an exact one, the header and table text that the kit sets single over the body) is written with
  `w:lineRule="auto"`: Word reads a rule left out as a multiple, but the layout and LibreOffice take the style's.
  Elsewhere a multiple is written with no rule, as before, so existing documents keep their bytes.

## Frames and floating pictures

`frame(blocks, { x, y, width, height? }, options)` sets blocks in a text frame placed on the page: a cover's title
over its artwork, a label, an address window. Every paragraph the blocks make carries one frame (`w:framePr` with
`w:x`, `w:y`, `w:w`), so Word and `document-kit/layout` set them as one box there.

- `x`, `y` are points from the top-left corner of the page, or of the margins with `relativeTo: 'margin'`; `width`
  points. `height` is a least height (`heightRule: 'atLeast'`, the default) or exact (`'exact'`, the content cut to
  it); absent, the frame is as tall as its content.
- `wrap`: `around` (default, Word's own; text flows beside it), `notBeside` (above and below only), `through` or
  `none` (the text takes no notice: for a cover whose flow is empty, or text that must not move). `space`: points
  kept between the frame and text flowing round it.
- It holds what a pinned block holds (headings, paragraphs, lists, `lines` fields, a signature, tick boxes, spacers,
  pictures without a caption); a table or a floating picture is refused with an error naming it. A pinned block or
  a frame inside a frame joins it.
- It is placed on the page its anchor falls on (where the block stands in the flow). In a section's `header` it is
  placed on every page the header prints on; in `firstPage.header`, on the section's first page.

`floatingImage(image, box, options)` places a picture out of the flow (Word's anchored drawing, no wrapping):

- `box: 'page'` covers the whole page (a cover background: give the picture the page's proportions); otherwise
  `{ x, y, width, height }` in points from the page's corner, or the margins' with `relativeTo: 'margin'`. The
  picture is stretched to the box.
- `behind: true` (default) puts it behind the text; `false` in front. `alt` is its description. Behind the text means
  under everything the body paints, its fills too: a shaded cell or paragraph over a page ground keeps its fill in
  Word, the PDF and the HTML alike.
- Pictures in the same layer stack in document order: one placed later stands over one placed earlier (Word's
  `relativeHeight`, written 1, 2, 3… through the document; the `docx` library's own default was the picture's height,
  which put a large ground over a small mark placed after it).
- Its anchor is an empty paragraph one twip tall, so it takes no room in the flow. In a `header` it is on every page
  of the section (a background for every page), in `firstPage.header` on the first page only.
- `image: null` (a recipe's picture the data did not give) places nothing and is listed to complete.

A designed cover is a section with no margins, its artwork in `firstPage.header` (or `header`) and its words in
frames:

```ts
section(
  [frame([heading(1, 'Annual review')], { x: 54, y: 120, width: 420 }, { wrap: 'through' })],
  { page: { paper: 'letter', margins: { top: 0, right: 0, bottom: 0, left: 0, header: 0, footer: 0 } },
    header: [floatingImage(artwork, 'page', { alt: 'Cover artwork' })], footer: [] },
)
```

## Embedded fonts

`writeDocx(doc, { embedFonts })` embeds fonts in the file (Word's *Embed fonts*), by family name, so the document
shows in its own faces where they are not installed, and `document-kit/layout` draws with them:

```ts
const fonts: NamedFont[] = [
  { family: 'Brand Sans', bytes: regular },             // weight 400, upright: the regular face
  { family: 'Brand Sans', weight: 700, bytes: bold },   // 600 and heavier: the bold face
  { family: 'Brand Sans', style: 'italic', bytes: italic },
]
const docx = await writeDocx(doc, { embedFonts: fonts })
fontEmbeddingReport(fonts, 'docx')   // [{ family, weight, style, slot, format, fsType, embedded, reason }, …]
```

- Word knows four faces a family: regular, bold, italic, bold italic. A face of weight 600 or more fills the bold
  slot; where two faces want one slot, the one nearer 400 (or 700) has it. A light or black cut is a family of its
  own (`Brand Sans Light`). Runs and styles name the family (`font: 'Brand Sans'`).
- Each face is checked against its own licence flags (OS/2 `fsType`) and against what Word can carry, and only
  those `fontEmbeddingReport(fonts, 'docx')` marks `embedded` go in: **TrueType outlines** whose flags allow
  **installable or editable** embedding. Refused, with the reason in the report: OpenType faces with CFF outlines
  (`.otf`; Word embeds TrueType only: install the face, or supply a TrueType build), preview & print embedding (an
  editable Word file cannot use it), restricted and bitmap-only faces, a face with no OS/2 table, a collection, a
  family with no regular or bold face. A refused face is never dropped silently: `writeDocx` leaves it out, the
  report says why, and the application can carry the list (a custom property, say).
- Faces go in whole, never subset, obfuscated as ECMA-376 §17.8.1 says, under `w:embedRegular`, `w:embedBold`,
  `w:embedItalic`, `w:embedBoldItalic` in the font table, with `w:embedTrueTypeFonts` set. The obfuscation key is
  made from the font's own bytes, so the same fonts write the same file.
- The licence flags are the font's statement; the vendor's licence terms still go with the file.

## Signature options

| Option | Default | |
|---|---|---|
| `label`, `labelPosition` | | `above` the block, or `below` the name and lines |
| `labelStyle`, `labelGap` | bold, 4 | The label's look; points between it and the picture or signing room |
| `image`, `imageBox`, `imageUpscale` | none, 140 × 55, true | The picture is fitted into the box, keeping its shape, and sits on the rule |
| `signingSpace` | 36 | Room to sign by hand when there is no picture, points |
| `name`, `title`, `titleLayout`, `titleSeparator` | | `sameLine`: "Name, Title"; `nextLine`: the title under the name. A null name or title is a blank to complete |
| `nameStyle` | bold | |
| `lines` | | Rows under the name ("Date: …"); a row with an empty label prints its value alone |
| `align` | `left` | `left`, `center` or `right` (the rule and the lines follow it) |
| `ruleWidth`, `ruleThickness`, `ruleColor`, `ruleGap` | 216, 0.75, body colour, 2 | The rule's length and weight; points between it and the name |
| `spaceAbove`, `spaceBelow` | 12, 2 | |
| `keepTogether` | true | Every paragraph kept with the next and whole, so the block never splits across pages |

## Text rules

`document.textRules` states, as data, the automatic touches an application gives its text: web and e-mail addresses
as links, a species name in italics, a footnote marker raised, a formula's digits set low. The Word writer applies
them to the text's ordinary runs before it writes (`applyTextRules`), so the layout and the PDF follow, and the
document an application keeps is left as it was. Each rule says what to find and what to do with it:

```json
"textRules": [
  { "id": "links", "builtIn": "links", "in": ["paragraph", "value"] },
  { "id": "species", "terms": ["M. exemplum"], "style": { "italic": true }, "in": ["cell", "note"] },
  { "id": "markers", "builtIn": "trailingMarker", "markers": "a-h", "in": ["cell", "label"] },
  { "id": "footnotes", "builtIn": "leadingMarker", "markers": "a-h", "auto": false },
  { "id": "formulas", "builtIn": "formulaDigits", "formulas": ["CaCO3", "NO3"] },
  { "id": "run", "pattern": "Run (\\d+)", "group": 1, "style": { "bold": true } }
]
```

| Key | |
|---|---|
| `pattern`, `flags`, `group` | A regular expression from the safe subset below (`flags`: `i` only); the style or link goes on the whole match or on one capture group, by number or name |
| `terms`, `ignoreCase` | Words or phrases, matched whole (a term starting or ending in a letter does not match inside a longer word), longest first |
| `builtIn` | `links` (e-mail addresses → `mailto:`, `https://` and other schemes, `www.…`, and bare domains `example.org/page` on the top-level domains in `domains`, a common list by default); `trailingMarker` ("Total a", "Lead a, b": the letters after a space at the end, raised); `leadingMarker` ("a Dried at 105 °C", "bAir dried": the letter opening the text, raised); `formulaDigits` (the digits of the listed `formulas`, standing alone, set low) |
| `style` | A text style (`italic`, `bold`, `superscript`, `subscript`, `color`, `underline`, …) laid over the matched text's own. A built-in's own default (markers superscript, formula digits subscript) is replaced by it |
| `link`, `schemes` | `link: true` makes the match a link when it reads as an address of an allowed scheme (`https`, `http`, `mailto` by default); the `links` built-in takes `schemes` too. A trailing `.,;:!?` and a closing bracket the address did not open stay outside the link |
| `in` | Only in these kinds of text: `paragraph`, `heading`, `list`, `label` (a field's label), `value` (a field's value), `cell`, `caption`, `note` (notes, `note` and `small` paragraphs, placeholders), `signature`, `choice` (a question), `header`, `footer` (everything in a header or footer). Everywhere when absent |
| `id`, `auto` | `auto: false`: the rule applies only where a block or a recipe filter names it by `id` |

**Where.** A block's `textRules` decides which rules touch it: absent, the automatic ones; a list of ids, the
automatic ones and those; `false`, none. Blocks in a table's cells and in a pinned block take their container's
setting unless they give their own.

**Order.** A block's rules apply one after another in the order the document lists them, each to the runs the ones
before it made. A rule sees a stretch of consecutive text runs as one text (a match can run across two runs of
different looks); a link, a field, a break, a tab, a blank, a picture or a tick box ends a stretch. **Text inside a
link** (a run with `link`, given by the application or made by an earlier rule) **is never searched again**: nothing
is linked twice and an address's digits are never set low. So list `links` first. Hits of one rule that overlap are
taken left to right, the later dropped.

**Limits** (`TEXT_RULE_LIMITS`, each can be changed through `applyTextRules(document, limits)` or
`compileTextRules(rules, limits)`): patterns come from JSON a person may edit, so they are read through a safe subset
and every search is bounded.

| Limit | Default | |
|---|---|---|
| `rules` | 100 | rules in a document |
| `patternLength` | 200 | characters in a pattern |
| `repeat` | 100 | the most times `*`, `+` or `{n,}` repeat (they are read as `{0,100}`, `{1,100}`, `{n,100}`); an explicit `{n,m}` may not ask for more |
| `cost` | 12,000 | the product over a pattern's repeats of (most times + 1): two `+` (101² = 10,201) pass, three do not |
| `terms`, `termLength` | 1,000, 100 | terms in a list (and formulas), characters in one |
| `textLength` | 10,000 | characters of one stretch a pattern or term list searches; a longer stretch is left as it is and reported in `applyTextRules(…).skipped` (the built-ins, which are linear, search any length) |
| `matches` | 20,000 | matches in one document, every rule together; past it, `TextRuleError` |

Refused, each named in the error's `problems`: back-references (`\1`, `\k<…>`), look-arounds (`(?=`, `(?!`,
`(?<=`, `(?<!`), Unicode property escapes (`\p{…}`), control escapes, a repeat of a group that holds a repeat or an
alternative (`(a+)+`, `(a|b)*`, the patterns that can run away), a `{` that is not a repeat, flags other than `i`, a
rule that says neither `style` nor `link`, an unknown key or style key, and a built-in without what it needs
(`formulaDigits` without `formulas`). A worst case of what is allowed searches 10,000 characters in well under a
second (`tests/model/text-rules.test.ts`).

`compileTextRules(rules)` reads a list once; `ruleText(richText, compiled, { use: ids })` applies named rules to one
value (what the recipes' `rules` filter does).

## Spreadsheets

`spreadsheet(bytes, { sheet: 'Budget', range: 'A1:F60' })` prints one sheet of an `.xlsx` as a native Word table, so
it is set, searched and edited as the rest of the document is, and the layout breaks it across pages like any table.
The Word writer reads the workbook (`document-kit/ream`'s Excel reader) and turns the sheet into a `table` block
(`spreadsheetTable(block, { width })` from `document-kit/docx` gives that table to an application):

- **What prints**: `sheet` by tab name or 0-based position (the first visible sheet when absent); `range` in A1 form,
  else the sheet's print area, else A1 to the last cell holding a value, a fill or a border. Hidden rows and columns
  are left out.
- **The look**: the sheet's column widths (measured as Excel does, in the workbook's Normal font), merged cells,
  fonts (face, size, bold, italic, underline, strike, colour; rich text run by run), solid fills, cell borders
  (hair, thin, medium, thick, double, dotted, dashed), horizontal and vertical alignment (Excel's General alignment:
  numbers right, text left), indents, wrapped text, custom row heights (as least heights) and the grid lines when
  the sheet prints them (`gridLines` overrides). Text that does not wrap runs on into the empty cells beside it, as
  Excel prints it. No row breaks across a page.
- **Values**: number formats as Excel applies them (currency, thousands, percent, dates, coloured sections). A
  formula prints the value the file cached; a formula with none (openpyxl writes none) is evaluated from the other
  cells, on any sheet of the workbook: arithmetic, comparisons, cell and range references, defined names, shared
  formulas and most of Excel's common functions (`SUM`, `AVERAGE`, `ROUND`, `IF`, `VLOOKUP`, …). A formula that
  refers to itself prints `#REF!`. `formulas: 'evaluate'` recalculates every formula, ignoring cached values.
- **Pages**: the leading rows repeat at the top of each page (`headerRows`; else the sheet's print titles, else its
  frozen rows). A sheet set to fit its page width is scaled, text and columns together, to the section's text
  width; `scale` (percent) sets it. Columns still wider than the text width are narrowed to fit. Give the section the
  workbook's page set-up to print it as the workbook does: `{ paper: 'tabloid', orientation: 'landscape' }` is
  17 × 11 in.
- Caption (numbered with the tables), note, `align`, `spaceAfter` as a table's. `workbook: null` prints a placeholder
  box (`placeholder`, else "Spreadsheet: *sheet*") and is listed to complete (kind `spreadsheet`).

Charts, pictures, shapes, comments and conditional formats on the sheet are not printed.

## Keeping table rows together

Word keeps a row on the page of the next row when the paragraphs of the row carry keep-with-next, and does not break a
row that says `cantSplit`. The writer says both that way:

- a row's `keepWithNext: true` puts `w:keepNext` on every paragraph of its cells. A `header` row that has a row after it
  keeps with it unless it says `keepWithNext: false`, so a heading row is never left alone at the foot of a page.
  Repeating header rows (`w:tblHeader`) work as before.
- a table's `keepTogether: true` keeps every row but the last with the next and makes every row `cantSplit`. A table
  taller than a page still breaks, between rows. A row's own `keepWithNext` or `cantSplit` wins over it.

`document-kit/layout` honours both (see `docs/layout.md`).

## Theme

The theme is how a document looks, so one application can be Times with a green band and the other Arial and grey,
and the kit knows neither: body font, size and colour; title, subtitle, six heading levels, caption, note, small,
header and footer text (font, size, colour, bold, italic, caps, alignment, spacing before and after, line height as a multiple or in points; a heading level's `keepWithNext`, true by default, false to let a heading end a page);
link colour; the placeholder mark (`placeholderMark`, a text style); the table look (text size, line height `line`,
header fill and text, group fill, note text, label fill, outer and inner borders, cell padding, least heights of
header and other rows `headerRowHeight` and `rowHeight`, spacer height, placeholder box); the caption words and separator (`Table 3: Results`, or `Figure 2. View`); the blank rule's
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
value is not printed at all, and its separating space goes with it — unless it says `missing: 'placeholder'`: then
it prints with each missing slot as `[slot]`, marked. A placeholder blank (`show: 'placeholder'`) prints `[what]`,
marked. Both are listed to complete like the rest. An unanswered tick-box question (`chosen` null) and a null answer
are listed as `choice` and `answer`.

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
`metadata.created`, writes that time in `docProps/core.xml`. An embedded font's obfuscation key (random in the library)
is made from the font's bytes. The same document gives the same `word/document.xml` always, and the same bytes
whenever `created` is given.

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
| Footnote markers in cells and footnote lines | text rules `trailingMarker` (cells, labels) and `leadingMarker` (footnote lines, named by their blocks); or `sup('a')` runs |
| Chemistry (CaCO3) | text rule `formulaDigits` with the app's list of formulas; or `sub('3')` runs |
| E. coli in italics | text rule `terms` with `style: { italic: true }` |
| `align` (label/value/header), `exportLayoutPreset` density | `columnAlign`, cell `align`, table `size`, `cellPadding`, `columnWidths` |
| Wide tables: `separateTables` / `stackedBlocks` | several `table`s, or one table with `header` rows inside it (they do not repeat); the split itself stays in the app's `table-layout.ts` |
| `tableContinuation` (repeat / none / auto) | row `repeat` and table `repeatHeader: false` |
| `keepTableTitleWithTable` | always on for captions above a table |
| Captions "Table N: …", `captionMode` none / explicit | `caption`, `numbered: false` / `'silent'` |
| Figures (`layout` wide / standard / half) and placeholders | `figure` with a `fit` box; `image: null` and `placeholder` |
| Attachments (image page, companion file, external report) | `figure` with `pageBreakBefore`, or a placeholder `figure` |
| Signature block (label, picture or wet-ink gap, rule, name and title) | `signature` with `titleLayout: 'sameLine'`, `align: 'center'`, `signingSpace`, `ruleWidth`, `ruleGap`, `spaceAbove`/`spaceBelow` |
| Service Disclaimer at the foot of page 1, Contact at the foot of the last page (`pinBottom`, `CONTACT_LIFT_LINES`) | `pinned` with `spaceBelow` |
| E-mail and phone on one line when they fit | `fields` row `sameLine: true` |
| An unfilled `[placeholder]` marked in its sentence | `blank` with `show: 'placeholder'`; a sentence's `missing: 'placeholder'` |
| Header code; footer reference + "Page X of Y" | section `header` paragraph; footer paragraph with `tabs: [{ at: 'end', align: 'right' }]`, `pageNumber()`, `pageCount()` |
| Links in prose (URLs, e-mail) | text rule `links` (scheme list, trailing punctuation outside); or `link` runs |
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
| `choices` (tick boxes, boxed, answer) | `choices` (`boxed`, `answer`); hand edits of the tick and the answer in recipes |
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

- A frame that holds a table (Word's frames take paragraphs), or a frame given a fixed point on a page other than
  the one its anchor falls on.
- A floating picture cropped or framed (`crop` and `border` are on `figure` only), or text wrapped round it
  (`square`, `tight`): it floats with no wrapping, behind or in front.
- Shapes and vector drawing: a scale bar with an arrow, a mark between matrix cells, an X on a grid, a north arrow.
  The app draws them as a PNG or builds them from table cells. SVG pictures are not written.
- Tick boxes typed as characters (☐ in a text run) are drawn by the layout's face of boxes too, but only the
  `checkbox` run is a real check box in Word.
- A picture border other than a solid line (dashed, double): the layout draws a picture's frame solid only. Crop and
  border are on `figure`; a `figureGrid` item and an inline picture take neither yet.
- A table pinned to the foot of the page (Word frames hold paragraphs only).
- Fitting a `sameLine` field by measuring: the kit does not measure text, so a row that does not fit wraps inside
  the row before it's paragraph rather than becoming a paragraph of its own.
- A label printed inside a picture's corner (the `inset` placement of a figure pair); labels go above or below.
- A list of figures (the contents' Images list); only headings make a table of contents.
- Internal links and bookmarks the app names; `link` takes URIs only.
- Footnotes and endnotes as Word notes (footnote markers are superscript runs).
- Automatic splitting of a too-wide table into blocks, and automatic placing of wide tables on landscape pages.
- A custom block that opens a section of its own.
- JPEG orientation from EXIF: a rotated photo is drawn as stored.
