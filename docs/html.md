# document-kit/html

`writeHtml(document, options)` writes the block model as one self-contained HTML page, styled for print: the
same document `writeDocx` writes, read the same way, so a print view in the browser looks like the Word file and
the PDF `document-kit/layout` makes from it.

```ts
import { writeDocx } from 'document-kit/docx'
import { writeHtml } from 'document-kit/html'
import { layoutDocx } from 'document-kit/layout'

const fonts = [
  { family: 'Brand Sans', bytes: regular },
  { family: 'Brand Sans', weight: 700, bytes: bold },
]
const docx = await writeDocx(doc, { embedFonts: fonts })     // the Word file, the faces inside it
const laid = await layoutDocx(docx, { fonts: standIns, namedFonts: fonts })
const pdf = laid.toPdf()                                      // the PDF and its preview
const html = writeHtml(doc, { fonts })                         // the print view, a string
```

## What it reads, as the Word writer reads it

- Custom blocks are expanded through `options.registry`, the document's text rules applied, captions numbered in
  reading order, footnotes numbered by first use, `sourceStyles` marks laid on, `nullCells` honoured.
- The look comes from the same data: the theme and the named styles (`document.styles`) are resolved once
  (`styleTable` in `src/model/styles.ts`) and written as CSS classes, `.dk-ps-<id>` for each paragraph style
  (`Normal`, `Title`, `Heading1`…, `Caption`, `KitNote`, `KitSmall`, `Header`, `Footer`, `KitTable` and every
  defined style, each with its look merged through the styles it is based on) and `.dk-cs-<id>` for each character
  style. A run's own look (font, size, colour, weight, slant, underline, strike, raised and lowered, caps, small
  caps, highlight, tracking as `letter-spacing`) is written on its span.
- Fonts: `options.fonts` (the `NamedFont`s `writeDocx` embeds and `layoutDocx` draws with) become `@font-face` rules
  with the font inline as base64, each face at the weight of the slot it fills (400 or 700), so the browser picks
  the face Word and the layout pick. A face is carried only when `fontEmbeddingReport(fonts, 'html')` allows it:
  TrueType or OpenType (CFF), with installable, editable or preview & print embedding; a restricted or bitmap-only
  face is left out. Every font name is written as a stack: the face, its bundled stand-in (`Carlito` for Calibri…,
  as `document-kit/layout` substitutes), then the generic kind.
- Pictures are inline (`data:` URLs). Nothing in the page refers outside it; links the document holds stay links.
- `metadata` becomes `<title>` and `author`, `description`, `keywords` meta tags; each custom property a
  `<meta name="…" content="…">`.

## Pages, approximated

The browser breaks the lines, not the layout engine, so pages are approximated with CSS:

- Each section is a run of **sheets** of its paper size (`@page dk-s<n> { size: … }` and `page: dk-s<n>`, so a
  printed section takes its own paper), a new sheet at each `pageBreak` or `pageBreakBefore` in its body. The
  margins are the sheet's padding; the header sits `margins.header` from the top, the footer `margins.footer` from
  the foot; a section's first sheet takes `firstPage`'s header and footer; a section that gives none keeps the
  section before it's, as in Word.
- `page.verticalAlign` `center` or `bottom` sets the body in the middle or at the foot of the sheet.
- Frames and floating pictures are placed on their sheet with `position: absolute`, against the page or the
  margins as the block says; a floating picture behind the text sits under it (`z-index: -1` in the sheet's own
  stacking context), one in front over it. Those from a header are placed on every sheet the header is on.
- Pinned blocks and footnotes go at the foot of the sheet they fall on, the notes under a short rule.
- Page fields count sheets: `page` prints the sheet's number in its section's numbering (`start`, `format`),
  `pages` the number of sheets, `sectionPages` the section's. A contents list prints each heading's sheet number.
- Tab stops are honoured for one line: the text after each tab is placed at its stop (left, centred or right; the
  "Page 2 of 9" at the right edge of a footer). Without stops a tab advances to the next half inch.
- Paragraph spacing adds as Word adds it (no collapsing margins: every flow is a flex column), a line multiple is
  set on 1.2 em (about a face's own line), and the first block of a sheet, a frame or a header takes no space above.
- Every `style` attribute is escaped as an attribute value (`styleAttr` in `src/html/css.ts`): a font stack's quoted
  names, or a font named with quotes, no longer end the attribute and lose what follows. Font names in the style
  sheet are CSS strings with quotes, backslashes and angle brackets escaped, and the style sheet holds no `<`, so no
  name can close it.

## Known gaps

- **Sheets do not overflow onto the next.** A sheet whose text does not fit grows on screen; in print it runs on
  to another sheet of paper with no header or footer, and page numbers after it are a page short. Break the body
  where Word breaks it (`pageBreak`), or use the PDF where pagination matters: the PDF is the layout's, the HTML the
  browser's.
- **Text does not flow round a frame.** A frame's `wrap: 'around'` or `notBeside` is laid over the text, not round
  it; use frames in room the flow leaves empty (a cover, a margin) or `wrap: 'through'`.
- **Line breaks and line heights are the browser's.** With the same faces the lines are close to the PDF's but not
  identical; with no faces given, the browser's own stand-in is used.
- **Tabs** beyond the paragraph's stops, and stops in a paragraph that wraps, are approximated.
- **Table borders** collapse as CSS collapses them (the wider wins), not by Word's precedence rules.
- **Repeating header rows** repeat in print (`thead`) only when the browser breaks the table itself.
- Word's paragraph shading and borders between paragraphs of the same look (`w:between`) are drawn per paragraph.
