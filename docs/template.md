# Word templates

`document-kit/template` fills a Word document's tags from data. You write the document in Word as you
want it to look, type tags where the data goes, and the kit copies the document and replaces only the
tags. Fonts, colours, tables, headers, page layout: everything else stays exactly as you made it.

OpenDocument text files (`.odt`, `.ott`, from LibreOffice Writer) work the same way.

The engine is officeParser's (see `docs/template-upstream.md`), so officeParser's own documentation of
`OfficeTemplate` also applies.

## Writing a template

Type tags as ordinary text, with two curly braces on each side. Spaces inside the braces do not matter:
`{{name}}` and `{{ name }}` are the same tag.

### Fields

```text
Dear {{ contact }},
Results for {{ site }}, received {{ received }}.
```

- A field is replaced by the value of the same name in the data.
- **The value takes the formatting of the tag.** Make `{{ total }}` bold and the total prints bold. If
  only part of a tag is formatted, the formatting where the tag *starts* is used.
- Word sometimes splits what you typed into several pieces behind the scenes (after a spelling check or
  an edit in the middle). That does not matter: the tag is still found.
- A name is letters, digits, `_`, `-` and `.`, with no spaces. `{{ first name }}` is not a field (see
  "Text that looks like a tag" below); write `{{ first_name }}`.
- **Data inside data**: `{{ client.address.city }}` reads `city` inside `address` inside `client`.
  `{{ samples.0.id }}` reads the first item of a list. A data key that itself contains a dot
  (`"report.date"`) is read as typed first.
- A line break in a value (`\n`) becomes a line break in the document.

### When a field has no value

What happens is set by the program filling the template (`onMissing`):

| `onMissing` | A field the data does not have |
|---|---|
| `'keep'` (the default) | is left as typed, `{{ lot }}` and all, so you can see what was not filled. A condition or loop that uses it is left as typed too. |
| `'empty'` | prints nothing. In a condition it counts as false; a loop over it repeats nothing. |
| `'error'` | stops the fill with `TEMPLATE_FIELD_MISSING`, naming the field. |

A value that is present but empty (`null`) always prints nothing, whatever `onMissing` says.

### Conditions

Keep or drop text depending on the data:

```text
Status: {{ if approved }}approved{{ else }}pending{{ endif }}.
```

Or whole paragraphs, with each tag in a paragraph of its own:

```text
{{ if remarks }}
Remarks: {{ remarks }}
{{ else if urgent }}
Urgent: no remarks yet.
{{ else }}
No remarks.
{{ endif }}
```

- **Where the tags sit decides what they cover.** Both tags in one paragraph: the text between them.
  Tags in paragraphs of their own: the paragraphs between them, and the tag paragraphs disappear, so no
  blank line is left behind. Tags in different cells of a table: whole rows (see Tables).
- A condition is a field (`{{ if approved }}`) or a comparison: `==`, `!=`, `<`, `<=`, `>`, `>=`,
  combined with `and`, `or`, `not` and brackets. Compare with numbers (`total > 100`), text in quotes
  (`status == "final"`), `true`, `false` or `null`.
- These count as false: `false`, `null`, empty text, `0`, and an empty list.
- Text in the data that is a number (`"250"`) compares with a number in the tag as that number. Two texts
  compare as text, letter by letter, even if both look like numbers, so `"250" > "1000"` is true; the
  program can set `numericTextComparison: 'number'` to compare them as numbers.
- `elseif`, `elif` and `end if` are not tags here: write `else if` and `endif`. The error says so.
- Word's curly quotes are fine.

### Loops and tables

Repeat text once per item of a list:

```text
{{ for sample in samples }}
{{ sample.id }}: {{ sample.result }}
{{ endfor }}
```

- The word after `for` (`sample`) is the name you use for the current item inside the loop.
- Inside a loop, `loop.number` counts from 1, `loop.index` from 0, and `loop.first`, `loop.last` and
  `loop.length` describe the pass. A list on one line:
  `{{ for t in tags }}{{ t }}{{ if not loop.last }}, {{ endif }}{{ endfor }}` prints `red, blue`.
- Loops and conditions can be nested.

**A table row per item**: put the `for` tag at the start of the row's first cell and `endfor` at the end
of its last cell. Leave the heading row outside the loop.

| Sample | Result |
|---|---|
| `{{ for s in samples }}{{ s.id }}` | `{{ s.result }}{{ endfor }}` |

The row is repeated once per sample, with the row's own formatting. An empty list leaves only the
heading row. Close the loop in a later row and that group of rows repeats together. A condition across
cells keeps or drops whole rows the same way, each branch starting a row of its own.

### Formatting numbers and dates

Write a value through `format` and a picture, as Word's own mail merge does:

```text
{{ total | format "#,##0.00" }}            1,234,567.89
{{ total | format "$#,##0.00" }}           $1,234,567.89
{{ change | format "0.0;(0.0)" }}          (3.5)   for a negative value
{{ share | format "0.0%" }}                12.5%
{{ issued | format "d MMMM yyyy" }}        5 March 2026
{{ issued | format "dd/MM/yyyy HH:mm" }}   05/03/2026 14:30
{{ site | upper }}                         EXAMPLE FIELD STATION
```

- **Numbers**: `0` is a digit always printed, `#` a digit printed when there is one, `,` groups thousands,
  `.` is the decimal point, `%` makes a percentage. Up to three parts: positive, negative and zero,
  separated by `;`. Text in `'single quotes'` is printed as is. Values are rounded half away from zero, as
  Word and Excel do. The picture always uses `,` and `.`; the program decides what is printed for them
  (`numberSeparators`, for example `1 234,50`).
- **Dates**: `d dd ddd dddd` (day), `M MM MMM MMMM` (month: capital M), `yy yyyy`, `H HH` (24-hour),
  `h hh` with `AM/PM` (12-hour), `m mm` (minutes: small m), `s ss`. Excel's `dd/mm/yyyy` is `dd/MM/yyyy`
  here.
- A date in the data is a date, or text such as `2026-03-05` or `2026-03-05T14:30:00Z`. Dates are printed
  in UTC unless the program sets an offset (`utcOffsetMinutes`). **A date written without a time is
  midnight UTC**, so with an offset west of Greenwich (Vancouver is −480) it prints as the day before.
  Send dates without times with no offset, or as full times.
- Month and day names are English unless the program gives others (`dateNames`).
- `upper` and `lower` change case; filters chain left to right: `{{ code | lower | upper }}`. The program
  can add filters of its own (`formatters`), used the same way: `{{ name | initials }}`. The sentence
  filters (below) are there too.
- A number picture on text that is not a number, or a date picture on something that is not a date,
  stops the fill with `TEMPLATE_DATA_INVALID`.

### Sentence filters

Four filters help a sentence read right around what fills it. They work the same in a Word template and in
a JSON recipe, and chain with the others:

```text
{{ count | plural "sample was" "samples were" | capitalize }} received.     Samples were received.
{{ n }} {{ n | plural "bottle" }}                                         1 bottle / 3 bottles
{{ names | prose }}                                                       A, B and C
{{ names | prose "or" }}                                                  A, B or C
{{ ids | ranges | prose }}                                                S-1–S-3 and S-7
```

| Filter | What it does |
|---|---|
| `capitalize` | A plain lower-case first word gets a capital ("three" → "Three"), for a fill that opens a sentence. A word cased on purpose ("pH", "mL") and a value that is one letter alone (a footnote marker) are left as they are. |
| `plural "one" "many"` | The first word for a count of exactly one, the second otherwise (the first with an `s` when the second is left out). The count is a number, a number written as text, or a list (its length). Anything else stops the fill with an error naming the filter. |
| `prose "and" ", "` | A list joined as a sentence says it: "a", "a and b", "a, b and c". The first argument is the last joining word (`and` by default), the second what goes between the others (`, ` by default). Empty items are left out. A value that is not a list is printed as it is. |
| `ranges "–" 3` | A list with each run of consecutive values collapsed to its first and last: numbers (`1, 2, 3, 5` → `1–3`, `5`) or texts that differ only in a number at their end (`S-1`, `S-2`, `S-3`). The first argument is the dash (`–` by default), the second the shortest run collapsed (3 by default; shorter runs are listed). It gives a list, in the order given, for `prose` to join. |

A program's own filter of the same name (`formatters`) is used instead of the kit's.

### Headers, footers and the rest

Tags work in the body, headers, footers, footnotes, endnotes, comments and text boxes. A tag cannot
start in one of those and end in another. Deleted text kept by Track Changes is ignored, so turn Track
Changes off, or accept your changes, before saving the template.

### Text that looks like a tag

Text that is almost a tag is left exactly as typed, and the program is warned (`TEMPLATE_NOT_A_TAG`) so it
can tell you:

- a name with a space: `{{ first name }}`
- a misspelt filter: `{{ total | fromat "0.00" }}`
- a brace missing: `{{ name }` or `{ name }}`
- braces typed full-width, as some keyboards do: `｛｛name｝｝`

Ordinary braces in text, `{like this}`, are never reported.

### Mistakes in the template

These stop the fill with `TEMPLATE_SYNTAX`, naming the tag and the part of the document it is in:

- an `if` or `for` that is never closed, or an `endif`/`endfor` with nothing to close;
- `endfor` closing an `if`, or the reverse;
- an `else` after the `else`;
- a block whose tags cannot enclose content together: one in a table and the other outside it, one in a
  header and the other in the body, one inside a text box or a tracked change and the other outside.

`templateTags` (below) finds all of these before any data is used.

## Using it from a program

Everything runs in the browser or a worker: no server, no network, no file paths. The template goes in as
bytes (`Uint8Array`, `ArrayBuffer`) or a `Blob`/`File`, and the filled document comes back as bytes.

```ts
import { checkTemplate, fillTemplate, fillTemplates, templateTags } from 'document-kit/template';

// What the template asks for, without any data: to build a form, or to check a template on upload.
const tags = await templateTags(file);           // file: a File from <input type="file">, or bytes
// [{ kind: 'field', tag: '{{ contact }}', part: 'word/document.xml', fields: ['contact'], loops: [] },
//  { kind: 'for', tag: '{{ for s in samples }}', fields: ['samples'], item: 's', ... }, ...]

// What is wrong with the data, all at once, without writing anything.
const problems = await checkTemplate(file, { data });
// [{ code: 'TEMPLATE_FIELD_MISSING', field: 's.result', part: 'word/document.xml',
//    loops: [{ list: 'samples', index: 2 }], message: "Template field 's.result' has no matching value ..." }]

// One document.
const bytes = await fillTemplate(file, { data, onMissing: 'error' });   // Uint8Array, a .docx

// One document per entry (a mail-merge); the template is read once.
const docs = await fillTemplates(file, { data: [first, second] });
```

| Function | Returns |
|---|---|
| `fillTemplate(template, { data, ...options })` | `Promise<Uint8Array>`: the filled document (`.docx`, or `.odt` for an OpenDocument template) |
| `fillTemplates(template, { data: [...], ...options })` | `Promise<Uint8Array[]>`, one per entry |
| `templateTags(template, options?)` | `Promise<TemplateTagInfo[]>`: each field, `if`, `else if`, `else` and `for` tag, and text that looks like a tag (`kind: 'text'`, with `reason`) |
| `checkTemplate(template, { data, ...options })` | `Promise<TemplateProblem[]>`: what a fill with `onMissing: 'error'` would stop for; empty when the data is enough |
| `checkTemplates(template, { data: [...], ...options })` | `Promise<TemplateProblem[][]>`, one list per entry |

A check follows the data as a fill does: only the branch a condition picks is checked, every pass of
every loop is, and nothing inside a missing list is checked until the list is given.

### Options

The names are officeParser's.

| Option | Default | |
|---|---|---|
| `onMissing` | `'keep'` | `'keep'`, `'empty'` or `'error'` (see "When a field has no value") |
| `delimiters` | `{ start: '{{', end: '}}' }` | Other tag marks, such as `<<` and `>>` |
| `numberSeparators` | `{ decimal: '.', group: ',' }` | What a number picture prints for its point and thousands |
| `dataNumberSeparators` | `.` for the point, no groups | How numbers written as text in the data are read: `{ decimal: ',', group: '.' }` reads `1.234,50`; a list tries several ways and refuses text they read differently |
| `numericTextComparison` | text, with a warning | `'number'` compares two texts that are numbers as numbers; `'text'` as text, without the warning |
| `dateNames` | English | `months`, `monthsShort`, `days`, `daysShort` (Sunday first), `am`, `pm` |
| `utcOffsetMinutes` | `0` | Minutes added to UTC before a date is printed (−480 for Pacific Standard Time) |
| `formatters` | none | Filters of your own: `{ initials: (value) => ... }`. One named `capitalize`, `plural`, `prose` or `ranges` is used in place of the kit's; `format`, `upper` and `lower` cannot be replaced |
| `fileType` | | `'docx'`, `'odt'` or `'ott'`: only changes the wording of the unsupported-format error; the format is read from the bytes |
| `decompressionLimits` | 512 MB, 10,000 entries, 2,000,000 XML elements, 16 MiB repeated | Bounds for an untrusted template: `maxUncompressedBytes`, `maxZipEntries`, `maxXmlElements`, `maxRepeatedContent` (what loops and repeated values may add; past it the fill stops with `OUTPUT_TOO_LARGE`) |
| `onWarning` | | Called with each warning (`TEMPLATE_NOT_A_TAG`, `TEMPLATE_NUMBERS_COMPARED_AS_TEXT`, `UNRECOGNIZED_CONFIG_OPTION`, `INVALID_CONFIG_VALUE`, `TEMPLATE_CHECK_LIMITED`) and each error before it is thrown. Without it, warnings are not reported anywhere |

An option the kit does not know is reported as `UNRECOGNIZED_CONFIG_OPTION` and has no effect.

### Errors

Every refusal is a `TemplateError`. Its `code` (also `error.officeIssue.code`, as in officeParser) says
what happened, and `error.officeIssue.details` says where:

| Code | When |
|---|---|
| `TEMPLATE_FIELD_MISSING` | A field has no value and `onMissing` is `'error'` (`details` is the field) |
| `TEMPLATE_SYNTAX` | The template's tags are wrong (`details.reason`, `tag`, `part`, and for `elseif` and the like a `suggestion`) |
| `TEMPLATE_DATA_INVALID` | A value does not fit its tag: a loop over something that is not a list, a list printed as a field, a format on the wrong kind of value |
| `TEMPLATE_UNSUPPORTED_FORMAT` | Not a `.docx`, `.odt` or `.ott` |
| `TEMPLATE_ENCRYPTED` | The template is password-protected. The kit does not decrypt templates: save it without a password |
| `OUTPUT_TOO_LARGE`, `XML_ELEMENT_LIMIT_EXCEEDED`, `ZIP_*`, `MAX_NESTING_DEPTH_EXCEEDED` | A limit in `decompressionLimits`, or a damaged file |
| `IMPROPER_ARGUMENTS`, `INVALID_INPUT` | The call itself: data that is not an object (or a list for the batch forms), a malformed option, a template that is not bytes or a Blob |

```ts
import { isTemplateError, OfficeErrorType } from 'document-kit/template';

try {
  await fillTemplate(file, { data, onMissing: 'error' });
} catch (e) {
  if (isTemplateError(e, OfficeErrorType.TEMPLATE_FIELD_MISSING)) showMissing(e.officeIssue.details);
  else throw e;
}
```

### Same bytes every time

The same template and data always give byte-identical output, on any machine and in any time zone: every
file in the package is stamped 2001-01-01 00:00, and numbers and dates are formatted by the kit, not by
the browser's locale.
