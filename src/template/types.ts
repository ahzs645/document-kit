/**
 * The template engine's public and shared types, taken from officeParser 8.9.0's `src/types.ts`
 * (MIT, Harsh Ankur): only what the template engine uses. The error and warning codes keep
 * officeParser's names and values so its documentation still applies; codes for parsing,
 * generation, PDF, OCR and decryption are left out, and `TEMPLATE_ENCRYPTED` is document-kit's own
 * (see docs/template-upstream.md).
 *
 * @module template/types
 */

/** Error codes a template call rejects with (`error.officeIssue.code`). */
export enum OfficeErrorType {
    /** Arguments passed to the function are missing or invalid */
    IMPROPER_ARGUMENTS = 'IMPROPER_ARGUMENTS',
    /** Input type is not a supported type (Uint8Array, ArrayBuffer or Blob) */
    INVALID_INPUT = 'INVALID_INPUT',
    /**
     * The template is password-protected (an encrypted Word file, or an OpenDocument file with
     * encrypted entries). document-kit does not decrypt templates: save the template without a
     * password. (document-kit's own code; officeParser decrypts with Node's `crypto`.)
     */
    TEMPLATE_ENCRYPTED = 'TEMPLATE_ENCRYPTED',
    /** A template render was given a document type it cannot template */
    TEMPLATE_UNSUPPORTED_FORMAT = 'TEMPLATE_UNSUPPORTED_FORMAT',
    /** A template placeholder had no matching data field and `onMissing: 'error'` was set */
    TEMPLATE_FIELD_MISSING = 'TEMPLATE_FIELD_MISSING',
    /** A template's block tags cannot be read or do not enclose content (see {@link TemplateSyntaxDetails}) */
    TEMPLATE_SYNTAX = 'TEMPLATE_SYNTAX',
    /** A template's data does not fit how a tag uses it (see {@link TemplateDataDetails}) */
    TEMPLATE_DATA_INVALID = 'TEMPLATE_DATA_INVALID',
    /** ZIP entry count exceeds limit */
    ZIP_ENTRY_COUNT_LIMIT_EXCEEDED = 'ZIP_ENTRY_COUNT_LIMIT_EXCEEDED',
    /** ZIP uncompressed size limit exceeded */
    ZIP_SIZE_LIMIT_EXCEEDED = 'ZIP_SIZE_LIMIT_EXCEEDED',
    /** A template's XML holds more elements than `decompressionLimits.maxXmlElements` */
    XML_ELEMENT_LIMIT_EXCEEDED = 'XML_ELEMENT_LIMIT_EXCEEDED',
    /** ZIP data yielded no readable entries (corrupt, truncated, or not a ZIP archive) */
    ZIP_NO_ENTRIES_FOUND = 'ZIP_NO_ENTRIES_FOUND',
    /** ZIP data is truncated: the End of Central Directory record is absent */
    ZIP_TRUNCATED = 'ZIP_TRUNCATED',
    /** Document element/structure nesting exceeded the safe recursion depth */
    MAX_NESTING_DEPTH_EXCEEDED = 'MAX_NESTING_DEPTH_EXCEEDED',
    /** The output grew past what the render may add by repetition, or past what a string can hold */
    OUTPUT_TOO_LARGE = 'OUTPUT_TOO_LARGE',
}

/** Warning codes a template call reports through `onWarning` (`issue.code`). */
export enum OfficeWarningType {
    /**
     * A template holds text that looks like a tag but is not one (see `TemplateTextReason`), which a render
     * writes as typed: `{{first name}}`, `{{ name }`, `｛｛nom｝｝`. Raised once by `fillTemplate` and
     * `checkTemplate`, with each such text, its part and why in `details.texts` (at most 1,000) and
     * how many there are in `details.found`; `templateTags` lists them as tags of kind `text`, and
     * raises it only when it found more than it lists.
     */
    TEMPLATE_NOT_A_TAG = 'TEMPLATE_NOT_A_TAG',
    /** A template check found as many problems as it lists (`details.limit`: 1,000 for one data entry, or 10,000 in all for a batch, `details.batch`) and stopped there, in `details.entry` of a batch, with `details.unchecked` entries after it not checked; fix these and check again */
    TEMPLATE_CHECK_LIMITED = 'TEMPLATE_CHECK_LIMITED',
    /**
     * A template condition ordered two texts that both read as numbers (`{{ if a > b }}` with `'250'` and
     * `'1000'`) and, as `numericTextComparison` was not given, compared them as text, character by character.
     * Raised by `fillTemplate` and `checkTemplate` once per tag (its first such comparison),
     * with the tag, the texts and whether comparing them as numbers answers otherwise in `details` (see
     * {@link TemplateNumbersAsTextDetails}). Set `numericTextComparison` to `'number'` to compare them as
     * numbers, or to `'text'` to keep comparing them as text without the warning.
     */
    TEMPLATE_NUMBERS_COMPARED_AS_TEXT = 'TEMPLATE_NUMBERS_COMPARED_AS_TEXT',
    /** A config option was passed that this version does not recognize (e.g. a key renamed in a major release); it had no effect */
    UNRECOGNIZED_CONFIG_OPTION = 'UNRECOGNIZED_CONFIG_OPTION',
    /**
     * An option was given a value it does not accept (`onMissing`, `numericTextComparison`, a
     * `decompressionLimits` entry). The option's default was used instead; the message names the
     * option, the value and what it accepts.
     */
    INVALID_CONFIG_VALUE = 'INVALID_CONFIG_VALUE',
}

/** Limits applied to decompressing and filling a template (the same shape and defaults as officeParser's). */
export interface DecompressionLimits {
    /**
     * Maximum allowed total uncompressed size (in bytes) of files extracted from a ZIP archive.
     * Default is 536870912 (512 MB).
     */
    maxUncompressedBytes?: number;
    /**
     * Maximum allowed number of entries (files and directories) in a ZIP archive.
     * Default is 10000.
     */
    maxZipEntries?: number;
    /** Not used by the template engine; accepted so an officeParser configuration carries over. Default 1000000. */
    maxTableCells?: number;
    /**
     * Maximum number of XML elements read from the template's parts (comments, CDATA sections and
     * processing instructions count too), and with them its tags and their tokens. Past the limit the
     * call fails with `XML_ELEMENT_LIMIT_EXCEEDED`. Default is 2000000.
     */
    maxXmlElements?: number;
    /** Not used by the template engine; accepted so an officeParser configuration carries over. Default 67108864. */
    maxRawContentLength?: number;
    /**
     * Maximum amount of content a render adds by repetition: each copy of a loop's content after its
     * first, each copy of a note a repeated reference points at, each use of a value after its first
     * (what the value outgrows its placeholder by), each block or field written again as typed, and each
     * field, loop or branch evaluated. Past the limit the render rejects with `OUTPUT_TOO_LARGE`.
     * Default is 16777216.
     */
    maxRepeatedContent?: number;
}

/** A single issue (a warning or an error) reported by a template call. */
export interface OfficeIssue {
    /** The severity of the issue. */
    type: 'warning' | 'info' | 'error';
    /** Human-readable message text. */
    message: string;
    /** A unique code for programmatic handling. */
    code: OfficeWarningType | OfficeErrorType;
    /** Additional context: the tag, the part, the limit, as each code describes. */
    details?: any;
}

/** Where an issue is reported: a call's `onWarning`. */
export interface IssueHandler {
    onWarning?: ((issue: OfficeIssue) => void) | undefined;
}


/**
 * A value in a template's data. Text, numbers, true/false and dates fill a field (`null`/`undefined`
 * render as empty text); a list is what `{{ for item in list }}` repeats over; an object holds named
 * values reached by a path (`{{ customer.address.city }}`).
 */
export type TemplateValue = string | number | boolean | Date | null | undefined | readonly TemplateValue[] | { readonly [key: string]: TemplateValue };

/**
 * The data for one document from {@link fillTemplate}. A key `name` fills the field
 * `{{name}}` (see `delimiters`), and a path `{{ a.b }}` reads `b` of the object at `a`; a key that
 * itself holds a dot (`"invoice.total"`) is read as typed before it is read as a path. Fields absent
 * from the data are governed by `onMissing`.
 */
export type TemplateData = Record<string, TemplateValue>;

/** Why a template's tags were rejected with `TEMPLATE_SYNTAX`. */
export type TemplateSyntaxReason =
    /**
     * A tag cannot be read; `detail` says why. Text that only starts like a block tag (`{{ If you call }}`)
     * stays text, and is named so only when the blocks around it fail to pair. Named at once: a block
     * keyword spelt as other template languages spell it (`elseif`, `elif`, `end if`), and a filter, on a
     * field or on a value in a condition, given what it does not take (`format` with a picture that cannot
     * be read, `upper "x"`).
     */
    | 'unreadable'
    /** A block's opening tag has no closing tag. */
    | 'unclosed'
    /** A closing tag, or an `else`, has no block to belong to. */
    | 'unopened'
    /** A closing tag closes a different kind of block than the one open (`endfor` for an `if`). */
    | 'mismatched'
    /** An `else` or `else if` follows the block's `else`. */
    | 'else-after-else'
    /** A block's tags sit where no content lies between them as a whole: in different tables, cells across tables, a text box and the text around it, different parts, or inside an element a block cannot be cut at; `detail` names it. */
    | 'crossing'
    /** An `else` or `else if` in a table shares a row with another of the block's tags; in a table each branch is its own rows. */
    | 'row-branch'
    /** Two blocks overlap without one holding the other. */
    | 'overlap'
    /** The part holding block tags is not well-formed XML. */
    | 'malformed-xml';

/** The `details` of a `TEMPLATE_SYNTAX` error. */
export interface TemplateSyntaxDetails {
    reason: TemplateSyntaxReason;
    /** The tag as typed, delimiters included. */
    tag: string;
    /** The document part holding it, e.g. `word/document.xml`. */
    part: string;
    /** More about the reason, where there is more to say. */
    detail?: string | undefined;
    /**
     * The tag as it should be typed, delimiters included, where the reason says how: a block keyword
     * spelt as other template languages spell it (`{{ elseif x }}`, `{{ elif x }}`) is `{{ else if x }}`,
     * and `{{ end if }}` is `{{ endif }}`. The spaces typed around the tag's text are kept.
     */
    suggestion?: string;
}

/** Why a template's data was rejected with `TEMPLATE_DATA_INVALID`. */
export type TemplateDataReason =
    /** A `for` tag's list is not a list. */
    | 'not-a-list'
    /** A field's value (or the value a case filter changes) is a list, an object, a function or a symbol, which has no text of its own. */
    | 'not-text'
    /** A condition compares two values of different kinds (text with a number), or a list or object. */
    | 'incomparable'
    /** A filter cannot write the value: a number picture given text that is not a number, a date picture given one that is not a date. */
    | 'unformattable'
    /** Text read as a number two ways of a `dataNumberSeparators` list read as different numbers (`1.250`: 1250 or 1.25); `readings` holds each. */
    | 'ambiguous-number';

/** One way of a `dataNumberSeparators` list that read a text as a number, and what it read. */
export interface TemplateNumberReading {
    /** The way's place in the list, from 0. */
    separators: number;
    /** Its decimal point and group separator (`''` for none). */
    decimal: string;
    group: string;
    /** The number it read. */
    value: number;
}

/** The `details` of a `TEMPLATE_DATA_INVALID` error. */
export interface TemplateDataDetails {
    reason: TemplateDataReason;
    /** The tag as typed, delimiters included. */
    tag: string;
    /** The document part holding it. */
    part: string;
    /** The kind of the value (`text`, `number`, `true/false`, `date`, `list`, `object`, `function`, `symbol`, `empty`). */
    kind: string;
    /** For `incomparable`, the kind of the other value. */
    otherKind?: string;
    /** For `unformattable`, the filter, and what it takes (`number` or `date` for `format`). */
    filter?: string;
    takes?: string;
    /** For `ambiguous-number`, each way that read the text and what it read. */
    readings?: TemplateNumberReading[];
}

/** The `details` of a `TEMPLATE_NUMBERS_COMPARED_AS_TEXT` warning: the first such comparison of its tag a render or check met. */
export interface TemplateNumbersAsTextDetails {
    /** The tag as typed, delimiters included. */
    tag: string;
    /** The document part holding it. */
    part: string;
    operator: '<' | '<=' | '>' | '>=';
    /** The two texts compared, as the condition had them (after their filters), each at most 80 characters (a longer one is its first 77 and `...`). */
    left: string;
    right: string;
    /** Whether comparing them as numbers gives the other answer (`'250' > '1000'`: true). Left out when a text reads as two numbers (a `dataNumberSeparators` list). */
    differs?: boolean;
    /** For a batch, the data entry it was met in, from 0. */
    entry?: number;
}

/** A loop a template tag is read in: the list as typed in its `for` tag (`lines`, `order.lines`), and the pass, from 0. */
export interface TemplateLoopPlace {
    list: string;
    index: number;
}

/**
 * Why text that looks like a tag is not one (see {@link TemplateTagInfo}, kind `text`). Exactly these,
 * and nothing else is reported (single braces, `{x}`, JSON, `{{}}` are not):
 * - `space-in-name`: text between the delimiters that is a name with white space in it, made of letters,
 *   digits, `_ . -` and white space, starting with a letter or `_`, at most 64 characters: `{{first name}}`
 *   (a field's name has no spaces; `{{first_name}}` with a data key `first_name` is one).
 * - `unknown-filter`: a field, or a value in a condition, through a word after `|` that is no filter,
 *   built-in or in `formatters`: `{{ total | fromat "0.00" }}`, `{{ if answer | lowr == "yes" }}`.
 * - `missing-delimiter`: for delimiters of one character typed twice or more (the default `{{ }}`), a
 *   name between a start delimiter and an end delimiter one of which is a character short: `{{ name }`,
 *   `{ name }}` (a short start inside braces still open, or a short end whose brace a later one closes,
 *   as code nests them, is not reported).
 * - `full-width`: a name between delimiters typed with full-width characters (U+FF01 to U+FF5E), as a
 *   Japanese or Chinese keyboard types them: `｛｛nom｝｝`.
 */
export type TemplateTextReason = 'space-in-name' | 'unknown-filter' | 'missing-delimiter' | 'full-width';

/**
 * A tag of a template, as {@link templateTags} lists it: a field, or a block's opening or
 * branch tag (`endif` and `endfor` close what they open and are not listed); or (`kind: 'text'`) text
 * that looks like a tag but is not one, which a render writes as typed. In the order of each part.
 */
export interface TemplateTagInfo {
    /**
     * `field` (`{{ name }}`), `if`, `elseif`, `else` or `for`; or `text`: text that looks like a tag but is
     * not one (`{{first name}}`, `{{ name }`, `｛｛nom｝｝`), written as typed, with why in `reason`, no
     * `fields`, and the loops it is inside. At most 1,000 such texts are listed for a template; when there
     * are more, a `TEMPLATE_NOT_A_TAG` warning says how many (`details.found`).
     */
    kind: 'field' | 'if' | 'elseif' | 'else' | 'for' | 'text';
    /** The tag as typed, delimiters included (a very long one is cut short); for `text`, the text as typed. */
    tag: string;
    /** The document part holding it, e.g. `word/document.xml` or `word/header1.xml`. */
    part: string;
    /**
     * What the tag reads from the data, each as typed: a field's path, every path in a condition, a
     * loop's list. A path may start with a loop's item (`line.qty`) or `loop` (`loop.last`), which the
     * enclosing `for` tags name.
     */
    fields: string[];
    /**
     * A field's filters, by name, in order; for an `if` or `else if`, the filters its condition's values go
     * through, by name, each once, in the order first written.
     */
    filters?: string[];
    /** A `for` tag's item name (`line` in `{{ for line in lines }}`). */
    item?: string;
    /** For `text`: why it is not a tag. */
    reason?: TemplateTextReason;
    /**
     * The `for` tags this tag is inside, outermost first, each as its item name and list as typed. A note
     * only loops refer to is listed as its part holds it, outside them: its fields may start with the
     * item of a loop that refers to it.
     */
    loops: { item: string; list: string }[];
}

/** A problem {@link checkTemplate} finds with a template's data: what a render with `onMissing: 'error'` would reject. */
export interface TemplateProblem {
    /** `TEMPLATE_FIELD_MISSING` (the data does not have a field the template reads) or `TEMPLATE_DATA_INVALID` (a value does not fit its tag). */
    code: OfficeErrorType.TEMPLATE_FIELD_MISSING | OfficeErrorType.TEMPLATE_DATA_INVALID;
    /** What a render would say. */
    message: string;
    /** For a missing field: the field as typed (`name`, `line.qty`). */
    field?: string;
    /** For a value that does not fit: why, and the tag (see {@link TemplateDataDetails}). */
    details?: TemplateDataDetails;
    /** The document part the tag is in. */
    part: string;
    /** The loops it was read in, outermost first, with each one's pass: `line.qty` in the third line is `[{ list: 'lines', index: 2 }]`. Frozen, and shared by the problems of one pass. */
    loops: TemplateLoopPlace[];
}

/**
 * A filter of the caller's own, named in `TemplateConfig.formatters` and written in a tag after `|`
 * (`{{ name | initials }}`, `{{ total | round 2 }}`). It is given the field's value as the data holds it
 * (a list or an object too) and the strings and numbers written after its name, and returns the value
 * to write (or to pass to the next filter).
 */
export type TemplateFormatter = (value: TemplateValue, ...args: (string | number)[]) => TemplateValue;

/** The words a date picture writes: month and day names, in full and short, and the AM/PM markers. */
export interface TemplateDateNames {
    /** Twelve names, January first (`MMMM`). */
    months?: string[];
    /** Twelve short names (`MMM`). */
    monthsShort?: string[];
    /** Seven names, Sunday first (`dddd`). */
    days?: string[];
    /** Seven short names (`ddd`). */
    daysShort?: string[];
    /** The markers `AM/PM` writes (`am/pm` writes them in lower case, `A/P` their first letter). */
    am?: string;
    pm?: string;
}

/**
 * Options for {@link fillTemplate}. `data` is the only required field: a single
 * {@link TemplateData} produces one document, an array of them produces one document per entry
 * (a mail-merge). An option render does not take, at the top or inside `delimiters`,
 * `numberSeparators`, `dateNames` or `decompressionLimits`, is reported (`UNRECOGNIZED_CONFIG_OPTION`).
 */
export interface TemplateConfig {
    /** The field values. One object -> one rendered document; an array -> one document per entry. */
    data: TemplateData | TemplateData[];
    /**
     * Tag delimiters. Defaults to `{{` and `}}`, so `{{name}}` is a field and `{{ if paid }}` a block
     * tag. A tag is matched even when the source splits it across several runs (a common Word quirk).
     * Delimiters that are not an object, or one that is not text, are refused (`IMPROPER_ARGUMENTS`).
     */
    delimiters?: { start: string; end: string };
    /**
     * What to do with a field the data does not have, in a placeholder, a condition or a loop's list:
     * - `'keep'` (default): leave it as typed. A placeholder keeps its text; a block (`if` or `for`)
     *   whose condition or list uses the field is left exactly as typed, tags and content included.
     * - `'empty'`: treat it as empty. A placeholder is replaced with nothing; in a condition the field
     *   is `null`; a loop over it repeats nothing.
     * - `'error'`: reject with `TEMPLATE_FIELD_MISSING`.
     * A key that is present but `null`/`undefined` is not missing: it renders as empty, is `null` in a
     * condition, and a loop over it repeats nothing. Another value is reported (`INVALID_CONFIG_VALUE`)
     * and `'keep'` used.
     */
    onMissing?: 'keep' | 'empty' | 'error';
    /**
     * The separators a number picture (`{{ total | format "#,##0.00" }}`) writes. A picture is always
     * written with `,` for groups and `.` for the decimal point, as in Word; these are what appear in
     * the document. Default `{ decimal: '.', group: ',' }` (`{ decimal: ',', group: '.' }` writes
     * `1.250,50`). A value that is not an object of text is refused (`IMPROPER_ARGUMENTS`).
     */
    numberSeparators?: { decimal?: string; group?: string };
    /**
     * How a number written as text in the data is read, by a number picture, where a condition compares it
     * with a number, and where two texts are compared under `numericTextComparison`: its decimal point and the separator of its groups of three digits.
     * `{ decimal: ',', group: '.' }` reads `12,50` and `1.234,50`; `{ decimal: ',', group: ' ' }` reads
     * `1 234,50` (a space separator matches a no-break space too). Each is one character; `group` may be
     * `''` (the default when left out) for none, and `decimal` defaults to `.`. A group separator where
     * groups of three do not follow makes the text no number (`1.5` with `.` for groups), nor is one
     * after a first group starting with 0 (`0.500`), so an ambiguous text is refused rather than misread. Not given (the default), text is read with `.` for the point and
     * no group separator (`1250.5`). `numberSeparators` says how numbers are written, this how they are
     * read, so data written one way can fill a document written another.
     *
     * A list gives several ways, each tried: `[{ decimal: ',', group: '.' }, { decimal: '.' }]` reads `12,50`,
     * `1.234,50` and `12.50`. Text read by more than one way is taken when they read the same number (`1250`),
     * and refused when they differ (`1.250`: 1250 one way, 1.25 the other) with `TEMPLATE_DATA_INVALID`
     * (`reason: 'ambiguous-number'`, each reading in `details.readings`); so such a list refuses a whole number
     * written with one group separator (`1.000`), which `{ decimal: ',', group: '.' }` alone reads as 1000. One
     * to 8 ways; one object, or a list of one, reads as before. A value of the wrong shape is refused
     * (`IMPROPER_ARGUMENTS`).
     */
    dataNumberSeparators?: { decimal?: string; group?: string } | { decimal?: string; group?: string }[];
    /**
     * How a condition compares two texts that both read as numbers, as `dataNumberSeparators` reads them
     * (by default with `.` for the point: `'250'`, `'-3.5'`, `'1e3'`), such as two columns of a spreadsheet,
     * where every value is text:
     * - not given (the default): as text, character by character (`'250' > '1000'` is true), and an order
     *   comparison (`<`, `<=`, `>`, `>=`) of two such texts is reported, once per tag in a render or a check,
     *   with a `TEMPLATE_NUMBERS_COMPARED_AS_TEXT` warning;
     * - `'text'`: as text, without the warning;
     * - `'number'`: as numbers, for every operator, exactly: `'250' < '1000'`, `'007' == '7'` and `'1.0' == '1'`
     *   are true, and two 20-digit ids that differ in their last digit stay different. Text too large for
     *   a double (past about 1.8e308: `'1e400'`, a whole number of 310 digits) is no number, as for a number
     *   picture, so it compares as text; a very small one (`'1e-400'`) is a number, and compared exactly.
     *   Text that a `dataNumberSeparators` list reads as two different numbers is refused (`ambiguous-number`).
     * Text beside text that is no number compares as text. Text compared with a number typed in the tag
     * (`{{ if total > 100 }}`) is read as a number whatever this says. Another value is reported
     * (`INVALID_CONFIG_VALUE`) and the default used.
     */
    numericTextComparison?: 'text' | 'number';
    /**
     * The month and day names and AM/PM markers a date picture writes. Default English. A value that is
     * not an object of names, or a list of the wrong length, is refused (`IMPROPER_ARGUMENTS`).
     */
    dateNames?: TemplateDateNames;
    /**
     * Minutes added to UTC before a date picture writes a date: 0 (the default) writes dates in UTC, 330
     * in India's time, -300 in New York's in winter. A fixed offset, the same on every machine, unlike
     * the machine's own time zone.
     */
    utcOffsetMinutes?: number;
    /** Filters of the caller's own, by name, beside the built-in `format`, `upper` and `lower` (see {@link TemplateFormatter}). */
    formatters?: Record<string, TemplateFormatter>;
    /**
     * Optional hint for the template's format: a Word document (`docx`), or an OpenDocument text (`odt`)
     * or text template (`ott`, rendered as an `.odt`). The format is always detected from the bytes, so
     * this hint does not influence detection: it only refines the `TEMPLATE_UNSUPPORTED_FORMAT` message
     * when an unsupported file is passed.
     */
    fileType?: 'docx' | 'odt' | 'ott';
    /**
     * Limits on decompressing the (untrusted) template zip, same shape and defaults as the parser's.
     * Guards against zip-bomb templates. Defaults to 512 MiB / 10000 entries. `maxXmlElements` (2,000,000)
     * bounds the elements the template's parts hold, as for parsing, and its tags and their tokens with them (`XML_ELEMENT_LIMIT_EXCEEDED`). `maxRepeatedContent`
     * (16 MiB) bounds what a document adds by repetition: each copy of a loop's content after its first,
     * each copy of a note a repeated reference points at, each use of a value after its first (what the
     * value outgrows its placeholder by), each block or field written again as typed, and each field, loop
     * or branch evaluated (a few characters, plus its condition's size). Past the
     * limit the render rejects with `OUTPUT_TOO_LARGE` (nested loops over a long list, or a million
     * placeholders of one value, would otherwise write without bound).
     */
    decompressionLimits?: DecompressionLimits;
    /**
     * Optional callback for non-fatal issues. Also suppresses the console fallback that would
     * otherwise print a thrown error, so a caller handling the rejection is not double-notified.
     */
    onWarning?: (issue: OfficeIssue) => void;
}

/**
 * A structural stand-in for the web `Blob`/`File`, so a template can be passed as one without the DOM
 * library in this package's types. Any object with an `arrayBuffer()` method qualifies.
 */
export interface BlobLike {
    arrayBuffer(): Promise<ArrayBuffer>;
}
