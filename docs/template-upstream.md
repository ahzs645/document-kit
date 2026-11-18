# The template engine's upstream: officeParser

`src/template` is officeParser's template engine (`OfficeTemplate`), moved into document-kit to run in a
browser. It is ours to change; this file records what was taken and every change made, so a later
officeParser release can be compared and brought in by hand.

| | |
|---|---|
| Project | officeParser, <https://github.com/harshankur/officeParser> |
| Version | 8.9.0 (`package.json`; CHANGELOG entry dated 2026-11-18), from a downloaded source tree with no Git history, so no commit hash is recorded |
| Licence | MIT, Copyright (c) 2019 Harsh Ankur (`NOTICE-OFFICEPARSER.md`) |
| Taken | 2026-11-18 |

## Files taken

| officeParser | document-kit | What |
|---|---|---|
| `src/template/templateEngine.ts` | `src/template/templateEngine.ts` | Whole file |
| `src/template/templateTags.ts` | `src/template/templateTags.ts` | Whole file |
| `src/template/templateFormat.ts` | `src/template/templateFormat.ts` | Whole file |
| `src/template/odtTemplate.ts` | `src/template/odtTemplate.ts` | Whole file |
| `src/OfficeTemplate.ts` | `src/template/officeTemplate.ts` | Whole file, the class turned into functions |
| `src/types.ts` | `src/template/types.ts` | The template types (`TemplateValue` to `TemplateConfig`), `DecompressionLimits`, `OfficeIssue`, `BlobLike`, and the error and warning codes the engine raises |
| `src/utils/errorUtils.ts` | `src/template/errors.ts` | The template's error and warning messages, `shownTo80`, `getOfficeError`, `buildOfficeError`, `logWarning` and the reporting they share |
| `src/utils/zipUtils.ts` | `src/template/zip.ts` | `extractFiles`, `maxUncompressedBytesOf` |
| `src/utils/configUtils.ts` | `src/template/defaults.ts` | `decompressionLimitsOf`, `PROTOTYPE_POLLUTION_KEYS` |
| `src/defaults.ts` | `src/template/defaults.ts` | `DEFAULT_TEMPLATE_CONFIG`, and the parser's `decompressionLimits` defaults (as `DEFAULT_DECOMPRESSION_LIMITS`) |
| `src/utils/sanitize.ts` | `src/template/sanitize.ts` | `escapeXml`, `stripInvalidXmlChars` |
| `src/crypto/cfb.ts` | `src/template/cfb.ts` | Whole file, to recognise an encrypted Word file |
| `src/crypto/ooxmlCrypto.ts` | `src/template/cfb.ts` | `isEncryptedOoxml` only |
| `LICENSE` | `NOTICE-OFFICEPARSER.md` | The licence text |

Nothing else was taken: none of officeParser's parsers, generators, converter, PDF, OCR, CLI or
decryption code. `src/template/utf8.ts` and `src/template/index.ts` are document-kit's own.

## Changes

### Runs in a browser

- **Input is bytes, never a path.** `TemplateInput` is `Uint8Array | ArrayBuffer | BlobLike`. The path
  branch of `readInput` (`assertNode`, `import('fs')`, `FILE_DOES_NOT_EXIST`) is gone, and with it
  `src/utils/envUtils.ts`. A `Uint8Array` is used as given, without a copy, as officeParser's `Buffer`
  view of it was.
- **No `Buffer`.** Every `Buffer.from(bytes).toString('utf8')` is `utf8Text(bytes)` and every
  `Buffer.from(text, 'utf8')` is `utf8Bytes(text)` (`utf8.ts`). The decoder is
  `new TextDecoder('utf-8', { ignoreBOM: true })`: a `TextDecoder` drops a byte-order mark by default and
  Node's `Buffer` keeps it, so this keeps a part's text exactly as officeParser read it. In `zip.ts`,
  `Buffer.concat`, `Buffer.from(chunk)` and `lastIndexOf` are a small concat, `chunk.slice()` and a byte
  search.
- **No abort signal in the unzip.** `extractFiles` took the parser's `config.abortSignal` and yielded
  event-loop turns (`setImmediate`, `MessageChannel`, timers) while a signal was given. A template call
  takes no signal (`TemplateConfig` never had one), so that path was dead here and is removed, with
  `eventLoopTurns`, `checkAbortSignal` and `getAbortError`. Limits, errors and their order are unchanged.
- No `node:` import, `fs`, `process`, `require`, `window` or `document` remains in `src/template`. The
  one global used beyond the language is `TextDecoder`/`TextEncoder` (in every browser, worker and Node),
  and `console` as officeParser's last-resort reporter when a call is given no options object at all
  (every public function passes one, so it is not reached through them).

### Encrypted templates are refused, not decrypted

officeParser decrypts a password-protected template (`password`, `onPassword`, `maxPasswordAttempts`)
with `src/crypto/*`, which imports Node's `crypto` (`createDecipheriv`, `createHash`, `pbkdf2Sync`).
**document-kit leaves encrypted templates out.** Porting to WebCrypto does not go cleanly:

- ECMA-376 *standard* encryption is AES-ECB, which WebCrypto does not offer; it would have to be built
  from AES-CBC one block at a time.
- *Agile* encryption decrypts 4,096-byte segments with AES-CBC and no padding; WebCrypto's AES-CBC always
  checks PKCS#7 padding, so each segment needs a forged padding block appended.
- Its key is SHA-512 iterated 100,000 times, one `await crypto.subtle.digest` each.
- OpenDocument's padding (W3C/ISO 10126) is not PKCS#7 either, and older ODF files use Blowfish, which
  WebCrypto does not have at all.

Each of those is a workaround with its own risk, for a feature neither application uses. Instead:

- A password-protected Word file (a CFB container with `EncryptionInfo` and `EncryptedPackage`, found
  with officeParser's own `cfb.ts`) and an OpenDocument file whose manifest holds `encryption-data` reject
  with a new code, **`TEMPLATE_ENCRYPTED`**, whose message says the kit does not decrypt templates and to
  save the template without a password.
- `password`, `onPassword` and `maxPasswordAttempts` are removed from `TemplateConfig` and from the
  options a call recognises, so one given is reported as `UNRECOGNIZED_CONFIG_OPTION`, then the template
  is refused as encrypted.
- `odtEntries(entries, decrypted)` keeps its second parameter (always `false` here), so the file stays
  as officeParser has it.
- A CFB container that is not encrypted OOXML (a legacy `.doc`) is refused with
  `TEMPLATE_UNSUPPORTED_FORMAT`, as in officeParser.

### OpenDocument templates are kept

`.odt` and `.ott` templates still work: carrying `odtTemplate.ts` cost nothing beyond the file, and the
ported ODT tests pass.

### API and errors

- **The class is functions.** `OfficeTemplate.render`, `.tags` and `.check` are `render`, `tags` and
  `check` in `officeTemplate.ts`, bodies unchanged, and `index.ts` exposes them with plain names:
  `fillTemplate` (one data object), `fillTemplates` (a list), `templateTags`, `checkTemplate` (one data
  object) and `checkTemplates` (a list). The single and batch forms refuse data of the other shape with
  `IMPROPER_ARGUMENTS`, where `render`'s overloads chose by the data. `renderTemplate` (the alias) is not
  carried.
- **Option names are officeParser's**, so its `TemplateConfig` documentation applies, less the three
  password options.
- **Errors are `TemplateError`s.** officeParser throws a plain `Error` carrying `officeIssue`; here it is
  a `TemplateError` (a subclass of `Error`) carrying the same `officeIssue`, with `code` read from it, and
  `isTemplateError(e, code?)`. The message header is `[document-kit template]: ` where officeParser's is
  `[OfficeParser]: `; the messages after it are officeParser's, but for `INVALID_INPUT`, which names the
  input kinds the kit takes. `getWrappedError` and `getWarningMessage` are not carried.
- **The codes keep officeParser's names and values** (`OfficeErrorType`, `OfficeWarningType`), cut to the
  ones the engine raises, plus `TEMPLATE_ENCRYPTED`. `OfficeIssue` loses `node` and `page`, which only the
  parser sets.
- The engine's internal `OfficeParserConfig` parameter type is `IssueHandler` (`{ onWarning? }`), which is
  all the engine read of it.
- `DecompressionLimits` keeps all six limits so an officeParser configuration carries over unreported;
  `maxTableCells` and `maxRawContentLength` are documented as unused by templates.

### Sentence filters

- **Four filters of document-kit's own** (`sentenceFilters.ts`, new; `templateEngine.ts` `isFilter` and the
  `default` branch of `compileFilters`, marked `document-kit:`): `capitalize`, `plural`, `prose` and
  `ranges`, shared with JSON recipes (`docs/template.md`, "Sentence filters"). A caller's `formatters` entry of
  the same name is used instead, so a template that already had one of these names as its own filter fills as
  before; only `format`, `upper` and `lower` stay reserved names (`BUILT_IN_FILTER_NAMES`, unchanged). A tag
  using one of the four is now a filter where it was text that looks like a tag (`TEMPLATE_NOT_A_TAG`) when no
  formatter of that name was given. Pinned by `tests/template/sentence-filters.test.ts`; officeParser's own
  tests are unchanged and pass.

### Type fixes

- `TemplateSyntaxDetails.detail` is `string | undefined` where it was `string`, because the engine sets it
  to `undefined` and the kit's `exactOptionalPropertyTypes` says so. No runtime change.
- `cfb.ts` defines `DecryptionError` itself, where it imported it from `wrongPassword.ts`.
- `TEMPLATE_OPTIONS` (the recognised option names) and its users read the kit's types.

### Type checking

The kit's `tsconfig.json` sets `noUncheckedIndexedAccess`, which officeParser's does not. Under it, the
engine has **217 errors on 169 lines** (146 in `templateEngine.ts`, 33 in `templateFormat.ts`, 26 in
`templateTags.ts`, 10 in `odtTemplate.ts`, 2 in `cfb.ts`), every one an array or record read the
compiler cannot prove is in range. Making them pass would mean a non-null assertion at each, so they are
left as they are. Every other option holds: `src/template/tsconfig.json` is the kit's with only
`noUncheckedIndexedAccess` off, and `npx tsc -p src/template/tsconfig.json --noEmit` passes.
`src/template/tsconfig.build.json` writes the folder's declarations to `dist/types/template` under it.

The root check (`npx tsc -p tsconfig.json`) still includes `src/template` and reports those 217 errors.
Proposed root change, not made (the root configuration is not this folder's to edit):

- `tsconfig.json` and `tsconfig.build.json`: add `"src/template"` to `exclude`. Nothing outside the folder
  imports from it.
- `package.json`: `typecheck` also runs `tsc -p src/template/tsconfig.json --noEmit`, and `build` also runs
  `tsc -p src/template/tsconfig.build.json`.

### Pinned timestamps

officeParser already pins every zip entry's time with `new Date(2001, 0, 1, 0, 0, 0)`, built from local
fields. fflate reads a Date's local fields (`getFullYear()`...), so that is 2001-01-01 00:00 in the zip in
every time zone. Unchanged. `tests/template/timezone.test.ts` fills one template under
America/Vancouver, Pacific/Pago_Pago (UTC−11), UTC, Asia/Kolkata and Pacific/Kiritimati (UTC+14) in one
process, and checks the bytes are identical and every central-directory entry reads 2001-01-01 00:00. For
contrast, `new Date(Date.UTC(1980, 0, 1))` handed to fflate under `TZ=America/Vancouver` throws
`date not in range 1980-2099`; the engine never builds one.

## Tests

`tests/template/upstream` ports officeParser's template tests:

| officeParser | Here | |
|---|---|---|
| `test/testExhaustive.ts`: `testTemplate` to `testTemplateBlocks` (lines 4037–6257), `testTemplateMistypedOperators`, with `docxParts` and `ooxmlHelpers` | `exhaustive.test.ts` | 18 vitest cases, one per function, the function bodies as they are |
| `test/helpers/docxTemplateCases.ts`, `test/files/template/docx-template-digests.json` | `docxTemplateCases.ts`, `tests/template/files/template/` | Each case's output must hash as officeParser 8.6.0's did |
| `test/security/testSanitization.ts`: the template cases of `parserHardeningTests`, `odtTemplateHardeningTests`, `templateTextBoundTests`, with `check`, `linear` and the memory readings | `hardening.test.ts` | 3 vitest cases; linear time and heap bounds |
| `test/generator/testOfficeGenerator.ts`: `runTemplatePropertyTests`; `test/helpers/lookalikeReference.ts`, `odtSpaceRuns.ts` | `properties.test.ts`, `lookalikeReference.ts`, `odtSpaceRuns.ts` | Random-input property checks |
| `test/parser/testOfficeParser.ts`: the hostile `<w:t>` run and the split-run fill | `properties.test.ts` | |
| `test/testExhaustive.ts`, config validation: a template's `maxPasswordAttempts` and `maxRepeatedContent` | `properties.test.ts` | |

Fixtures copied, all small and made by officeParser's own scripts: `invoice.docx`, `invoice.odt`,
`invoice.ott`, `shapes.odt`, `near-tags.odt`, `docx-template-digests.json`, and the synthetic encrypted
`agile.docx`, `standard.docx`, `encrypted.odt` (password `test123`, from
`scripts/generate-crypto-fixtures.mjs`). Not copied: the real LibreOffice-encrypted files (96 KB and
more), `conditions.docx` (no test reads it), and the parser fixtures `test.ods`/`test.odp`.

Changes to the tests, each marked `document-kit:` in the code:

- **`harness.ts`** stands in for what the tests take from the rest of officeParser. `OfficeTemplate`
  reads a path in the test, then calls the kit. `OfficeParser.parseOffice` is not officeParser's parser:
  it reads a filled `.docx` or `.odt` back as body paragraphs, tables (rows of tab-separated cells) and
  headers, with each Word run as a text node marked bold from its own properties. The tests use it for
  `includes` checks and two bold checks. `parseXmlString` is fast-xml-parser's validator. `withinTime` and
  `scaledMs` are officeParser's, with its calibration (which times its parser) replaced by a fixed factor,
  `DOCUMENT_KIT_TEST_SPEED_FACTOR` or 3.
- `testTemplate`'s time-zone case ran a child process under `tsx`; it is replaced by
  `tests/template/timezone.test.ts`.
- `testOdtTemplate`'s encrypted case expects `TEMPLATE_ENCRYPTED` (with and without the password) where it
  decrypted `real-libreoffice.odt`; it reads the synthetic `encrypted.odt`.
- `testOdtTemplate` builds the smallest `.ods` and `.odp` packages where it read the parser's fixtures.
- `testTemplateBlocks` no longer parses a Markdown file with bad limits (no parser), so it counts one
  `INVALID_CONFIG_VALUE` warning where it counted two.
- The config check of `maxPasswordAttempts` now expects it reported as unrecognised, and the template
  refused as encrypted.
- The property checks run officeParser's fast-mode sizes (20,000 texts, 1,000 paragraphs) unless
  `DOCUMENT_KIT_TEMPLATE_FULL=1`.
- `hardening.test.ts` reaches V8 through `createRequire`, and each vitest case fails when any `check` in
  it failed, naming them.

document-kit's own tests are `tests/template/fill.test.ts` (templates built with the `docx` library) and
`timezone.test.ts`.

## Bringing in a later officeParser

1. Diff officeParser's `src/template/*`, `src/OfficeTemplate.ts` and the pieces listed above against
   8.9.0, not against this folder.
2. Apply the upstream change here by hand, keeping the changes above (bytes only, no `Buffer`, no
   decryption, `IssueHandler`, the folder's own type-check).
3. Port the matching upstream tests into `tests/template/upstream` and run `npx vitest run tests/template`
   in a time zone west of Greenwich as well as at UTC.
4. Update the version and this file.
