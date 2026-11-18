# document-kit agent instructions

document-kit is the shared document framework for PGMaps (forestry reports) and nalsreport (laboratory
reports). Both take it as a Git submodule at `vendor/document-kit`. It holds what the two have in common
and nothing that belongs to one of them.

## Rules that hold everywhere

- Everything under `src/` runs in a browser. No `node:` imports, no `fs`, no `process`, no server.
  Tests and scripts may use Node.
- No network access from `src/`. Fonts, templates, images and documents arrive as bytes from the caller.
  A missing font is reported as a loss, never fetched.
- No DOM access in `src/` unless a file's header says it is browser-only and why.
- No application content: no laboratory, university, forestry, licensee or client names, no real reports,
  no real people. Tests use made-up data. The one exception is the engines' own public test documents
  (`tests/ream/fixtures/real`, each under its stated licence in its `NOTICE.md`).
- Nothing either application can do today is removed by moving it here. The kit adds a common base; an
  application keeps its own writers until the kit's output is proven the same.

## Layout

| Folder | What it is | Public entry |
|---|---|---|
| `src/model` | The block model: a document as data before layout | `document-kit` |
| `src/docx` | Writes the block model as a Word file (the `docx` library) | `document-kit/docx` |
| `src/layout` | Word file to PDF, page pictures (SVG) and HTML, with fonts passed in | `document-kit/layout` |
| `src/template` | Fills a Word template's tags from data | `document-kit/template` |
| `src/recipe` | Reports from JSON components and recipes: compile, render, layering, own recipes, code components | `document-kit/recipe` |
| `src/ream` | The layout engine itself, brought in from reamkit | `document-kit/ream` |
| `fonts` | Open fonts that stand in for the common Office faces | `document-kit/fonts/*` |

Tests mirror the folders under `tests/`. `tests/browser` runs the built package in a real browser with
every outside request blocked.

## Code brought in from other projects

- `src/ream` is reamkit (MIT, Alexandr Krassavin). `src/template` is officeParser's template engine (MIT,
  Harsh Ankur). They are ours to change. Keep `NOTICE-REAMKIT.md` and `NOTICE-OFFICEPARSER.md`, and record
  every change from upstream in `docs/ream-upstream.md` or `docs/template-upstream.md`, with the upstream
  version it was taken from.
- Never copy, port or consult Carbone's source. Its licence forbids building a document generator with it.
- Bring in other code only when its licence is MIT, BSD, Apache-2.0 or public domain, with its notice.

## Checks

`npm run check` runs the type check, the unit tests, the build and the browser test. Run it before a
commit. The vendored engines' own tests are part of the unit tests and must keep passing in any time zone.
