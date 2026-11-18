# Recipes: reports built from JSON

`document-kit/recipe` builds a report from JSON. A **component** is a piece of a report written as the block model
(`docs/model.md`) in JSON, with `{{ tags }}` where data goes. A **recipe** is JSON that names components, in sections.
A recipe is **compiled** from its components into one self-contained JSON object, then **rendered** against data into
a kit `Document`, which `writeDocx` writes and `document-kit/layout` lays out.

```ts
import { compileRecipe, renderRecipe, sharedLibrary } from 'document-kit/recipe'
import { writeDocx } from 'document-kit/docx'

const library = sharedLibrary({ components: myComponents, recipes: myRecipes }) // the kit's shared components + yours
const compiled = compileRecipe('lab-report', library)          // one JSON object: store it, compare it, ship it
const { document, parts, toComplete, edits } = renderRecipe(compiled, data, { sources, edits: savedEdits })
const bytes = await writeDocx(document)
```

Everything runs in a browser or a worker: no Node, no DOM, no network.

## JSON or code?

**JSON by default.** A JSON component is reusable (any recipe uses it), storable (it is a file), shareable (send the
file), comparable (the compiled recipe names each component with the hash of its content) and overridable by a person
without a release (`layerLibraries`). Tags and control nodes say most of what a report part does: print a value,
format it, test a condition, repeat per item, pick a wording.

Code is left three jobs:

1. **Computing the data** a recipe reads. Arithmetic, joins and judgements belong in the application, which hands the
   recipe plain values.
2. **Custom block kinds and filters**: a block the model does not have (a rating matrix) through the model's custom
   block registry (`defineCustomBlocks`), and a filter (`{{ name | initials }}`) through `filters`.
3. **Code components**, for a part that computes or branches beyond what tags and control nodes can say (below).

A code component is the escape hatch, not the default. To move one to JSON later: compute its values into the data (a
`summary` object the application fills), keep its layout in a JSON component that prints `{{ summary.… }}`, and
replace the code component's id in the recipes with the JSON one's (a `replace` change keeps the part's id and so its
hand edits).

## Components

```json
{
  "kind": "component",
  "id": "greeting",
  "title": "Greeting",
  "description": "Optional words for a person choosing it.",
  "params": {
    "name": { "required": true },
    "closing": { "default": "Regards" }
  },
  "variants": {
    "formal": { "title": "Formal", "blocks": [{ "type": "paragraph", "content": "Dear {{ params.name }}," }] },
    "warm": { "title": "Warm", "with": { "closing": "All the best" } }
  },
  "blocks": [
    { "type": "paragraph", "content": "Hello {{ params.name }}." },
    { "type": "paragraph", "content": "{{ params.closing }}" }
  ]
}
```

| Key | |
|---|---|
| `kind` | `"component"` |
| `version` | Optional; a version newer than this kit reads (1) is refused |
| `id` | Letters, digits, `-`, `_`, `.`; unique in its library, and never a code component's id |
| `title`, `description` | Words for a person; the title is a part's title unless the recipe gives one |
| `params` | Each parameter: `default` (any JSON, tags read in the caller's scope) or `required: true`, and `description` |
| `variants` | Alternatives: each a `title` and either `blocks` (another body) or `with` (other parameter values), or both |
| `blocks` | The body: block templates (below) |
| `meta` | Extra details of the application's own (below) |

A parameter is read as `{{ params.name }}`. The caller's `with` wins over a variant's `with`, which wins over the
default. A `with` naming a parameter the component does not declare is an error, as is a required one not given.

### Block templates

A block template is the model's JSON (`docs/model.md`: `heading`, `paragraph`, `list`, `fields`, `table`, `figure`,
`figureGrid`, `signature`, `choices`, `pinned`, `pageBreak`, `toc`, `spacer`, `custom`) in which any string may hold tags. Every key is
checked against the model at compile time: a misspelt key is an error naming its JSON path.

Two shorthands the model does not have:

- a table's or figure's `caption` may be written as text (`"caption": "Results"`), which is `{ "text": "Results" }`;
  written as `{ "text": "{{ … }}", "numbered": … }`, it is left out whole when its text has no value, and a lone tag
  may give it whole from the data (`{ text, numbered }`);
- an optional text key (a table's or a figure's `note`, a fields block's or a contents block's `title`, a figure's
  `placeholder`, a signature's `label`, a picture's `label`) whose lone tag has no value is left out.

A picture (`image` of a figure, a signature or an inline image) comes from the data as bytes (`Uint8Array`, or
`{ bytes }`): JSON cannot hold it. One the data lacks is `null`: a placeholder box, listed to complete.

### Tags

The tag language is the Word template engine's, the same reader and the same formatters (`docs/template.md`): one
syntax for Word templates and JSON components.

```text
{{ client.name }}                       a field; data inside data, list items by number: {{ samples.0.id }}
{{ total | format "#,##0.00" }}         number and date pictures, upper, lower, and the application's filters
{{ n | plural "sample" "samples" }}     the sentence filters: capitalize, plural, prose, ranges (docs/template.md)
{{ run.date | missing "placeholder" "run date" }}   this value's missing-value rule (below)
{{ unit | rules "formulas" }}           named text rules applied to this value (below)
{{ if approved }}yes{{ else }}no{{ endif }}             inside one string
{{ for t in tags }}{{ t }}{{ if not loop.last }}, {{ endif }}{{ endfor }}
```

Names a tag reads:

| Name | |
|---|---|
| `params.*` | The component's parameters |
| `options.*` | The recipe's options (constants; the render may lay others over them) |
| a loop's item, `loop.*` | Inside a `for`: `loop.index` (from 0), `loop.number` (from 1), `loop.first`, `loop.last`, `loop.length` |
| anything else | The data |

So `params`, `options` and `loop` are not data names. A component does not see its caller's loop items: what it needs
comes through its parameters.

**A string that is exactly one tag gives the raw value** (a list, an object, a number), so data can be handed to a
table's rows, a parameter, a custom block's data: `"rows": "{{ table.rows }}"`. Any text around the tag, even a space,
makes the string text. Every `{{` starts a tag; a tag that does not read is an error at compile time.

### Control nodes

Control nodes stand wherever a block, a table row, a cell, a field row, a list item, a picture of a grid or an inline
run stands: in any list. The items they give are spliced into the list.

```json
{ "if": "status == \"final\" and total > 0", "then": [ … ], "else": [ … ] }
{ "for": "row in table.rows", "do": [ … ], "empty": [ … ] }
{ "use": "data-table", "with": { "rows": "{{ table.rows }}" }, "variant": "compact" }
```

- `if` takes a condition as the Word engine reads one (`==`, `!=`, `<`, `<=`, `>`, `>=`, `and`, `or`, `not`,
  brackets, filters), written without braces. A value the data lacks is false.
- `for` takes `item in list`. A list the data lacks repeats nothing (and prints `empty`).
- `use` puts a component in place, with `with` and `variant`. A component used where rows stand gives rows, where
  runs stand gives runs; a code component stands only where a block does.

A control node in a custom block's `data` is read the same way, so a key named `if`, `for` or `use` there makes the
object a control node.

## Recipes

```json
{
  "kind": "recipe",
  "id": "lab-report",
  "name": "Laboratory report",
  "document": { "metadata": { "title": "{{ report.title }}", "created": "2026-03-12T09:00:00Z" }, "theme": { } },
  "options": { "showRemarks": true },
  "sections": [
    {
      "id": "report",
      "page": { "paper": "letter", "orientation": "portrait", "margins": { "top": 54 }, "verticalAlign": "top" },
      "header": [{ "use": "running-header", "with": { "left": "Made-up Testing Laboratory", "right": "{{ report.code }}" } }],
      "footer": [{ "use": "page-footer", "with": { "text": "Report {{ report.number }}" } }],
      "parts": [
        { "use": "report-title", "id": "title", "with": { "title": "{{ report.title }}" } },
        { "use": "sample-notes", "id": "notes", "variant": "brief" },
        { "use": "results-table", "id": "results", "for": "table in tables", "with": { "table": "{{ table }}" } },
        { "use": "signature", "id": "sign-off", "if": "signer", "pageBreakBefore": true, "with": { "name": "{{ signer.name }}" } }
      ]
    }
  ]
}
```

| Key | |
|---|---|
| `kind`, `version`, `id`, `name`, `description` | As for components |
| `extends` | Another recipe's id: this recipe is that one with `changes` |
| `document` | `metadata` and `theme` as the model takes them; strings may hold tags. `textRules` as the model takes them (`docs/model.md`, "Text rules"), carried as written: no tags are read in them. They are checked at compile, each problem named by its place (`recipe/lab:document.textRules[2].pattern`) |
| `missing` | The missing-value rule of every part that does not give its own: `blank` (default) or `placeholder` (below) |
| `meta` | Extra details of the application's own (below) |
| `options` | Constants the components read as `options.*` |
| `sections` | Each: `id` (for changes), `page`, `header`, `footer`, `firstPage: { header, footer }`, `pageNumbers`, `parts`. Headers and footers are block templates, `use` included. Left out, a header is the section before's (Word's rule); a section with another page shape should give its own, since a tab stop at the right edge is measured on the page it was written for |
| `changes` | For an extending recipe (below) |
| `own` | `{ "basedOn": id }` on a person's own recipe |

A **part** is `{ "use", "id"?, "title"?, "variant"?, "with"?, "if"?, "for"?, "pageBreakBefore"?, "missing"?, "pin"?,
"meta"? }`. `missing` is the part's missing-value rule; `pin` (`"bottom"`, or `{ "spaceBelow": 24 }`) pins the part's
blocks to the foot of the page they fall on, as one `pinned` block (`docs/model.md`, "Pinned blocks"); Its id (the
component's when not given) is how changes and hand edits name it, unique in the recipe. `for` prints the part once per
item, and the item is seen by `with` and `if`; `if` is tested on each pass.

### Extending a recipe

```json
{
  "kind": "recipe",
  "id": "lab-report-interim",
  "name": "Laboratory report, interim",
  "extends": "lab-report",
  "changes": [
    { "set": "title", "with": { "status": "Interim" } },
    { "add": { "use": "notice", "id": "interim-notice", "with": { "text": "{{ pending.count }} results to follow." } }, "after": "title" },
    { "set": "notes", "variant": "brief" }
  ]
}
```

The base is resolved first (its own base before it), then the extending recipe's `sections` (if it gives any) replace
its base's, its `document` is merged over the base's key by key, its `options` laid over the base's, and its `changes`
applied in order:

| Change | |
|---|---|
| `{ "add": part, "before" \| "after": id }` or `{ "add": part, "section": id \| index }` | A part, by another or at the end of a section (the last when neither) |
| `{ "remove": id }` | |
| `{ "replace": id, "by": part }` | The part keeps its id (and its hand edits) unless `by` names its own |
| `{ "move": id, "before" \| "after" \| "section": … }` | |
| `{ "set": id, "variant", "with", "if", "for", "title", "pageBreakBefore", "missing", "pin", "meta" }` | Each given key replaces the part's; `with` merges by parameter; `null` removes |
| `{ "setSection": id \| index, "page", "header", "footer", "firstPage", "pageNumbers" }` | `null` removes |
| `{ "addSection": section, "after": id \| index }` | |

A change naming a part or section the recipe does not have is an error.

## Libraries and a person's own changes

A library is `{ "components": [...], "recipes": [...] }`. `sharedLibrary(own)` is the kit's shared components plus the
application's.

A person's changes travel as **overrides**, themselves JSON:

```json
{
  "kind": "library-overrides",
  "version": 1,
  "components": { "notice": { "kind": "component", "id": "notice", … }, "old-part": null },
  "recipes": { "own-lab-1a2b": { "kind": "recipe", … } }
}
```

An item replaces the shipped one of its id (or is added); `null` removes the shipped one.
`layerLibraries(shipped, overrides, { code })` gives `{ library, status, problems }`: each id's status is `shipped`,
`modified`, `added` or `removed` (a change equal to the shipped item is `shipped`), and `problems` names a recipe left
using a component that is gone. An override cannot replace a **code** component: a JSON component with a code
component's id is named in `problems` and refused by the compile (`component-conflict`); give it another id and point
the recipe at it. `withOverride`, `withRemoval`, `withoutOverride`, `overridesFileText` and `readOverridesFile` make
and keep the file.

**Own recipes**: `copyRecipe(recipe, takenIds, { as: 'copy' | 'extension' })` makes one (an extension `extends` the
original, so later fixes to it still arrive), `editRecipe(recipe, { name, options, changes })` changes it (changes go
straight into its own sections, or into its `changes` when it extends another), and `recipeFileText` /
`readRecipeFile` share it, an import taking a fresh id and name when its own are in use.

## Code components

```ts
import { defineCodeComponents } from 'document-kit/recipe'

const code = defineCodeComponents([{
  id: 'review-summary',
  title: 'Summary',
  version: '1',
  params: { assessor: { required: true } },
  needs: ['viewpoints[].share'],            // added to the compiled recipe's needs
  uses: ['notice'],                         // JSON components its blocks may `use`
  build: ({ params, variant, data, options, value, source }) => [
    { type: 'paragraph', content: `${count(value('viewpoints'))} viewpoints see it.`, source: 'computed' },
    { type: 'fields', rows: [{ label: 'Assessed by', value: '{{ params.assessor }}' }] },
    { use: 'notice', with: { text: 'Above the limit.' } },
  ],
}])

const compiled = compileRecipe('review', library, { code })
renderRecipe(compiled, data, { code })
```

A recipe or a JSON component uses one exactly as a JSON component: `{ "use": "review-summary", "with": … }`, with
`variant` (declared in `variants`), `if`, `for` and `pageBreakBefore`. The compiled recipe keeps a marked reference
`{ "code": "review-summary", "params": { … } }` and lists it in `code`; the JSON components it `uses` travel in the
compiled recipe's `embedded`. Rendering without it fails with `missing-code`, naming it; `checkRecipe` reports it.

`build` returns block templates: the same JSON a component's `blocks` holds, so its strings may carry tags (read in the
code component's scope: its parameters, the data, the options), its lists control nodes, and a `use` of a JSON
component it declared in `uses` (another, `undeclared-use`) or of another code component. Pictures as bytes pass
through. Its blocks are checked as a component's are, and to-complete, provenance and hand edits apply to them as to any
part's. Code components nesting through what they return count against the render's `includeDepth`.

## The compiled recipe

`compileRecipe(recipeId | recipe, library, { code?, filters?, limits? })` returns one object with no reference left but
code components:

```json
{
  "kind": "compiled-recipe",
  "version": 1,
  "recipe": { "id": "letter", "name": "Letter" },
  "from": {
    "recipes": [{ "id": "letter", "hash": "sha256:40df…" }],
    "components": [{ "id": "fields", "hash": "sha256:683d…" }, { "id": "greeting", "hash": "sha256:cdcc…" }],
    "code": [],
    "hash": "sha256:81fd…"
  },
  "document": {},
  "options": {},
  "sections": [{
    "parts": [{
      "part": { "id": "greeting", "component": "greeting", "title": "Greeting" },
      "params": { "name": "{{ client.name }}" },
      "blocks": [{ "type": "paragraph", "content": "Hello {{ params.name }}." }]
    }]
  }],
  "needs": [{
    "path": "client.name", "uses": ["value"], "inLoop": false,
    "places": [
      { "part": "greeting", "at": "sections[0].parts[0].params.name", "use": "value", "inLoop": false },
      { "part": "greeting", "at": "sections[0].parts[0].blocks[0].content", "use": "value", "inLoop": false }
    ]
  }],
  "customKinds": [], "filters": [], "code": []
}
```

- A part: its mark (`part`), its `if`, `for`, `pageBreakBefore` and `variant`, its parameter bindings (`params`, read in
  the part's scope at render) and its `blocks` (every `use` inlined as `{ "component", "variant", "params", "blocks" }`)
  or its `code`.
- `needs`: every data path the recipe reads, with `[]` for each list item (`tables[].rows[].value`), how it is read
  (`value` printed, `condition` tested, `list` looped over), whether in a loop, and each place (the part and the JSON
  path in the compiled recipe). Paths read through a parameter are followed to the data they are bound to. A code
  component's declared reads are marked `declared`. An application checks its data against this, or builds a form from
  it.
- `customKinds`: custom block kinds the document holds. `filters`: filters other than `format`, `upper`, `lower`.
  `code`: code components, with the parts using them.
- `from`: the recipes and components it was compiled from, each with the SHA-256 of its JSON, code components with the
  version they declared, and one hash over all: two compilations from the same sources say so.

Compiling is deterministic: the same inputs give the same JSON text (keys in sorted order), whatever the key or library
order. `readCompiledRecipe` reads one back, refusing a newer version.

## Rendering

`renderRecipe(compiled, data, options)` returns `{ document, parts, toComplete, edits }`:

| Option | |
|---|---|
| `registry` | The model's custom block registry: the to-complete list walks what custom blocks become; given, every kind the recipe uses must be in it |
| `filters` | `{ name: (value, ...args) => value }` |
| `code` | Code components |
| `sources` | Provenance by data path (below) |
| `edits`, `editSource` | Hand edits (below); edited values carry `editSource` (`edited` by default) |
| `options` | Laid over the recipe's options |
| `numberSeparators`, `dataNumberSeparators`, `dateNames`, `utcOffsetMinutes`, `numericTextComparison` | As the Word engine's (`docs/template.md`) |
| `limits` | Below |

- `document`: a kit `Document`, ready for `writeDocx` (pass the same `registry`) and the layout.
- `parts`: per part instance, its `id`, `component`, `title`, `section`, where its blocks are in the section's body
  (`start`, `end`), how many top-level blocks it printed (0 when its `if` was false or it printed nothing; a pinned
  part prints one), how many values it left to complete, how many times it was printed (`repeats`), and the part's
  and its component's `meta` (`meta`, `componentMeta`).
- `toComplete`: the model's `toComplete` items (`docs/model.md`), each with its `partId` (null for a header or footer),
  in reading order, and a `value` kind for a value the data lacks at a place that prints no blank (a number, a plain
  text key).

### Missing values

Three rules. The first two both applications have, and are the default (`blank`); the third is chosen by
`"missing": "placeholder"` on a recipe, on a part, or for one value through the `missing` filter:

- **A value the data lacks in a field, a cell or text prints blank and is listed.** A field's value or a cell given by a
  lone tag becomes `null` (the model prints a blank rule or an empty cell); a tag inside text becomes a blank rule run.
  `null` and `""` count as no value.
- **A sentence whose slot is missing is not printed and is listed.** Use the model's sentence run, whose values are
  tags: `{ "type": "sentence", "label": "Visibility", "template": "It is {visibility} from {viewpoint}.",
  "values": { "visibility": "{{ vp.visibility }}", "viewpoint": "{{ vp.name }}" } }`. A sentence's `template` takes
  `{slots}`, never tags.
- **Placeholder: a value the data lacks prints its bracketed name, marked, and is listed.** A field's value, a cell or
  a tag inside text prints `[name]` in the theme's `placeholderMark` look (a yellow highlight by default), so an
  unfilled value cannot pass for prose: "Run on [metals run date]." The name is the field's label for a field's
  value, else the tag's path, or the one the value gives: `{{ run.date | missing "placeholder" "metals run date" }}`.
  A sentence prints with its missing slots so (`[visibility]`) rather than being left out (a sentence may say its own
  `"missing": "drop"` or `"placeholder"`). Places that print nothing for a missing value (an optional note or title, a
  number) are unchanged. Listed to complete as a `blank` or a `sentence`, like the others.

`{{ x | missing "blank" }}` asks for the first rule for one value inside a placeholder part. The compiled recipe keeps
the recipe's rule (`missing`) and each part's.

A condition reads a missing value as false; a missing list repeats nothing. `checkRecipe(compiled, data, options)`
reports, without making the document, every data path a printed tag or a loop reads that the data does not have (by its
data path with indexes, `rows.1.value`, in every pass of every loop and the branch each condition takes), and every
code component, filter or custom kind not given. A condition's missing value is not reported (it is false by design).

The to-complete list printed in the document is the shared `to-complete` component: the render fills it with every
other item.

### Provenance

A value says where it came from in the data, two ways:

- a `sources` map by data path: `{ "client": "office", "tables[].rows": "instrument", "viewpoints.*.visibility":
  "reviewer" }`, `[]` and `*` matching any item; the most specific key wins;
- a wrapped value in the data: `{ "$value": "Easy to see", "$source": "reviewer" }`, unwrapped wherever it is read.

A printed value keeps its source as the model's `source` on its run (`writeDocx(…, { sourceStyles })` marks it), a
sentence takes its slots' source when they agree, and a blank carries the source its value would have had (a field row
or cell `source`, which the to-complete item names). A source travels through parameters.

### Hand edits

`edits` maps a key to a replacement text. Each part fills its blocks, then its edits are applied: a field's value, a
signature's name, a body cell, a paragraph or a list item becomes the edit (`""` blanks it, printed blank and listed),
carrying `editSource`. The key is the part instance id and the place in the part:
`part|scope|kind|label`, where scope is the heading the value sits under (or `#id` of a block with an `id`), kind is
`field`, `name`, `cell`, `paragraph` or `item`, and label the field's label, or a cell's table caption, row (named by
its first cell, never its place) and column header; a repeated key gains `#2`, `#3`. Keys name no recipe, so two
recipes sharing a part share its edits.

A **tick-box question** (`choices` block) gives two more kinds: `choice`, whose edit is one of the question's
options (or `""` for none ticked) and is the option ticked, and `answer`, the free-text answer line. The key is the
same shape, labelled by the question: `part||choice|Is the objective met?`. Its target lists the `options`; an edit
that is not one of them is not applied and is reported in `orphans`. A question with several ticks is not editable by
tick. Blocks in a pinned block take edits like any others.

`edits.targets` lists every editable value with its key, its printed `original` and its `edited` text when edited.
`edits.orphans` lists edits for this recipe's parts whose place was not printed this time: kept, reported, never applied
elsewhere. `edits.otherParts` lists edits for parts this recipe does not have.

### Extra details: meta

A recipe, a component and a part may carry `meta`: any JSON object of the application's own (its sources, tags,
counts, its own ids). It is read as JSON and nothing more — no tags are read in it, and no key in it is checked —
and carried through compile untouched: the recipe's to `compiled.recipe.meta`, a part's to its mark
(`part.meta`), its component's to `part.componentMeta`, and both to the rendered part. Unknown keys anywhere else
are still errors.

### Text rules in a recipe

The recipe's `document.textRules` go to the document, where the Word writer applies them (`docs/model.md`, "Text
rules"). A block template says which touch it with the model's `textRules` key (a list of rule ids, or `false`). One
value can take named rules wherever it is printed with the `rules` filter: `{{ unit | rules "formulas" }}` applies
those rules (whatever their `auto` and `in`) to the value and prints the runs they make, inside text too. A `rules`
filter naming an id the document does not have is an error at compile (`bad-tag`).

## Limits

| Limit | Default | Where |
|---|---|---|
| `includeDepth` | 16 components nested by `use` (code components through what they return included) | compile (`COMPILE_LIMITS`), render (`RENDER_LIMITS`) |
| `extendsDepth` | 16 recipes in an `extends` chain | compile |
| `nodes` | 200,000 JSON nodes walked in the compiled templates | compile |
| `loopPasses` | 100,000 passes of every loop together | render |
| `blocks` | 50,000 blocks made, at every depth | render |
| condition nesting | 32 levels of brackets or `not` | the Word engine's reader |

Each is lowered or raised with `limits`. Past one, the compile or render stops with `limit`.

## Errors

Every refusal is a `RecipeError` with a `code` and a `path` (`component/notice:blocks[0].content`,
`recipe/lab-report:sections[0].parts[2].with.title`, or a compiled recipe's JSON path at render):

| Code | |
|---|---|
| `invalid` | A file that does not read; `problems` lists each with its path. Unknown keys are problems |
| `version` | A newer format version |
| `unknown-recipe`, `unknown-component`, `unknown-variant`, `unknown-part` | A name that does not resolve |
| `missing-param`, `unknown-param` | |
| `cycle` | A component that uses itself, or a recipe that extends itself, through any chain |
| `bad-tag` | A tag, condition or loop that does not read; a tag in a sentence's template |
| `unknown-key`, `bad-node` | A key the model does not have there; a node that does not stand there |
| `duplicate-part`, `component-conflict`, `limit` | |
| `unknown-filter`, `missing-code`, `missing-custom-kind` | What the application did not give |
| `invalid` | Also a text rule that does not read, each problem by its place |
| `not-text`, `bad-value` | A list or an object printed as text; a value a filter or place cannot take |
| `code-failed`, `undeclared-use` | A code component that threw or returned no list; one using a component it did not declare |

## The shared components

`sharedComponents()` (JSON in `src/recipe/components/`):

| Id | Parameters |
|---|---|
| `cover-page` | `title`*, `subtitle`, `lines` ([{ label, value }]), `note` |
| `contents` | `title` ("Contents"), `levels` (2) |
| `fields` | `rows`* ([{ label, value, what?, sameLine? }]), `heading`, `level`, `layout` (lines / table), `labelWidth`, `size`, `rowSpacing`, `sameLineGap` |
| `text-section` | `heading`, `level`, `paragraphs` (text or rich text each) |
| `notice` | `text`*, `title`, `fill`, `rule` |
| `data-table` | `columns`* (labels), `rows` (each a list of cells, `{ group }`, `{ note }` or `{ header: [labels] }`), `caption`, `numbered` (true; false: the caption alone, taking no number; `"silent"`), `note`, `widths`, `align`, `size`, `padding` ({ top, bottom, left, right }), `line`, `repeatHeader`, `autoFit`, `empty` |
| `figure` | `image`*, `caption`, `numbered` (true), `note`, `placeholder`, `width`, `height` (300) |
| `figure-page` | `title`*, `image`*, `caption`, `facts`, `height` (360), `level`, `placeholder` |
| `signature` | `name`*, `label` ("Signed"), `labelPosition`, `labelGap`, `image`, `imageWidth` (140), `imageHeight` (55), `inkGap` (room to sign by hand), `title`, `titleLayout` (`row`: a "Title:" line, the default; `sameLine`: "Name, Title"; `nextLine`), `titleSeparator`, `date`, `dateLine` (true), `lines` ([{ label, value }] in place of the title and date lines), `align`, `ruleWidth`, `ruleThickness`, `ruleColor`, `ruleGap`, `spaceAbove`, `spaceBelow`, `keepTogether` (true) |
| `to-complete` | `title`, `level`, `none` |
| `page-footer` | `text`, `page`, `of`: "text … Page X of Y" |
| `running-header` | `left`, `right` |

(* required.)

## Worked examples

`tests/recipe/fixtures/` holds two made-up report families built from one library: the kit's shared components and a
few of their own.

- **Laboratory** (`recipes/laboratory.json`): a running header and a "Page X of Y" footer, a title, the addressee as
  fields (one blank; the phone on the e-mail's line), notes with two wordings (`components/sample-notes.json`,
  variants) under the placeholder rule (a missing slot printed `[basis]`), one results table component used once per
  table in the data (`for`), footnote lines whose markers a rule named by their blocks raises
  (`components/footnotes.json`), a signature with its options, and a contact block pinned to the foot of the last
  page. Its document states text rules (links, a term in italics, raised cell markers, formula digits) and it carries
  `meta` on the recipe, a part and a component.
  `recipes/laboratory-interim.json` extends it: its status in the title, an interim notice after it, the brief notes.
- **Review** (`recipes/review.json`): a centred cover section, contents, project fields with blanks, a summary computed
  by a code component (`code.ts`, which uses the shared `notice`), one heading and two sentences per viewpoint (one
  left out), a boxed tick-box question with an answer line (`components/tick-question.json`) that a hand edit can
  change, a rating matrix drawn by a custom block kind (the application's registry, `library.ts`), a landscape
  section with a wide table, an 11 × 17 figure page, and the to-complete list.

`data/` holds their made-up data; `tests/recipe/end-to-end.test.ts` runs them to PDF and reads the pages.

## Type checking

`src/recipe` reads tags with `src/template`'s own reader and formatters, which type-check under that folder's tsconfig
(the kit's options with `noUncheckedIndexedAccess` off). A program that imports them checks them under its own options,
so `src/recipe` has its own tsconfig with the same one relaxed option (`src/recipe/tsconfig.json`), run by `typecheck`,
and its own declaration build (`src/recipe/tsconfig.build.json`), run by `build`; the root configs leave the folder out,
as they leave out `src/template`. `tests/recipe` is checked with `tsconfig.integration.json`, for the same reason.

## What a recipe cannot say yet

- A variant chosen by the data at render: a variant is chosen at compile time (compile again with another `variant`, or
  `set` it in an own recipe).
- A key read by a computed name (`row[column.key]`): rows are lists of cells, or the data gives objects in the model's
  shape.
- Literal `{{` in text, and an object in a custom block's `data` with a key `if`, `for` or `use` (it is read as a
  control node): give such values through the data.
- Section-level changes from a component (a component cannot open a section or set its page; the recipe does).
- Everything in `docs/model.md`'s list of what the model cannot say yet.
