/**
 * The JSON shapes of the recipe system: components, recipes, libraries, overrides and the compiled recipe. Every shape
 * here is plain JSON (what `JSON.parse` gives), so it can be stored, shared, compared and edited as data. The guide is
 * docs/recipe.md.
 */

/** A JSON value. */
export type Json = null | boolean | number | string | Json[] | { [key: string]: Json }
export type JsonObject = { [key: string]: Json }

/** The format version this kit writes and reads. A file of a newer version is refused. */
export const RECIPE_FORMAT_VERSION = 1

/* ------------------------------------------------------------------ components */

/** A parameter a component takes. */
export interface ParamSpec {
  /** Used when the caller gives no value. May hold tags, read in the caller's scope. */
  default?: Json
  /** The caller must give a value (a default never satisfies it). */
  required?: boolean
  description?: string
}

/** An alternative body (`blocks`) or alternative parameter values (`with`) for a component. */
export interface ComponentVariant {
  title: string
  description?: string
  blocks?: Json[]
  with?: JsonObject
}

/** A reusable piece of a document, written as the block model with `{{ tags }}` and control nodes. */
export interface Component {
  kind: 'component'
  version?: number
  id: string
  title: string
  description?: string
  params?: Record<string, ParamSpec>
  variants?: Record<string, ComponentVariant>
  blocks: Json[]
  /** Extra details of the application's own (sources, tags, an id): any JSON, carried through compile untouched. */
  meta?: JsonObject
}

/**
 * What a value nobody gave prints: `blank` (the default) — a blank rule, or an empty cell, for a value, and a sentence
 * left out for a missing slot; `placeholder` — the value's bracketed name, marked (`[run date]`), and a sentence
 * printed with its missing slots so. Listed to complete either way.
 */
export type MissingRule = 'blank' | 'placeholder'

/** A part pinned to the foot of the page: `"bottom"`, or the room kept under it. */
export type PartPin = 'bottom' | { at?: 'bottom'; spaceBelow?: number }

/* ------------------------------------------------------------------ recipes */

/** One part of a recipe: a component, used once (or once per item of `for`). */
export interface PartRef {
  use: string
  /** The part's id in this recipe (hand edits and `changes` name it); the component's id when absent. */
  id?: string
  /** The part's title in `parts` and the edit list; the component's title when absent. */
  title?: string
  variant?: string
  with?: JsonObject
  /** What a value the data lacks prints in this part; the recipe's rule when absent. */
  missing?: MissingRule
  /** Pin the part's blocks to the foot of the page they fall on. */
  pin?: PartPin
  /** Extra details of the application's own: any JSON, carried to the compiled part and the rendered part. */
  meta?: JsonObject
  /** A condition (no braces): the part prints only when it holds. */
  if?: string
  /** `item in list`: the part prints once per item; `item` is seen by `with` and `if`. */
  for?: string
  pageBreakBefore?: boolean
}

export interface RecipeSection {
  /** A name `changes` can use for the section; its index otherwise. */
  id?: string
  page?: Json
  header?: Json[]
  footer?: Json[]
  firstPage?: { header?: Json[]; footer?: Json[] }
  pageNumbers?: Json
  parts: PartRef[]
}

/** Where a part is put: before or after another part, or at the end of a section. */
export interface PartAnchor {
  before?: string
  after?: string
  /** A section's id or index; the last section when absent. Used when neither `before` nor `after` is given. */
  section?: string | number
}

/** One change a derived recipe makes to its base. Applied in order. */
export type RecipeChange =
  | ({ add: PartRef } & PartAnchor)
  | { remove: string }
  | { replace: string; by: PartRef }
  | ({ move: string } & PartAnchor)
  | {
      set: string
      variant?: string | null
      with?: Record<string, Json | null>
      if?: string | null
      for?: string | null
      title?: string | null
      pageBreakBefore?: boolean | null
      missing?: MissingRule | null
      pin?: PartPin | null
      meta?: JsonObject | null
    }
  | {
      setSection: string | number
      page?: Json
      header?: Json[] | null
      footer?: Json[] | null
      firstPage?: { header?: Json[]; footer?: Json[] } | null
      pageNumbers?: Json
    }
  | { addSection: RecipeSection; after?: string | number }

export interface RecipeDocument {
  metadata?: JsonObject
  theme?: JsonObject
  /** The document's text rules (the model's `TextRule`s), carried as written: no tags are read in them. */
  textRules?: Json[]
}

export interface Recipe {
  kind: 'recipe'
  version?: number
  id: string
  name: string
  description?: string
  /** Another recipe's id: this recipe is that one with `changes` applied. */
  extends?: string
  document?: RecipeDocument
  /** Constant settings the components read as `options.*`. A derived recipe's options are laid over its base's. */
  options?: JsonObject
  /** The sections. A derived recipe that gives them replaces its base's before `changes` apply. */
  sections?: RecipeSection[]
  changes?: RecipeChange[]
  /** An own recipe: the recipe it was copied or extended from. */
  own?: { basedOn: string | null }
  /** What a value the data lacks prints, for every part that does not say (`blank` by default). */
  missing?: MissingRule
  /** Extra details of the application's own: any JSON, carried to the compiled recipe untouched. */
  meta?: JsonObject
}

/* ------------------------------------------------------------------ libraries and overrides */

export interface Library {
  kind?: 'library'
  version?: number
  components: Component[]
  recipes: Recipe[]
}

/** A user's changes over a shipped library: an item as it should be, or null to remove the shipped one. */
export interface LibraryOverrides {
  kind: 'library-overrides'
  version: number
  components?: Record<string, Component | null>
  recipes?: Record<string, Recipe | null>
}

export type OverrideStatus = 'shipped' | 'modified' | 'added' | 'removed'

export interface LayeredLibrary {
  library: Library
  status: { components: Record<string, OverrideStatus>; recipes: Record<string, OverrideStatus> }
  /** References that no longer resolve (a recipe using a removed component), and ids that clash with code components. */
  problems: string[]
}

/* ------------------------------------------------------------------ the compiled recipe */

/** How a data path is read. */
export type NeedUse = 'value' | 'condition' | 'list'

/** One data path a recipe reads. `[]` in a path stands for every item of a list: `tables[].rows[].value`. */
export interface RecipeNeed {
  path: string
  /** Every way it is read: printed, tested, looped over. */
  uses: NeedUse[]
  /** Read inside a loop at some place. */
  inLoop: boolean
  places: NeedPlace[]
}

export interface NeedPlace {
  /** The part instance id, or null for a section's header or footer. */
  part: string | null
  /** The JSON path in the compiled recipe where it is read. */
  at: string
  use: NeedUse
  inLoop: boolean
  /** Declared by a code component rather than read by a tag. */
  declared?: true
}

/** A code component the compiled recipe needs at render. */
export interface CodeNeed {
  id: string
  title: string
  version?: string
  /** The part instances that use it (null: a header or footer). */
  parts: Array<string | null>
}

export interface CompiledPartMark {
  id: string
  component: string
  title: string
  /** The part's own `meta`, as written. */
  meta?: JsonObject
  /** Its component's `meta`, as written. */
  componentMeta?: JsonObject
}

/** A part as compiled: its mark, its condition and loop, and its body (a component instance or a code node). */
export interface CompiledPart {
  part: CompiledPartMark
  if?: string
  for?: string
  pageBreakBefore?: boolean
  /** A JSON component's variant, when one was chosen. */
  variant?: string
  missing?: MissingRule
  pin?: PartPin
  /** Parameter bindings, read in the part's scope at render (tags allowed). */
  params: JsonObject
  /** A JSON component's blocks, every `use` inlined. */
  blocks?: Json[]
  /** A code component's id: built by the application's code at render. */
  code?: string
}

export interface CompiledSection {
  id?: string
  page?: Json
  header?: Json[]
  footer?: Json[]
  firstPage?: { header?: Json[]; footer?: Json[] }
  pageNumbers?: Json
  parts: CompiledPart[]
}

export interface CompiledFrom {
  /** The recipe and each recipe it extends, nearest first, each with the hash of its JSON. */
  recipes: Array<{ id: string; hash: string }>
  /** Every JSON component inlined, with the hash of its JSON. */
  components: Array<{ id: string; hash: string }>
  /** Every code component named, with the version it declared. */
  code: Array<{ id: string; version?: string }>
  /** One hash over all of the above. */
  hash: string
}

export interface CompiledRecipe {
  kind: 'compiled-recipe'
  version: number
  recipe: { id: string; name: string; description?: string; meta?: JsonObject }
  from: CompiledFrom
  /** The recipe's missing-value rule, when it is not `blank`. */
  missing?: MissingRule
  document: RecipeDocument
  options: JsonObject
  sections: CompiledSection[]
  needs: RecipeNeed[]
  /** Custom block kinds the document holds (an application's registry expands them). */
  customKinds: string[]
  /** Filters other than the built-in `format`, `upper` and `lower`. */
  filters: string[]
  /** Code components the render needs. */
  code: CodeNeed[]
  /** JSON components a code component declared it may `use`, as written. */
  embedded?: Record<string, Component>
}
