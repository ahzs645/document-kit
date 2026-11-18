/**
 * Code components: a component whose body is a function the application registers, for a part that computes or
 * branches beyond what tags and control nodes can say. A recipe or a JSON component `use`s one exactly as it uses a
 * JSON component (same `with`, `variant`, `if`, `for`, `pageBreakBefore`, part boundary, to-complete, provenance and
 * hand edits). A compiled recipe keeps a marked reference to it (`{ "code": id, "params": … }`) and lists it in
 * `code`; rendering without it fails, naming it.
 *
 * `build` returns block templates: the same JSON a component's `blocks` holds, so its strings may carry `{{ tags }}`,
 * its lists control nodes, and a `use` of a JSON component it declared in `uses` (or of another code component).
 */
import type { Block } from '../model'
import { readComponent } from './read'
import type { Json, JsonObject, ParamSpec } from './types'

/** What a code component's `build` is given. */
export interface CodeBuildInput {
  /** Its parameters' values: the caller's `with`, the variant's, then the defaults; tags already read. */
  params: Readonly<Record<string, unknown>>
  /** The variant the caller chose, if any. */
  variant: string | null
  /** The whole data, as given to the render. */
  data: Readonly<Record<string, unknown>>
  /** The recipe's options. */
  options: Readonly<Record<string, unknown>>
  /** A data path's value, `[]`-free (`tables.0.rows`): wrapped values unwrapped, undefined when missing. */
  value(path: string): unknown
  /** The source the data gives a path (the `sources` map or a wrapped value's `$source`). */
  source(path: string): string | undefined
}

/** A block template, or a block of the model (pictures as bytes are fine: they are passed through). */
export type CodeBlock = Json | Block

export interface CodeComponent {
  id: string
  title: string
  description?: string
  /** Shown in the compiled recipe beside the id, so a stored compilation says what it was built against. */
  version?: string
  params?: Record<string, ParamSpec>
  variants?: Record<string, { title: string; with?: JsonObject }>
  /** Data paths it reads (`viewpoints[].visibility`), added to the compiled recipe's `needs`. */
  needs?: readonly string[]
  /** JSON components its blocks may `use`: the compile carries them in the compiled recipe (`embedded`). */
  uses?: readonly string[]
  /**
   * Extra details of the application's own, as a JSON component's `meta`: any JSON object, carried untouched to the
   * compiled part's mark (`componentMeta`) and the rendered part.
   */
  meta?: JsonObject
  build(input: CodeBuildInput): readonly CodeBlock[]
}

export interface CodeComponents {
  get(id: string): CodeComponent | undefined
  has(id: string): boolean
  readonly ids: readonly string[]
  /** A new set with one more component (or one replaced). */
  with(component: CodeComponent): CodeComponents
}

/** One code component, checked. */
export function defineCodeComponent(component: CodeComponent): CodeComponent {
  if (!/^[A-Za-z0-9][A-Za-z0-9_.-]{0,99}$/.test(component.id))
    throw new Error(`document-kit/recipe: a code component's id is not an id: ${JSON.stringify(component.id)}`)
  if (typeof component.build !== 'function')
    throw new Error(`document-kit/recipe: the code component "${component.id}" has no build function`)
  if (component.meta !== undefined) {
    // The same check a JSON component's meta takes: a JSON object, read as JSON and nothing more.
    try {
      readComponent({ kind: 'component', id: component.id, title: component.title, blocks: [], meta: component.meta })
    } catch (error) {
      throw new Error(`document-kit/recipe: the code component "${component.id}" has meta that is not a JSON object (${error instanceof Error ? error.message : String(error)})`)
    }
  }
  return component
}

/** A set of code components, passed to `compileRecipe` and `renderRecipe`. */
export function defineCodeComponents(components: readonly CodeComponent[] = []): CodeComponents {
  const map = new Map<string, CodeComponent>()
  for (const component of components) {
    defineCodeComponent(component)
    if (map.has(component.id)) throw new Error(`document-kit/recipe: two code components have the id "${component.id}"`)
    map.set(component.id, component)
  }
  return makeSet(map)
}

function makeSet(map: ReadonlyMap<string, CodeComponent>): CodeComponents {
  return {
    get: (id) => map.get(id),
    has: (id) => map.has(id),
    ids: [...map.keys()],
    with: (component) => {
      defineCodeComponent(component)
      const next = new Map(map)
      next.set(component.id, component)
      return makeSet(next)
    },
  }
}
