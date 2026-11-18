/**
 * Libraries as a person changes them, and their own recipes. `layerLibraries` lays a set of overrides (each item as it
 * should be, or null to remove a shipped one) over a shipped library and says, per item, whether it is shipped,
 * modified, added or removed. The overrides are themselves JSON, saved and loaded as a file. Own recipes are copied or
 * extended from another, changed, and shared as a file. Nothing here touches the shipped library. Pure.
 */
import type { CodeComponents } from './code'
import { applyChanges, partId } from './compile'
import { RecipeError } from './errors'
import { canonical, cloneJson, sameJson } from './json'
import { readLibrary, readOverrides, readRecipe } from './read'
import {
  RECIPE_FORMAT_VERSION,
  type Component,
  type Json,
  type JsonObject,
  type LayeredLibrary,
  type Library,
  type LibraryOverrides,
  type OverrideStatus,
  type Recipe,
  type RecipeChange,
} from './types'

export const emptyOverrides = (): LibraryOverrides => ({ kind: 'library-overrides', version: RECIPE_FORMAT_VERSION, components: {}, recipes: {} })

type Kind = 'component' | 'recipe'
const keyOf = (kind: Kind) => (kind === 'component' ? 'components' : 'recipes')

/** The overrides with one item set (added, or a shipped one changed). */
export function withOverride(overrides: LibraryOverrides, kind: Kind, item: Component | Recipe): LibraryOverrides {
  const key = keyOf(kind)
  return { ...overrides, [key]: { ...(overrides[key] ?? {}), [item.id]: cloneJson(item as unknown as Json) } }
}

/** The overrides with a shipped item removed. */
export function withRemoval(overrides: LibraryOverrides, kind: Kind, id: string): LibraryOverrides {
  const key = keyOf(kind)
  return { ...overrides, [key]: { ...(overrides[key] ?? {}), [id]: null } }
}

/** The overrides with an item's change dropped: it goes back to the shipped one (or away, if it was added). */
export function withoutOverride(overrides: LibraryOverrides, kind: Kind, id: string): LibraryOverrides {
  const key = keyOf(kind)
  const next = { ...(overrides[key] ?? {}) }
  delete next[id]
  return { ...overrides, [key]: next }
}

/**
 * The library as the overrides change it, and each item's status. Order: the shipped items in their order (a removed
 * one left out), then the added ones in the overrides' order. A changed item equal to the shipped one is `shipped`.
 * `problems` lists recipes that use a component no longer there or extend a recipe no longer there, and, given the
 * application's code components, any JSON component whose id is a code component's (the compile refuses it).
 */
export function layerLibraries(shippedIn: Library, overridesIn: LibraryOverrides, options: { code?: CodeComponents } = {}): LayeredLibrary {
  const shipped = readLibrary(shippedIn)
  const overrides = readOverrides(overridesIn)
  const layer = <T extends { id: string }>(base: readonly T[], changes: Record<string, T | null> | undefined) => {
    const status: Record<string, OverrideStatus> = {}
    const items: T[] = []
    const changed = changes ?? {}
    for (const item of base) {
      if (!Object.prototype.hasOwnProperty.call(changed, item.id)) {
        status[item.id] = 'shipped'
        items.push(cloneJson(item as unknown as Json) as unknown as T)
        continue
      }
      const next = changed[item.id]
      if (next === null || next === undefined) {
        status[item.id] = 'removed'
        continue
      }
      status[item.id] = sameJson(next, item) ? 'shipped' : 'modified'
      items.push(cloneJson(next as unknown as Json) as unknown as T)
    }
    for (const [id, next] of Object.entries(changed)) {
      if (status[id] !== undefined || next === null) continue
      status[id] = 'added'
      items.push(cloneJson(next as unknown as Json) as unknown as T)
    }
    return { items, status }
  }
  const components = layer(shipped.components, overrides.components)
  const recipes = layer(shipped.recipes, overrides.recipes)
  const library: Library = { kind: 'library', version: RECIPE_FORMAT_VERSION, components: components.items, recipes: recipes.items }
  const problems: string[] = []
  const componentIds = new Set(library.components.map((c) => c.id))
  const known = (id: string) => componentIds.has(id) || options.code?.has(id) === true
  const recipeIds = new Set(library.recipes.map((r) => r.id))
  for (const recipe of library.recipes) {
    if (recipe.extends !== undefined && !recipeIds.has(recipe.extends)) problems.push(`The recipe "${recipe.id}" extends "${recipe.extends}", which is not in the library`)
    const used = new Set<string>()
    for (const section of recipe.sections ?? []) for (const part of section.parts) used.add(part.use)
    for (const change of recipe.changes ?? []) {
      if ('add' in change) used.add(change.add.use)
      if ('replace' in change) used.add(change.by.use)
    }
    for (const id of used) if (!known(id)) problems.push(`The recipe "${recipe.id}" uses "${id}", which is not in the library`)
  }
  for (const component of library.components) {
    for (const id of usesIn(component.blocks)) if (!known(id)) problems.push(`The component "${component.id}" uses "${id}", which is not in the library`)
    if (options.code?.has(component.id))
      problems.push(`The JSON component "${component.id}" has a code component's id; an override cannot replace a code component (give it another id)`)
  }
  return { library, status: { components: components.status, recipes: recipes.status }, problems }
}

/** The component ids a template's `use` nodes name. */
function usesIn(value: unknown, out = new Set<string>()): Set<string> {
  if (Array.isArray(value)) value.forEach((item) => usesIn(item, out))
  else if (value && typeof value === 'object') {
    const record = value as Record<string, unknown>
    if (typeof record.use === 'string' && !('type' in record)) out.add(record.use)
    for (const member of Object.values(record)) usesIn(member, out)
  }
  return out
}

export const OVERRIDES_FILE_KIND = 'document-kit-library-overrides'

/** The overrides as a file's text: canonical JSON, so the same changes always write the same file. */
export function overridesFileText(overrides: LibraryOverrides): string {
  return `${JSON.stringify(canonical(readOverrides(overrides)), null, 2)}\n`
}

/** Overrides read from a file's text, refused with every problem named. */
export function readOverridesFile(text: string): LibraryOverrides {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    throw new RecipeError('invalid', 'the file is not JSON')
  }
  return readOverrides(parsed)
}

/* ------------------------------------------------------------------ own recipes */

/** Own recipes' ids start with this, so one never takes a shipped recipe's id. */
export const OWN_RECIPE_PREFIX = 'own-'
export const RECIPE_FILE_KIND = 'document-kit-recipe'

/** A fresh own recipe id from a name; `taken` holds every id in use. `random` makes it repeatable in tests. */
export function ownRecipeId(name: string, taken: ReadonlySet<string>, random: () => number = Math.random): string {
  const slug = name.toLowerCase().normalize('NFKD').replace(/[^a-z0-9]+/g, '-').replace(/^-+|-+$/g, '').slice(0, 40).replace(/-+$/g, '') || 'recipe'
  for (;;) {
    const id = `${OWN_RECIPE_PREFIX}${slug}-${Math.floor(random() * 36 ** 4).toString(36).padStart(4, '0')}`
    if (!taken.has(id)) return id
  }
}

/**
 * A recipe to change as one's own: a full `copy` (its sections and changes duplicated), or an `extension` that
 * `extends` it with no changes yet, so later fixes to the original still reach it. Named "… (your copy)".
 */
export function copyRecipe(
  recipe: Recipe,
  taken: ReadonlySet<string>,
  options: { as?: 'copy' | 'extension'; name?: string; random?: () => number } = {},
): Recipe {
  const name = options.name ?? `${recipe.name.replace(/ \(your copy\)$/, '')} (your copy)`
  const id = ownRecipeId(name, taken, options.random)
  const basedOn = recipe.own ? recipe.own.basedOn : recipe.id
  if (options.as === 'extension') return { kind: 'recipe', id, name, extends: recipe.id, own: { basedOn: recipe.id } }
  const copy = cloneJson(recipe as unknown as Json) as unknown as Recipe
  return { ...copy, id, name, own: { basedOn } }
}

/**
 * A recipe with a change: its name, description, options or document replaced, and `changes` applied — straight to
 * its own sections, or added to its `changes` when it extends another (so they apply to the base). Checked on the way out.
 */
export function editRecipe(
  recipe: Recipe,
  change: { name?: string; description?: string | null; options?: JsonObject; document?: Recipe['document']; changes?: readonly RecipeChange[] },
): Recipe {
  const next: Recipe = cloneJson(recipe as unknown as Json) as unknown as Recipe
  if (change.name !== undefined) next.name = change.name
  if (change.description === null) delete next.description
  else if (change.description !== undefined) next.description = change.description
  if (change.options !== undefined) next.options = cloneJson(change.options)
  if (change.document !== undefined) next.document = cloneJson(change.document as Json) as Recipe['document'] & object
  if (change.changes?.length) {
    if (next.sections && next.extends === undefined) next.sections = applyChanges(next.sections, change.changes, `recipe/${recipe.id}`)
    else next.changes = [...(next.changes ?? []), ...cloneJson(change.changes as unknown as Json) as unknown as RecipeChange[]]
  }
  return readRecipe(next)
}

/** The part ids of a recipe's own sections, in order (an extension's are its base's, after compiling). */
export const recipePartIds = (recipe: Recipe): string[] => (recipe.sections ?? []).flatMap((section) => section.parts.map(partId))

/** A file sharing one recipe. */
export function recipeFileText(recipe: Recipe, exportedAt: Date): string {
  return `${JSON.stringify(canonical({ kind: RECIPE_FILE_KIND, version: RECIPE_FORMAT_VERSION, exportedAt: exportedAt.toISOString(), recipe: readRecipe(recipe) }), null, 2)}\n`
}

/**
 * Reads a shared recipe file. An imported recipe takes a fresh own id when its id is in use, and "(imported)" after
 * its name when that name is, so an import never overwrites. Refused with a `RecipeError` naming what is wrong.
 */
export function readRecipeFile(text: string, taken: { ids: ReadonlySet<string>; names: ReadonlySet<string> }, random?: () => number): Recipe {
  let parsed: unknown
  try {
    parsed = JSON.parse(text)
  } catch {
    throw new RecipeError('invalid', 'the file is not JSON')
  }
  if (typeof parsed !== 'object' || parsed === null || (parsed as { kind?: unknown }).kind !== RECIPE_FILE_KIND)
    throw new RecipeError('invalid', `the file is not a recipe file (its kind is not "${RECIPE_FILE_KIND}")`)
  const file = parsed as { version?: unknown; recipe?: unknown }
  if (typeof file.version !== 'number') throw new RecipeError('invalid', 'the recipe file gives no version')
  if (file.version > RECIPE_FORMAT_VERSION)
    throw new RecipeError('version', `the file is version ${file.version} of the recipe format; this kit reads version ${RECIPE_FORMAT_VERSION}`)
  let recipe = readRecipe(file.recipe)
  if (taken.names.has(recipe.name)) recipe = { ...recipe, name: `${recipe.name} (imported)` }
  if (taken.ids.has(recipe.id)) recipe = { ...recipe, id: ownRecipeId(recipe.name, taken.ids, random) }
  return recipe
}
