/**
 * Reading recipe JSON safely, by hand (no schema library): a component, a recipe, a library, a set of overrides and a
 * compiled recipe. Every problem is collected with its JSON path and the whole set refused at once (`RecipeError`
 * with code `invalid` and `problems`), never repaired by guessing. A key the format does not have is a problem, so a
 * misspelt key is never silently dropped. A file of a newer format version is refused with code `version`. The block
 * templates inside a component are checked where they are used, by the compile (which knows each block's place).
 */
import { jsonPath, RecipeError, type RecipeProblem } from './errors'
import { isRecord } from './json'
import {
  RECIPE_FORMAT_VERSION,
  type CompiledRecipe,
  type Component,
  type Json,
  type Library,
  type LibraryOverrides,
  type PartRef,
  type Recipe,
  type RecipeChange,
  type RecipeSection,
} from './types'

type Path = ReadonlyArray<string | number>

/** Ids are words a person can type and a key can hold: letters, digits, `-`, `_` and `.`. */
export const ID_PATTERN = /^[A-Za-z0-9][A-Za-z0-9_.-]{0,99}$/

class Checker {
  readonly problems: RecipeProblem[] = []
  problem(path: Path, message: string) {
    this.problems.push({ path: jsonPath(path), message })
  }
  /** Reports keys of `value` not in `allowed`. */
  keys(value: Record<string, unknown>, allowed: readonly string[], path: Path, what: string) {
    for (const key of Object.keys(value))
      if (!allowed.includes(key)) this.problem([...path, key], `${what} has no key "${key}" (it takes ${allowed.join(', ')})`)
  }
  string(value: unknown, path: Path, what: string, optional = false): value is string {
    if (value === undefined && optional) return false
    if (typeof value === 'string' && value.trim() !== '') return true
    this.problem(path, `${what} is not a non-empty string`)
    return false
  }
  id(value: unknown, path: Path, what: string, optional = false): value is string {
    if (value === undefined && optional) return false
    if (typeof value === 'string' && ID_PATTERN.test(value)) return true
    this.problem(path, `${what} is not an id (letters, digits, - _ . and at most 100 of them): ${JSON.stringify(value)}`)
    return false
  }
  /** `meta`: any JSON object, read as JSON and nothing more. */
  meta(value: unknown, path: Path) {
    if (value === undefined) return
    if (this.object(value, path, 'meta')) this.json(value, path, 'meta')
  }
  /** A missing-value rule. */
  missing(value: unknown, path: Path) {
    if (value !== undefined && value !== 'blank' && value !== 'placeholder') this.problem(path, 'missing is "blank" or "placeholder"')
  }
  /** A part's pin: "bottom", or { at: "bottom", spaceBelow }. */
  pin(value: unknown, path: Path) {
    if (value === undefined || value === 'bottom') return
    if (!this.object(value, path, 'pin')) return
    this.keys(value, ['at', 'spaceBelow'], path, 'pin')
    if (value.at !== undefined && value.at !== 'bottom') this.problem([...path, 'at'], 'pin.at is "bottom"')
    if (value.spaceBelow !== undefined && (typeof value.spaceBelow !== 'number' || !(value.spaceBelow >= 0))) this.problem([...path, 'spaceBelow'], 'pin.spaceBelow is a number of points from 0')
  }
  bool(value: unknown, path: Path, what: string) {
    if (value !== undefined && typeof value !== 'boolean') this.problem(path, `${what} is not true or false`)
  }
  json(value: unknown, path: Path, what: string, depth = 0) {
    if (depth > 200) return this.problem(path, `${what} nests more than 200 deep`)
    if (value === null || typeof value === 'string' || typeof value === 'boolean') return
    if (typeof value === 'number') {
      if (!Number.isFinite(value)) this.problem(path, `${what} holds a number that is not finite`)
      return
    }
    if (Array.isArray(value)) return value.forEach((item, i) => this.json(item, [...path, i], what, depth + 1))
    if (isRecord(value)) {
      for (const [key, member] of Object.entries(value)) this.json(member, [...path, key], what, depth + 1)
      return
    }
    this.problem(path, `${what} holds something that is not JSON (${typeof value})`)
  }
  blocks(value: unknown, path: Path, what: string, optional = false) {
    if (value === undefined && optional) return
    if (!Array.isArray(value)) return this.problem(path, `${what} is not a list of blocks`)
    this.json(value, path, what)
  }
  object(value: unknown, path: Path, what: string, optional = false): value is Record<string, unknown> {
    if (value === undefined && optional) return false
    if (isRecord(value)) return true
    this.problem(path, `${what} is not an object`)
    return false
  }
  version(value: Record<string, unknown>, path: Path, what: string) {
    const version = value.version
    if (version === undefined) return
    if (typeof version !== 'number' || !Number.isInteger(version) || version < 1)
      return this.problem([...path, 'version'], `${what}'s version is not a whole number from 1`)
    if (version > RECIPE_FORMAT_VERSION)
      throw new RecipeError(
        'version',
        `${what} is version ${version} of the recipe format; this kit reads version ${RECIPE_FORMAT_VERSION}`,
        jsonPath([...path, 'version']),
      )
  }
  done<T>(value: T, what: string): T {
    if (this.problems.length) {
      const first = this.problems[0]
      const more = this.problems.length > 1 ? ` and ${this.problems.length - 1} more` : ''
      throw new RecipeError('invalid', `${what} does not read: ${first?.message}${more}`, first?.path ?? '', this.problems)
    }
    return value
  }
}

const COMPONENT_KEYS = ['kind', 'version', 'id', 'title', 'description', 'params', 'variants', 'blocks', 'meta']
const PARAM_KEYS = ['default', 'required', 'description']
const VARIANT_KEYS = ['title', 'description', 'blocks', 'with']

function checkComponent(c: Checker, value: unknown, path: Path) {
  if (!c.object(value, path, 'A component')) return
  c.version(value, path, 'The component')
  c.keys(value, COMPONENT_KEYS, path, 'A component')
  if (value.kind !== 'component') c.problem([...path, 'kind'], 'A component\'s kind is not "component"')
  c.id(value.id, [...path, 'id'], 'The component\'s id')
  c.string(value.title, [...path, 'title'], 'The component\'s title')
  if (value.description !== undefined && typeof value.description !== 'string')
    c.problem([...path, 'description'], 'The description is not a string')
  if (c.object(value.params, [...path, 'params'], 'The params', true)) {
    for (const [name, spec] of Object.entries(value.params)) {
      const at = [...path, 'params', name]
      if (!/^[A-Za-z_][A-Za-z0-9_]*$/.test(name)) c.problem(at, `The parameter name "${name}" is not a name (letters, digits, _)`)
      if (!c.object(spec, at, `The parameter "${name}"`)) continue
      c.keys(spec, PARAM_KEYS, at, `The parameter "${name}"`)
      c.bool(spec.required, [...at, 'required'], 'required')
      if (spec.default !== undefined) c.json(spec.default, [...at, 'default'], 'The default')
      if (spec.required === true && spec.default !== undefined)
        c.problem(at, `The parameter "${name}" is required and has a default; it is one or the other`)
    }
  }
  if (c.object(value.variants, [...path, 'variants'], 'The variants', true)) {
    for (const [id, variant] of Object.entries(value.variants)) {
      const at = [...path, 'variants', id]
      c.id(id, at, 'A variant\'s id')
      if (!c.object(variant, at, `The variant "${id}"`)) continue
      c.keys(variant, VARIANT_KEYS, at, `The variant "${id}"`)
      c.string(variant.title, [...at, 'title'], 'The variant\'s title')
      c.blocks(variant.blocks, [...at, 'blocks'], 'The variant\'s blocks', true)
      if (c.object(variant.with, [...at, 'with'], 'The variant\'s with', true)) c.json(variant.with, [...at, 'with'], 'with')
      if (variant.blocks === undefined && variant.with === undefined)
        c.problem(at, `The variant "${id}" gives neither blocks nor with`)
    }
  }
  c.blocks(value.blocks, [...path, 'blocks'], 'The component\'s blocks')
  c.meta(value.meta, [...path, 'meta'])
}

const PART_KEYS = ['use', 'id', 'title', 'variant', 'with', 'if', 'for', 'pageBreakBefore', 'missing', 'pin', 'meta']

function checkPart(c: Checker, value: unknown, path: Path) {
  if (!c.object(value, path, 'A part')) return
  c.keys(value, PART_KEYS, path, 'A part')
  c.id(value.use, [...path, 'use'], 'The component a part uses')
  c.id(value.id, [...path, 'id'], 'The part\'s id', true)
  c.string(value.title, [...path, 'title'], 'The part\'s title', true)
  c.id(value.variant, [...path, 'variant'], 'The part\'s variant', true)
  if (c.object(value.with, [...path, 'with'], 'The part\'s with', true)) c.json(value.with, [...path, 'with'], 'with')
  c.string(value.if, [...path, 'if'], 'The part\'s condition', true)
  c.string(value.for, [...path, 'for'], 'The part\'s loop', true)
  c.bool(value.pageBreakBefore, [...path, 'pageBreakBefore'], 'pageBreakBefore')
  c.missing(value.missing, [...path, 'missing'])
  c.pin(value.pin, [...path, 'pin'])
  c.meta(value.meta, [...path, 'meta'])
}

const SECTION_KEYS = ['id', 'page', 'header', 'footer', 'firstPage', 'pageNumbers', 'parts']

function checkSection(c: Checker, value: unknown, path: Path, partsOptional = false) {
  if (!c.object(value, path, 'A section')) return
  c.keys(value, SECTION_KEYS, path, 'A section')
  c.id(value.id, [...path, 'id'], 'The section\'s id', true)
  if (value.page !== undefined) c.json(value.page, [...path, 'page'], 'The page')
  if (value.pageNumbers !== undefined) c.json(value.pageNumbers, [...path, 'pageNumbers'], 'The page numbers')
  c.blocks(value.header, [...path, 'header'], 'The header', true)
  c.blocks(value.footer, [...path, 'footer'], 'The footer', true)
  if (c.object(value.firstPage, [...path, 'firstPage'], 'The first page', true)) {
    c.keys(value.firstPage, ['header', 'footer'], [...path, 'firstPage'], 'The first page')
    c.blocks(value.firstPage.header, [...path, 'firstPage', 'header'], 'The first page\'s header', true)
    c.blocks(value.firstPage.footer, [...path, 'firstPage', 'footer'], 'The first page\'s footer', true)
  }
  if (value.parts === undefined && partsOptional) return
  if (!Array.isArray(value.parts)) return c.problem([...path, 'parts'], 'The section\'s parts are not a list')
  value.parts.forEach((part, i) => checkPart(c, part, [...path, 'parts', i]))
}

const ANCHOR_KEYS = ['before', 'after', 'section']

function checkAnchor(c: Checker, value: Record<string, unknown>, path: Path) {
  c.id(value.before, [...path, 'before'], 'before', true)
  c.id(value.after, [...path, 'after'], 'after', true)
  if (value.before !== undefined && value.after !== undefined) c.problem(path, 'A change gives both before and after')
  if (value.section !== undefined && typeof value.section !== 'number' && !(typeof value.section === 'string' && ID_PATTERN.test(value.section)))
    c.problem([...path, 'section'], 'section is not a section id or index')
}

function checkChange(c: Checker, value: unknown, path: Path) {
  if (!c.object(value, path, 'A change')) return
  if ('add' in value) {
    c.keys(value, ['add', ...ANCHOR_KEYS], path, 'An add change')
    checkPart(c, value.add, [...path, 'add'])
    checkAnchor(c, value, path)
  } else if ('remove' in value) {
    c.keys(value, ['remove'], path, 'A remove change')
    c.id(value.remove, [...path, 'remove'], 'The part removed')
  } else if ('replace' in value) {
    c.keys(value, ['replace', 'by'], path, 'A replace change')
    c.id(value.replace, [...path, 'replace'], 'The part replaced')
    checkPart(c, value.by, [...path, 'by'])
  } else if ('move' in value) {
    c.keys(value, ['move', ...ANCHOR_KEYS], path, 'A move change')
    c.id(value.move, [...path, 'move'], 'The part moved')
    checkAnchor(c, value, path)
  } else if ('set' in value) {
    c.keys(value, ['set', 'variant', 'with', 'if', 'for', 'title', 'pageBreakBefore', 'missing', 'pin', 'meta'], path, 'A set change')
    c.id(value.set, [...path, 'set'], 'The part set')
    if (value.missing !== null) c.missing(value.missing, [...path, 'missing'])
    if (value.pin !== null) c.pin(value.pin, [...path, 'pin'])
    if (value.meta !== null) c.meta(value.meta, [...path, 'meta'])
    if (value.with !== undefined && c.object(value.with, [...path, 'with'], 'with')) c.json(value.with, [...path, 'with'], 'with')
    for (const key of ['variant', 'if', 'for', 'title'])
      if (value[key] !== undefined && value[key] !== null && typeof value[key] !== 'string')
        c.problem([...path, key], `${key} is not a string or null`)
  } else if ('setSection' in value) {
    c.keys(value, ['setSection', 'page', 'header', 'footer', 'firstPage', 'pageNumbers'], path, 'A setSection change')
    for (const key of ['page', 'header', 'footer', 'firstPage', 'pageNumbers'])
      if (value[key] !== undefined) c.json(value[key], [...path, key], key)
  } else if ('addSection' in value) {
    c.keys(value, ['addSection', 'after'], path, 'An addSection change')
    checkSection(c, value.addSection, [...path, 'addSection'])
  } else c.problem(path, 'A change is none of add, remove, replace, move, set, setSection, addSection')
}

const RECIPE_KEYS = ['kind', 'version', 'id', 'name', 'description', 'extends', 'document', 'options', 'sections', 'changes', 'own', 'missing', 'meta']

function checkRecipe(c: Checker, value: unknown, path: Path) {
  if (!c.object(value, path, 'A recipe')) return
  c.version(value, path, 'The recipe')
  c.keys(value, RECIPE_KEYS, path, 'A recipe')
  if (value.kind !== 'recipe') c.problem([...path, 'kind'], 'A recipe\'s kind is not "recipe"')
  c.id(value.id, [...path, 'id'], 'The recipe\'s id')
  c.string(value.name, [...path, 'name'], 'The recipe\'s name')
  if (value.description !== undefined && typeof value.description !== 'string')
    c.problem([...path, 'description'], 'The description is not a string')
  c.id(value.extends, [...path, 'extends'], 'The recipe it extends', true)
  if (c.object(value.document, [...path, 'document'], 'The document', true)) {
    c.keys(value.document, ['metadata', 'theme', 'textRules', 'footnotes'], [...path, 'document'], 'The document')
    if (c.object(value.document.footnotes, [...path, 'document', 'footnotes'], 'The footnotes', true)) {
      for (const [id, paragraphs] of Object.entries(value.document.footnotes)) {
        if (!Array.isArray(paragraphs) || !paragraphs.length || !paragraphs.every(p => typeof p === 'string' && p.trim()))
          c.problem([...path, 'document', 'footnotes', id], 'A footnote needs one or more nonempty paragraphs')
      }
    }
    if (value.document.textRules !== undefined) {
      if (!Array.isArray(value.document.textRules)) c.problem([...path, 'document', 'textRules'], 'The text rules are not a list')
      else c.json(value.document.textRules, [...path, 'document', 'textRules'], 'The text rules')
    }
    if (c.object(value.document.metadata, [...path, 'document', 'metadata'], 'The metadata', true))
      c.json(value.document.metadata, [...path, 'document', 'metadata'], 'The metadata')
    if (c.object(value.document.theme, [...path, 'document', 'theme'], 'The theme', true))
      c.json(value.document.theme, [...path, 'document', 'theme'], 'The theme')
  }
  if (c.object(value.options, [...path, 'options'], 'The options', true)) c.json(value.options, [...path, 'options'], 'The options')
  c.missing(value.missing, [...path, 'missing'])
  c.meta(value.meta, [...path, 'meta'])
  if (value.sections !== undefined) {
    if (!Array.isArray(value.sections) || !value.sections.length) c.problem([...path, 'sections'], 'The sections are not a list of at least one')
    else value.sections.forEach((section, i) => checkSection(c, section, [...path, 'sections', i]))
  } else if (value.extends === undefined) c.problem([...path, 'sections'], 'A recipe that extends none gives its sections')
  if (value.changes !== undefined) {
    if (!Array.isArray(value.changes)) c.problem([...path, 'changes'], 'The changes are not a list')
    else value.changes.forEach((change, i) => checkChange(c, change, [...path, 'changes', i]))
  }
  if (value.own !== undefined) {
    if (!c.object(value.own, [...path, 'own'], 'own')) return
    c.keys(value.own, ['basedOn'], [...path, 'own'], 'own')
    if (value.own.basedOn !== null && typeof value.own.basedOn !== 'string')
      c.problem([...path, 'own', 'basedOn'], 'basedOn is not an id or null')
  }
}

/** A component read from JSON, or a `RecipeError` naming every problem. */
export function readComponent(value: unknown): Component {
  const c = new Checker()
  checkComponent(c, value, [])
  return c.done(value as Component, 'The component')
}

/** A recipe read from JSON, or a `RecipeError` naming every problem. */
export function readRecipe(value: unknown): Recipe {
  const c = new Checker()
  checkRecipe(c, value, [])
  return c.done(value as Recipe, 'The recipe')
}

/** A library read from JSON: components and recipes, each id once. */
export function readLibrary(value: unknown): Library {
  const c = new Checker()
  if (c.object(value, [], 'A library')) {
    c.version(value, [], 'The library')
    c.keys(value, ['kind', 'version', 'components', 'recipes'], [], 'A library')
    if (value.kind !== undefined && value.kind !== 'library') c.problem(['kind'], 'A library\'s kind is not "library"')
    for (const [key, check, what] of [
      ['components', checkComponent, 'component'],
      ['recipes', checkRecipe, 'recipe'],
    ] as const) {
      const list = value[key]
      if (list === undefined) continue
      if (!Array.isArray(list)) {
        c.problem([key], `The ${key} are not a list`)
        continue
      }
      const seen = new Set<string>()
      list.forEach((item: unknown, i: number) => {
        check(c, item, [key, i])
        const id = isRecord(item) ? item.id : undefined
        if (typeof id === 'string') {
          if (seen.has(id)) c.problem([key, i, 'id'], `Two ${what}s have the id "${id}"`)
          seen.add(id)
        }
      })
    }
  }
  const library = value as Partial<Library>
  return c.done({ ...library, components: library.components ?? [], recipes: library.recipes ?? [] } as Library, 'The library')
}

/** A set of overrides read from JSON: each component or recipe as it should be, or null to remove the shipped one. */
export function readOverrides(value: unknown): LibraryOverrides {
  const c = new Checker()
  if (c.object(value, [], 'Overrides')) {
    c.version(value, [], 'The overrides')
    c.keys(value, ['kind', 'version', 'components', 'recipes'], [], 'Overrides')
    if (value.kind !== 'library-overrides') c.problem(['kind'], 'The overrides\' kind is not "library-overrides"')
    if (typeof value.version !== 'number') c.problem(['version'], 'The overrides give no version')
    for (const [key, check] of [
      ['components', checkComponent],
      ['recipes', checkRecipe],
    ] as const) {
      if (!c.object(value[key], [key], `The ${key}`, true)) continue
      for (const [id, item] of Object.entries(value[key] as Record<string, unknown>)) {
        c.id(id, [key, id], 'An overridden id')
        if (item === null) continue
        check(c, item, [key, id])
        if (isRecord(item) && item.id !== id) c.problem([key, id, 'id'], `The item under "${id}" has the id "${String(item.id)}"`)
      }
    }
  }
  return c.done(value as LibraryOverrides, 'The overrides')
}

const COMPILED_KEYS = ['kind', 'version', 'recipe', 'from', 'missing', 'document', 'options', 'sections', 'needs', 'customKinds', 'filters', 'code', 'embedded']
const COMPILED_PART_KEYS = ['part', 'if', 'for', 'pageBreakBefore', 'variant', 'missing', 'pin', 'params', 'blocks', 'code']

/**
 * A compiled recipe read back (from storage or a file): its shape checked, a newer version refused. The block
 * templates inside are checked again when it is rendered.
 */
export function readCompiledRecipe(value: unknown): CompiledRecipe {
  const c = new Checker()
  if (c.object(value, [], 'A compiled recipe')) {
    if (typeof value.version !== 'number') c.problem(['version'], 'A compiled recipe gives no version')
    c.version(value, [], 'The compiled recipe')
    c.keys(value, COMPILED_KEYS, [], 'A compiled recipe')
    if (value.kind !== 'compiled-recipe') c.problem(['kind'], 'A compiled recipe\'s kind is not "compiled-recipe"')
    if (c.object(value.recipe, ['recipe'], 'The recipe')) {
      c.id(value.recipe.id, ['recipe', 'id'], 'The recipe\'s id')
      c.string(value.recipe.name, ['recipe', 'name'], 'The recipe\'s name')
    }
    if (c.object(value.from, ['from'], 'from')) c.string(value.from.hash, ['from', 'hash'], 'The hash')
    c.missing(value.missing, ['missing'])
    c.object(value.document, ['document'], 'The document')
    c.object(value.options, ['options'], 'The options')
    for (const key of ['needs', 'customKinds', 'filters', 'code'])
      if (!Array.isArray(value[key])) c.problem([key], `${key} is not a list`)
    if (!Array.isArray(value.sections) || !value.sections.length) c.problem(['sections'], 'The sections are not a list of at least one')
    else
      value.sections.forEach((section: unknown, s: number) => {
        if (!c.object(section, ['sections', s], 'A section')) return
        c.keys(section, SECTION_KEYS, ['sections', s], 'A section')
        if (!Array.isArray(section.parts)) return c.problem(['sections', s, 'parts'], 'The parts are not a list')
        section.parts.forEach((part: unknown, p: number) => {
          const at = ['sections', s, 'parts', p]
          if (!c.object(part, at, 'A part')) return
          c.keys(part, COMPILED_PART_KEYS, at, 'A compiled part')
          c.missing(part.missing, [...at, 'missing'])
          c.pin(part.pin, [...at, 'pin'])
          if (c.object(part.part, [...at, 'part'], 'The part\'s mark')) {
            c.keys(part.part, ['id', 'component', 'title', 'meta', 'componentMeta'], [...at, 'part'], 'A part\'s mark')
            c.id(part.part.id, [...at, 'part', 'id'], 'The part\'s id')
            c.id(part.part.component, [...at, 'part', 'component'], 'The part\'s component')
            c.string(part.part.title, [...at, 'part', 'title'], 'The part\'s title')
          }
          c.object(part.params, [...at, 'params'], 'The params')
          if ((part.blocks === undefined) === (part.code === undefined))
            c.problem(at, 'A compiled part holds either blocks or a code component, and only one')
          c.blocks(part.blocks, [...at, 'blocks'], 'The blocks', true)
        })
      })
    if (c.object(value.embedded, ['embedded'], 'embedded', true))
      for (const [id, component] of Object.entries(value.embedded)) checkComponent(c, component, ['embedded', id])
    c.json(value, [], 'A compiled recipe')
  }
  return c.done(value as CompiledRecipe, 'The compiled recipe')
}

/** The problems of a value read as a component, recipe or library, without throwing: empty when it reads. */
export function problemsOf(read: (value: unknown) => unknown, value: unknown): RecipeProblem[] {
  try {
    read(value)
    return []
  } catch (error) {
    if (error instanceof RecipeError) return error.problems.length ? [...error.problems] : [{ path: error.path, message: error.message }]
    throw error
  }
}

export type { Json, PartRef, RecipeChange, RecipeSection }
