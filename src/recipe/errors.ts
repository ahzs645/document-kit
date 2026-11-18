/** The recipe system's one error type: a code saying what went wrong, and where, as a JSON path. */

export type RecipeErrorCode =
  /** A JSON file that does not read: every problem is in `problems`. */
  | 'invalid'
  /** A file of a newer format version than this kit reads. */
  | 'version'
  | 'unknown-recipe'
  | 'unknown-component'
  | 'unknown-variant'
  /** A change names a part or section the recipe does not have. */
  | 'unknown-part'
  | 'missing-param'
  | 'unknown-param'
  /** A component that uses itself, or a recipe that extends itself, through any chain. */
  | 'cycle'
  /** A `{{ tag }}`, a condition or a loop that does not read. */
  | 'bad-tag'
  /** A key the block model does not have at that place. */
  | 'unknown-key'
  /** A node that is not what its place takes (a block with no type, a row that is a number). */
  | 'bad-node'
  | 'duplicate-part'
  /** A JSON component and a code component with the same id. */
  | 'component-conflict'
  | 'limit'
  | 'unknown-filter'
  /** A code component the render or the compile was not given. */
  | 'missing-code'
  | 'missing-custom-kind'
  /** A list or an object printed where text goes. */
  | 'not-text'
  /** A value a filter or a place cannot take. */
  | 'bad-value'
  /** A code component's build threw or returned something that is not blocks. */
  | 'code-failed'
  /** Blocks from a code component that `use` a component it did not declare. */
  | 'undeclared-use'

export interface RecipeProblem {
  path: string
  message: string
}

export class RecipeError extends Error {
  readonly code: RecipeErrorCode
  /** The JSON path of the place (in the file read, the recipe compiled or the compiled recipe rendered). */
  readonly path: string
  /** For `invalid`: every problem found. */
  readonly problems: readonly RecipeProblem[]

  constructor(code: RecipeErrorCode, message: string, path = '', problems: readonly RecipeProblem[] = []) {
    super(`document-kit/recipe: ${message}${path ? ` (at ${path})` : ''}`)
    this.name = 'RecipeError'
    this.code = code
    this.path = path
    this.problems = problems
  }
}

export function isRecipeError(error: unknown, code?: RecipeErrorCode): error is RecipeError {
  return error instanceof RecipeError && (code === undefined || error.code === code)
}

/** A JSON path from its steps: `sections[0].parts[2].with.title`; a key that is not a name is quoted. */
export function jsonPath(steps: ReadonlyArray<string | number>): string {
  let out = ''
  for (const step of steps) {
    if (typeof step === 'number') out += `[${step}]`
    else if (/^[A-Za-z_$][\w$-]*$/.test(step)) out += out ? `.${step}` : step
    else out += `[${JSON.stringify(step)}]`
  }
  return out
}
