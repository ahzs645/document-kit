/**
 * document-kit/recipe: reports built from JSON. Components are JSON (the block model with `{{ tags }}` and control
 * nodes), a recipe is JSON that names components, and a recipe is compiled from its components into one
 * self-contained JSON object, then rendered against data into a kit `Document`. Code is left three jobs: computing the
 * data a recipe reads, registering custom block kinds and filters, and code components for parts that compute beyond
 * what tags say. The tag language is the Word template engine's (`document-kit/template`). Guide: docs/recipe.md.
 *
 * Runs anywhere: no Node, no DOM, no network.
 *
 * @module recipe
 */

export type * from './types'
export { RECIPE_FORMAT_VERSION } from './types'
export { isRecipeError, jsonPath, RecipeError, type RecipeErrorCode, type RecipeProblem } from './errors'
export { problemsOf, readCompiledRecipe, readComponent, readLibrary, readOverrides, readRecipe } from './read'
export {
  applyChanges,
  compileRecipe,
  COMPILE_LIMITS,
  partId,
  resolveRecipe,
  TO_COMPLETE_KIND,
  type CompileLimits,
  type CompileOptions,
} from './compile'
export {
  checkRecipe,
  renderRecipe,
  RENDER_LIMITS,
  type CheckProblem,
  type RecipeToCompleteItem,
  type RenderedPart,
  type RenderLimits,
  type RenderOptions,
  type RenderResult,
} from './render'
export { applyEdits, type EditKind, type EditTarget, type RecipeEdits } from './edits'
export {
  copyRecipe,
  editRecipe,
  emptyOverrides,
  layerLibraries,
  overridesFileText,
  OVERRIDES_FILE_KIND,
  OWN_RECIPE_PREFIX,
  ownRecipeId,
  readOverridesFile,
  readRecipeFile,
  recipeFileText,
  recipePartIds,
  RECIPE_FILE_KIND,
  withOverride,
  withoutOverride,
  withRemoval,
} from './library'
export {
  defineCodeComponent,
  defineCodeComponents,
  type CodeBlock,
  type CodeBuildInput,
  type CodeComponent,
  type CodeComponents,
} from './code'
export { sharedComponents, sharedLibrary } from './shared'
export { BUILT_IN_FILTERS } from './tags'
export { type FormatOptions, type RecipeFilter } from './values'
export { canonical, canonicalText, contentHash, sameJson } from './json'
