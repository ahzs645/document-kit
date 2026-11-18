// Small builders for the recipe tests: made-up components, recipes and libraries.

import type { Component, Json, JsonObject, Library, PartRef, Recipe } from '../../src/recipe'

export const comp = (id: string, blocks: Json[], extra: Partial<Component> = {}): Component => ({
  kind: 'component',
  id,
  title: extra.title ?? `Component ${id}`,
  ...extra,
  blocks,
})

export const recipe = (id: string, parts: PartRef[], extra: Partial<Recipe> = {}): Recipe => ({
  kind: 'recipe',
  id,
  name: `Recipe ${id}`,
  sections: [{ parts }],
  ...extra,
})

export const lib = (components: Component[], recipes: Recipe[] = []): Library => ({ components, recipes })

export const para = (content: Json, extra: JsonObject = {}): JsonObject => ({ type: 'paragraph', content, ...extra })
