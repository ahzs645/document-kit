/**
 * The components the kit ships, as JSON (`src/recipe/components/*.json`), generic enough for any report: a cover page,
 * a table of contents, label and value fields, a text section, a notice, a data table, a figure, a figure page, a
 * spreadsheet, a signature, the to-complete list, a "Page X of Y" footer and a running header. An application's library starts from
 * them and adds its own.
 */
import contents from './components/contents.json'
import coverPage from './components/cover-page.json'
import dataTable from './components/data-table.json'
import fields from './components/fields.json'
import figurePage from './components/figure-page.json'
import figure from './components/figure.json'
import notice from './components/notice.json'
import pageFooter from './components/page-footer.json'
import runningHeader from './components/running-header.json'
import signature from './components/signature.json'
import spreadsheet from './components/spreadsheet.json'
import textSection from './components/text-section.json'
import toComplete from './components/to-complete.json'
import { cloneJson } from './json'
import { readComponent } from './read'
import type { Component, Json, Library } from './types'

const SHIPPED: readonly unknown[] = [
  coverPage,
  contents,
  fields,
  textSection,
  notice,
  dataTable,
  figure,
  figurePage,
  spreadsheet,
  signature,
  toComplete,
  pageFooter,
  runningHeader,
]

/** The shared components, each read and checked, as fresh copies. */
export function sharedComponents(): Component[] {
  return SHIPPED.map((component) => readComponent(cloneJson(component as Json)))
}

/** A library of the shared components and the components and recipes given (an application's own). */
export function sharedLibrary(own: Partial<Pick<Library, 'components' | 'recipes'>> = {}): Library {
  return { kind: 'library', version: 1, components: [...sharedComponents(), ...(own.components ?? [])], recipes: [...(own.recipes ?? [])] }
}
