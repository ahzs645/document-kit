// The fixtures' one component library: the kit's shared components, the fixture components and the three recipes,
// read from the JSON files beside this one. Both report families draw on the same shared components.

import { readFileSync, readdirSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

import { caption, cell, defineCustomBlocks, headerRow, row, table } from '../../../src'
import { defineCodeComponents, sharedLibrary, type Component, type Library, type Recipe } from '../../../src/recipe'
import { reviewSummary } from './code'

const HERE = fileURLToPath(new URL('.', import.meta.url))
const readJson = (path: string): unknown => JSON.parse(readFileSync(path, 'utf8'))
const folder = (name: string) => readdirSync(join(HERE, name)).filter((f) => f.endsWith('.json')).sort().map((f) => readJson(join(HERE, name, f)))

export const fixtureComponents = (): Component[] => folder('components') as Component[]
export const fixtureRecipes = (): Recipe[] => folder('recipes') as Recipe[]
export const fixtureData = (name: 'laboratory' | 'review'): Record<string, unknown> => readJson(join(HERE, 'data', `${name}.json`)) as Record<string, unknown>

export function fixtureLibrary(): Library {
  return sharedLibrary({ components: fixtureComponents(), recipes: fixtureRecipes() })
}

export const code = defineCodeComponents([reviewSummary])

interface MatrixData {
  title: string
  rowHeader: string
  rows: string[]
  columns: string[]
  cells: string[][]
  mark?: { row: number; column: number }
}

/** The application's custom block: the matrix the review recipe's rating component holds. */
export const registry = defineCustomBlocks<{ matrix: MatrixData }>({
  matrix: (data) => [
    table(
      [
        headerRow([data.rowHeader, ...data.columns]),
        ...data.rows.map((label, r) =>
          row([
            cell(label, { text: { bold: true } }),
            ...data.columns.map((_, c) =>
              data.mark?.row === r && data.mark.column === c
                ? cell(data.cells[r]?.[c] ?? '', { shading: 'f4c7c3', text: { bold: true }, align: 'center' })
                : cell(data.cells[r]?.[c] ?? '', { align: 'center' }),
            ),
          ]),
        ),
      ],
      { caption: caption(data.title), ...(data.mark ? {} : { pending: [{ what: `${data.title}: the cell the assessment reads` }] }) },
    ),
  ],
})
