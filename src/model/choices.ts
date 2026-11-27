/** A tick-box question as the kit's own blocks, the same for every writer. */
import type { Block, ChoicesBlock, Inline, Theme } from './types'

const isChosen = (block: ChoicesBlock, option: string) =>
  Array.isArray(block.chosen) ? block.chosen.includes(option) : block.chosen === option

/** A tick box and its option's words, kept together. */
function choiceRuns(block: ChoicesBlock, option: string): Inline[] {
  const source = block.source ? { source: block.source } : {}
  return [{ type: 'checkbox', checked: isChosen(block, option), ...source }, { type: 'text', text: ` ${option}`, ...source }]
}

/**
 * A tick-box question as the kit's own blocks: the question and its boxes on one line (`inline`), the question over one
 * option a line (`list`), or a ruled box with the question at the left and the boxes at the right (`boxed`). The
 * answer line, when there is one, goes under the question.
 */
export function choiceBlocks(block: ChoicesBlock, theme: Theme): Block[] {
  const answer: Block[] =
    block.answer === undefined
      ? []
      : [
          {
            type: 'paragraph',
            content: block.answer === null ? [{ type: 'blank' }] : block.answer,
            text: { bold: true },
            ...(block.source ? { source: block.source } : {}),
          },
        ]
  const label = block.label === undefined ? [] : [block.label]
  if (block.boxed) {
    const look = theme.table
    const options: Block[] = block.options.map((option) => ({ type: 'paragraph', content: choiceRuns(block, option), spacing: { after: 0 } }))
    return [
      {
        type: 'table',
        columnWidths: [3, 1],
        borders: { insideVertical: look.outerBorder },
        rows: [
          {
            cantSplit: true,
            cells: [
              { blocks: [...label.map((content): Block => ({ type: 'paragraph', content, spacing: { after: 2 } })), ...answer] },
              { blocks: options, verticalAlign: 'center' },
            ],
          },
        ],
      },
    ]
  }
  if ((block.layout ?? 'inline') === 'list') {
    return [
      ...label.map((content): Block => ({ type: 'paragraph', content, keepWithNext: true, spacing: { after: 2 } })),
      ...block.options.map((option, i): Block => ({
        type: 'paragraph',
        content: choiceRuns(block, option),
        indent: { left: 14 },
        spacing: { after: 0 },
        keepWithNext: i < block.options.length - 1 || answer.length > 0,
      })),
      ...answer,
    ]
  }
  const runs: Inline[] = []
  if (block.label !== undefined) runs.push(...(Array.isArray(block.label) ? block.label : [block.label as Inline]), '   ')
  block.options.forEach((option, i) => {
    if (i) runs.push('   ')
    runs.push(...choiceRuns(block, option))
  })
  return [{ type: 'paragraph', content: runs, keepWithNext: answer.length > 0 }, ...answer]
}

