/** Pure helpers over inline text: reading it as plain text, filling sentence slots, dropping unfilled sentences. */
import type { Inline, RichText, SentenceRun, SlotValues } from './types'

/** Rich text as a list of runs. */
export function inlines(text: RichText): readonly Inline[] {
  if (typeof text === 'string') return [text]
  return isInlineList(text) ? text : [text]
}

const isInlineList = (text: RichText): text is readonly Inline[] => Array.isArray(text)

/** True when a value is rich text (a string, a run, or a list of runs) rather than an object of another kind. */
export function isRichText(value: unknown): value is RichText {
  if (typeof value === 'string' || Array.isArray(value)) return true
  return typeof value === 'object' && value !== null && typeof (value as { type?: unknown }).type === 'string'
}

export interface FilledSentence {
  /** The sentence with its slots filled, or null when a slot has no value. */
  text: string | null
  /** The slots with no value, in the order they appear, each once. */
  missing: string[]
}

const SLOT = /\{(\w+)\}/g

/**
 * Fills `{slot}`s from `values`. A slot whose value is null, undefined or an empty string leaves the whole sentence
 * unfilled (`text: null`): a sentence is never printed with a hole or a brace in it.
 */
export function fillSlots(template: string, values: SlotValues): FilledSentence {
  const missing: string[] = []
  const text = template.replace(SLOT, (_, key: string) => {
    const value = values[key]
    if (value === null || value === undefined || value === '') {
      if (!missing.includes(key)) missing.push(key)
      return ''
    }
    return String(value)
  })
  return { text: missing.length ? null : text, missing }
}

/** A sentence run's printed text, or null when it is left out. */
export const sentenceText = (run: SentenceRun): string | null => fillSlots(run.template, run.values).text

const isSpace = (inline: Inline | undefined): inline is string => typeof inline === 'string' && /^\s*$/.test(inline)

/**
 * The runs as printed: each sentence filled into a text run, and each left-out sentence dropped with the space that
 * separated it from its neighbour, so "A. B. C." without B reads "A. C.". Returns an empty list when nothing is left.
 */
export function resolveSentences(text: RichText): Inline[] {
  const list = inlines(text)
  if (!list.some((inline) => typeof inline !== 'string' && inline.type === 'sentence')) return [...list]
  const out: Inline[] = []
  let dropSpaceAfter = false
  for (const inline of list) {
    if (typeof inline !== 'string' && inline.type === 'sentence') {
      const filled = sentenceText(inline)
      if (filled === null) {
        // Drop the separator after this sentence, or the one before it when it ends the run.
        dropSpaceAfter = true
        continue
      }
      const { type: _type, template: _template, values: _values, label: _label, ...style } = inline
      out.push({ ...style, type: 'text', text: filled })
      dropSpaceAfter = false
      continue
    }
    if (dropSpaceAfter && isSpace(inline)) {
      dropSpaceAfter = false
      continue
    }
    dropSpaceAfter = false
    out.push(inline)
  }
  // A dropped last sentence leaves the space before it hanging.
  while (out.length && isSpace(out[out.length - 1])) out.pop()
  while (out.length && isSpace(out[0])) out.shift()
  return out
}

/** Rich text as plain text: fields read as `#`, a blank as underscores, an image as its alt text. */
export function plainText(text: RichText | null | undefined): string {
  if (text === null || text === undefined) return ''
  return inlines(text)
    .map((inline) => {
      if (typeof inline === 'string') return inline
      switch (inline.type) {
        case 'text':
          return inline.text
        case 'field':
          return '#'
        case 'break':
          return '\n'
        case 'tab':
          return '\t'
        case 'blank':
          return '______'
        case 'image':
          return inline.alt ?? ''
        case 'sentence':
          return sentenceText(inline) ?? ''
      }
    })
    .join('')
}

/** True when rich text prints nothing at all (no runs, or only empty strings). */
export function isEmptyText(text: RichText | null | undefined): boolean {
  if (text === null || text === undefined) return true
  return inlines(text).every((inline) => inline === '')
}
