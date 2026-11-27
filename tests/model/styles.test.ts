// The named styles of a document: the theme's and those it defines, resolved through what they are based on.
import { describe, expect, it } from 'vitest'

import { paragraphStyleId, resolveTheme, StyleError, styleTable, characterStyleId } from '../../src/model'

const theme = resolveTheme({ body: { font: 'Brand Sans', size: 10, color: '111111' }, title: { size: 28 } })

describe('styleTable', () => {
  it('resolves a defined style through its bases down to the body text', () => {
    const table = styleTable(theme, {
      paragraph: [
        { id: 'Dept', text: { size: 12, bold: true, characterSpacing: 0.5 } },
        { id: 'SubDept', basedOn: 'Dept', next: 'normal', text: { size: 10.5, bold: false } },
        { id: 'CoverTitle', basedOn: 'title', text: { color: 'ffffff' } },
      ],
      character: [
        { id: 'Accent', text: { bold: true } },
        { id: 'AccentRed', basedOn: 'Accent', text: { color: 'aa0000' } },
      ],
    })
    expect(table.paragraph.get('SubDept')).toMatchObject({
      basedOn: 'Dept',
      next: 'Normal',
      defined: true,
      text: { font: 'Brand Sans', size: 10.5, bold: false, characterSpacing: 0.5, color: '111111' },
    })
    expect(table.paragraph.get('CoverTitle')?.text).toMatchObject({ size: 28, color: 'ffffff', bold: true })
    expect(table.character.get('AccentRed')?.text).toEqual({ bold: true, color: 'aa0000' })
    expect(paragraphStyleId(table, 'SubDept')).toBe('SubDept')
    expect(paragraphStyleId(table, 'note')).toBe('KitNote')
    expect(paragraphStyleId(table, 'normal')).toBeUndefined()
    expect(characterStyleId(table, 'Accent')).toBe('Accent')
  })

  it('refuses malformed, reserved, doubled, unknown and circular styles, naming them', () => {
    expect(() => styleTable(theme, { paragraph: [{ id: '1st' }] })).toThrow(StyleError)
    expect(() => styleTable(theme, { paragraph: [{ id: 'Title' }] })).toThrow(/theme's own/)
    expect(() => styleTable(theme, { paragraph: [{ id: 'note' }] })).toThrow(/theme's own/)
    expect(() => styleTable(theme, { paragraph: [{ id: 'A' }], character: [{ id: 'a', text: {} }] })).toThrow(/defined twice/)
    expect(() => styleTable(theme, { paragraph: [{ id: 'A', basedOn: 'Nope' }] })).toThrow(/"Nope"/)
    expect(() => styleTable(theme, { paragraph: [{ id: 'A', basedOn: 'B' }, { id: 'B', basedOn: 'A' }] })).toThrow(/based on itself/)
    expect(() => styleTable(theme, { character: [{ id: 'A', basedOn: 'Gone', text: {} }] })).toThrow(/"Gone"/)
    const table = styleTable(theme, undefined)
    expect(() => paragraphStyleId(table, 'Missing')).toThrow(/"Missing"/)
    expect(() => characterStyleId(table, 'Missing')).toThrow(/"Missing"/)
  })
})
