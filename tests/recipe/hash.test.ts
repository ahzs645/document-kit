import { expect, it } from 'vitest'
import { sha256 } from '../../src/recipe/json'
import { createHash } from 'node:crypto'
it('hashes content with SHA-256, as Node does', () => {
  for (const t of ['abc', '', 'a'.repeat(55), 'a'.repeat(56), 'a'.repeat(64), 'é€😀'.repeat(40)]) expect(sha256(t)).toBe(createHash('sha256').update(t, 'utf8').digest('hex'))
})
