/** A section's page in points: its paper turned to its orientation, its margins, and the box text fills. */
import { PAPER } from './theme'
import type { Margins, PageSetup, Theme } from './types'

export interface PageBox {
  /** Page width as printed (landscape: the long side), points. */
  width: number
  height: number
  orientation: 'portrait' | 'landscape'
  margins: Margins
  /** The text width between the margins, points. */
  contentWidth: number
  /** The text height between the top and bottom margins, points. */
  contentHeight: number
}

export function pageBox(setup: PageSetup | undefined, theme: Theme): PageBox {
  const paperSpec = setup?.paper ?? theme.page.paper
  const paper = typeof paperSpec === 'string' ? PAPER[paperSpec] : paperSpec
  if (!paper || !(paper.width > 0) || !(paper.height > 0)) {
    throw new Error(`document-kit: unknown paper ${JSON.stringify(paperSpec)}`)
  }
  const orientation = setup?.orientation ?? theme.page.orientation
  const short = Math.min(paper.width, paper.height)
  const long = Math.max(paper.width, paper.height)
  const width = orientation === 'landscape' ? long : short
  const height = orientation === 'landscape' ? short : long
  const margins: Margins = { ...theme.page.margins, ...definedOnly(setup?.margins) }
  return {
    width,
    height,
    orientation,
    margins,
    contentWidth: width - margins.left - margins.right,
    contentHeight: height - margins.top - margins.bottom,
  }
}

function definedOnly<T extends object>(value: T | undefined): Partial<T> {
  if (!value) return {}
  return Object.fromEntries(Object.entries(value).filter(([, v]) => v !== undefined)) as Partial<T>
}
