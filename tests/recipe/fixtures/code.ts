// The review fixture's one purpose-built part: a summary that computes from several data values (how many
// viewpoints see the opening, the largest share seen, the nearest one), written as a code component beside the JSON
// ones. Its blocks are ordinary block templates: tags read its parameters and the data with provenance, and when a
// share passes a limit it uses the shared JSON `notice` component, which it declares in `uses`.

import { defineCodeComponent, type CodeComponent } from '../../../src/recipe'

interface Viewpoint {
  name: string
  share?: number
  distanceKm?: number
}

export const reviewSummary: CodeComponent = defineCodeComponent({
  id: 'review-summary',
  title: 'Summary',
  version: '1',
  params: { assessor: { required: true }, limit: { default: 0.15 } },
  needs: ['viewpoints[].name', 'viewpoints[].share', 'viewpoints[].distanceKm'],
  uses: ['notice'],
  build({ params, value }) {
    const viewpoints = (value('viewpoints') as Viewpoint[] | undefined) ?? []
    const seen = viewpoints.filter((vp) => (vp.share ?? 0) > 0)
    const widest = seen.reduce<Viewpoint | null>((best, vp) => (!best || (vp.share ?? 0) > (best.share ?? 0) ? vp : best), null)
    const nearest = seen.reduce<Viewpoint | null>((best, vp) => (!best || (vp.distanceKm ?? Infinity) < (best.distanceKm ?? Infinity) ? vp : best), null)
    const limit = Number(params.limit)
    const percent = (share: number | undefined) => `${Math.round((share ?? 0) * 100)}%`
    return [
      { type: 'heading', level: 1, content: 'Summary' },
      {
        type: 'paragraph',
        content: seen.length
          ? `${seen.length} of ${viewpoints.length} viewpoints see the opening. The most is seen from ${widest?.name} (${percent(widest?.share)}); the nearest is ${nearest?.name}, ${nearest?.distanceKm} km away.`
          : 'No viewpoint sees the opening.',
        source: 'computed',
      },
      ...(widest && (widest.share ?? 0) > limit
        ? [{ use: 'notice', with: { title: 'Above the limit.', text: `The share seen from ${widest.name} passes ${percent(limit)}.` } }]
        : []),
      {
        type: 'fields',
        rows: [
          { label: 'Viewpoints seeing it', value: String(seen.length), source: 'computed' },
          { label: 'Assessed by', value: '{{ params.assessor }}' },
          { label: 'Checked by', value: '{{ project.checker }}' },
        ],
      },
    ]
  },
})
