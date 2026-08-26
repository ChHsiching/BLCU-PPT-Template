import { splitEmphasis, emphasisStyle } from '../lib/layout'

// Inline runs of one string: **keyword** segments become emphasis runs
// (tokens.emphasis — green bold), the web counterpart of the pptx
// emphasis=True runs. Markers never survive into the DOM text.
export function EmphasisText({ text }) {
  return splitEmphasis(text).map(({ text: seg, emph }, i) =>
    emph
      ? <strong className="emph" key={i} style={emphasisStyle()}>{seg}</strong>
      : seg)
}
