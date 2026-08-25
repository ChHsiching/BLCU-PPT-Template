import manifest from '../manifest.json'
import {
  regionStyle,
  ptToPx,
  fontStack,
  titleSizePt,
  titleOf,
  textsOf,
} from '../lib/layout'

const T = manifest.typography

// cover + closing share one layout (closing is a content variant of cover in
// the manifest): centered title in the upper band, joined presenter line in
// the lower band.
export function CoverPage({ page, arch }) {
  const title = titleOf(page)
  const subtitle = textsOf(page).join(' ')
  const t = T.title
  const sub = T.cover_subtitle
  return (
    <>
      <div
        className="region-text cover-title"
        style={{
          ...regionStyle(arch.regions.title),
          fontFamily: fontStack(t.face, t.latin),
          fontSize: ptToPx(titleSizePt(title)),
        }}
      >
        {title}
      </div>
      {subtitle && (
        <div
          className="region-text cover-subtitle"
          style={{
            ...regionStyle(arch.regions.subtitle),
            fontFamily: fontStack(sub.face, sub.latin),
            fontSize: ptToPx(sub.size_pt),
          }}
        >
          {subtitle}
        </div>
      )}
    </>
  )
}
