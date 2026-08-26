import { TitleBar } from '../components/TitleBar'
import { FittedImage } from '../components/FittedImage'
import { EmphasisText } from '../components/EmphasisText'
import {
  regionStyle, textboxPadding, rhythmLineHeight, paraBeforePx,
  regionRole, regionRoleName, blocksOf, textsOf, titleOf,
} from '../lib/layout'

// chart-focus: one dominant chart image in the chart region + a short side
// comment (render_pptx.fill_chart_focus); the comment flows with the rhythm
// and the emphasis convention, like every body-flow role.
export function ChartFocusPage({ page, arch }) {
  const archetype = page.archetype
  const chart = blocksOf(page, 'image')[0]
  const comments = textsOf(page)

  return (
    <>
      <TitleBar text={titleOf(page)} arch={arch} archetype={archetype} />
      {chart && <FittedImage slot={arch.regions.chart} block={chart} archetype={archetype} />}
      {comments.length > 0 && (
        <div
          className="region-text"
          data-role={regionRoleName(archetype, 'comment')}
          data-rhythm="1"
          style={{
            ...regionStyle(arch.regions.comment),
            ...textboxPadding(),
            ...rhythmLineHeight(),
            ...regionRole(archetype, 'comment'),
          }}
        >
          {comments.map((t, i) => (
            <p key={i} style={{ marginTop: paraBeforePx() }}>
              <EmphasisText text={t} />
            </p>
          ))}
        </div>
      )}
    </>
  )
}
