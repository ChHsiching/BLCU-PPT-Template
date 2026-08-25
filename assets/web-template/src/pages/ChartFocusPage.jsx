import manifest from '../manifest.json'
import { TitleBar } from '../components/TitleBar'
import { assetUrl } from './TextImagePage'
import { regionStyle, ptToPx, fontStack, blocksOf, textsOf, titleOf } from '../lib/layout'

const T = manifest.typography

// chart-focus: one dominant chart image in the chart region + a short side
// comment at secondary body size (render_pptx.fill_chart_focus).
export function ChartFocusPage({ page, arch }) {
  const body = T.body
  const chart = blocksOf(page, 'image')[0]
  const comments = textsOf(page)
  const chartRegion = arch.regions.chart

  return (
    <>
      <TitleBar text={titleOf(page)} arch={arch} />
      {chart && (
        <div className="image-slot" style={regionStyle(chartRegion)}>
          <img src={assetUrl(chart.path)} alt={comments[0] ?? ''} draggable={false} />
        </div>
      )}
      {comments.length > 0 && (
        <div
          className="region-text"
          style={{
            ...regionStyle(arch.regions.comment),
            fontFamily: fontStack(body.face),
            fontSize: ptToPx(body.secondary_size_pt),
          }}
        >
          {comments.map((t, i) => (
            <p key={i}>{t}</p>
          ))}
        </div>
      )}
    </>
  )
}
