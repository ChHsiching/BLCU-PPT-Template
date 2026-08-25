import manifest from '../manifest.json'
import { regionStyle, ptToPx, fontStack, titleSizePt } from '../lib/layout'

const T = manifest.typography

// Content-page title bar: the page's title block in the archetype's title
// region, vertically centered like the template's title placeholder.
export function TitleBar({ text, arch }) {
  const t = T.title
  return (
    <div
      className="title-bar"
      style={{
        ...regionStyle(arch.regions.title),
        fontFamily: fontStack(t.face, t.latin),
        fontSize: ptToPx(titleSizePt(text)),
      }}
    >
      {text}
    </div>
  )
}
