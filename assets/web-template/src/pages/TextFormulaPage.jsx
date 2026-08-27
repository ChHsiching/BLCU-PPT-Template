import manifest from '../manifest.json'
import { Formula } from '../components/Formula'
import { TitleBar } from '../components/TitleBar'
import { EmphasisText } from '../components/EmphasisText'
import {
  regionStyle, textboxPadding, rhythmLineHeight, paraBeforePx, ptToPx,
  regionRole, regionRoleName, titleOf,
} from '../lib/layout'

const TK = manifest.typography.tokens
// The pptx math paragraphs share one rhythm (render_pptx._MATH_PARA_SHELL:
// 130% line + 16pt before, first paragraph included inside the box) —
// mirrored here as the formula blocks' top margin.
const MATH_GAP_PT = 16

// text-formula: ONE full-height content flow (the template author's own
// idiom) — centered formula lines and left-aligned body paragraphs
// interleaved in block order, mirroring render_pptx.fill_text_formula's
// ContentArea. The flow vertically centers in the region (the pptx side uses
// anchor="ctr"): a partial page reads as composed breathing room, not a
// bottom void.
export function TextFormulaPage({ page, arch }) {
  const archetype = page.archetype
  const blocks = page.blocks.filter(b => b.type === 'formula' || b.type === 'text')

  return (
    <>
      <TitleBar text={titleOf(page)} arch={arch} archetype={archetype} />
      <div
        className="region-text"
        data-role={regionRoleName(archetype, 'text')}
        data-rhythm="1"
        style={{
          ...regionStyle(arch.regions.content),
          display: 'flex', flexDirection: 'column', justifyContent: 'center',
          ...textboxPadding(),
          ...rhythmLineHeight(),
          ...regionRole(archetype, 'text'),
        }}
      >
        {blocks.map((b, i) => b.type === 'formula'
          ? (
            <div key={i} style={{ marginTop: ptToPx(MATH_GAP_PT) }}>
              <Formula latex={b.latex} />
            </div>
          )
          : (
            <p key={i} style={{ marginTop: paraBeforePx() }}>
              <EmphasisText text={b.text} />
            </p>
          ))}
      </div>
    </>
  )
}
