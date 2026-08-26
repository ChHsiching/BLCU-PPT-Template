import manifest from '../manifest.json'
import { Formula } from '../components/Formula'
import { TitleBar } from '../components/TitleBar'
import { regionStyle, ptToPx, fontStack, blocksOf, textsOf, titleOf } from '../lib/layout'

const T = manifest.typography

// text-formula: title + stacked display formulas + bottom note band; with no
// formulas the page is a pure-text page using the full-height region (both
// variants mirror render_pptx.fill_text_formula).
export function TextFormulaPage({ page, arch }) {
  const formulas = blocksOf(page, 'formula').map((b) => b.latex)
  const texts = textsOf(page)
  const accent = T.accent
  const body = T.body

  return (
    <>
      <TitleBar text={titleOf(page)} arch={arch} />
      {formulas.length > 0 && (
        <div
          className="formula-area"
          style={{
            ...regionStyle(arch.regions.formula),
            fontSize: ptToPx(T.formula.size_pt),
          }}
        >
          {formulas.map((latex, i) => (
            <Formula key={i} latex={latex} />
          ))}
        </div>
      )}
      {texts.length > 0 && (
        <div
          className="region-text"
          style={{
            ...regionStyle(formulas.length > 0 ? arch.regions.text : arch.regions.text_full),
            fontFamily: fontStack(formulas.length > 0 ? accent.face : body.face),
            fontSize: ptToPx(formulas.length > 0 ? accent.size_pt : body.size_pt),
            // 纯文字页与 pptx 同步：150% 行距 + 段前 12pt（模板 slide 21 实测）
            ...(formulas.length === 0 && { lineHeight: 1.5 }),
          }}
        >
          {texts.map((t, i) => (
            <p key={i} style={formulas.length === 0 ? { marginTop: i === 0 ? 0 : ptToPx(12) } : undefined}>
              {t}
            </p>
          ))}
        </div>
      )}
    </>
  )
}
