import manifest from '../manifest.json'
import { Formula } from '../components/Formula'
import { TitleBar } from '../components/TitleBar'
import { EmphasisText } from '../components/EmphasisText'
import {
  regionStyle, textboxPadding, rhythmLineHeight, paraBeforePx, ptToPx,
  regionRole, regionRoleName, blocksOf, textsOf, titleOf, inToPx,
} from '../lib/layout'

const TK = manifest.typography.tokens
// The pptx math paragraphs share one rhythm (render_pptx._MATH_PARA_SHELL:
// 130% line + 16pt before, first included inside the box) — mirrored as the
// stack's gap and top padding so formulas land where the pptx puts them.
const MATH_GAP_PT = 16
const INSET_V_IN = TK.spacing.textbox_inset_v_in

// text-formula: title + stacked display formulas + bottom note band; with no
// formulas the page is a pure-text page using the full-height region (both
// variants mirror render_pptx.fill_text_formula). role_bindings map text and
// text_full to the body role; body paragraphs flow with the rhythm and the
// emphasis convention.
export function TextFormulaPage({ page, arch }) {
  const archetype = page.archetype
  const formulas = blocksOf(page, 'formula').map((b) => b.latex)
  const texts = textsOf(page)

  return (
    <>
      <TitleBar text={titleOf(page)} arch={arch} archetype={archetype} />
      {formulas.length > 0 && (
        <div
          className="formula-area"
          style={{
            ...regionStyle(arch.regions.formula),
            ...textboxPadding(),
            paddingTop: inToPx(INSET_V_IN) + ptToPx(MATH_GAP_PT),
            rowGap: ptToPx(MATH_GAP_PT),
            fontSize: ptToPx(TK.roles.formula.size_pt),
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
          data-role={regionRoleName(archetype, 'text')}
          data-rhythm="1"
          style={{
            ...regionStyle(formulas.length > 0 ? arch.regions.text : arch.regions.text_full),
            // no-formula pages vertically center their paragraphs — the
            // pptx side mirrors this with anchor="ctr" on TextFullArea
            ...(formulas.length === 0 && {
              display: 'flex', flexDirection: 'column', justifyContent: 'center',
            }),
            ...textboxPadding(),
            ...rhythmLineHeight(),
            ...regionRole(archetype, 'text'),
          }}
        >
          {texts.map((t, i) => (
            <p key={i} style={{ marginTop: paraBeforePx() }}>
              <EmphasisText text={t} />
            </p>
          ))}
        </div>
      )}
    </>
  )
}
