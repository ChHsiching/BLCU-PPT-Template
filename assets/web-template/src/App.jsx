import { useCallback, useEffect, useMemo, useState } from 'react'
import deck from './deck.json'
import manifest from './manifest.json'
import { CoverPage } from './pages/CoverPage'
import { AgendaPage } from './pages/AgendaPage'
import { TextFormulaPage } from './pages/TextFormulaPage'
import { TextImagePage } from './pages/TextImagePage'
import { ChartFocusPage } from './pages/ChartFocusPage'
import { ContentBrand, CoverBrand } from './components/BrandLayer'
import { STAGE_W, STAGE_H, regionStyle, ptToPx, fontStack, buildSteps, firstStepOfPage } from './lib/layout'

const T = manifest.typography
const FORWARD_KEYS = new Set(['ArrowRight', 'ArrowDown', 'PageDown', ' ', 'Enter'])
const BACK_KEYS = new Set(['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace'])

function PageNumber({ arch, pageIndex }) {
  const region = arch.regions.page_number
  if (!region) return null // cover/closing carry none, like the pptx masters
  const t = T.page_number
  return (
    <div
      className="page-number"
      style={{
        ...regionStyle(region),
        fontFamily: fontStack(t.face, t.latin),
        fontSize: ptToPx(t.size_pt),
      }}
    >
      {pageIndex + 1}
    </div>
  )
}

function Slide({ page, pageIndex, revealed }) {
  const arch = manifest.archetypes[page.archetype]
  const props = { page, arch }
  const revealedProps = { ...props, revealed }
  const isCover = page.archetype === 'cover' || page.archetype === 'closing'
  return (
    <div className="slide">
      {/* master brand layer first: content text draws on the bands */}
      {isCover ? <CoverBrand /> : <ContentBrand />}
      {isCover ? (
        <CoverPage {...props} />
      ) : page.archetype === 'agenda' ? (
        <AgendaPage {...revealedProps} />
      ) : page.archetype === 'text-formula' ? (
        <TextFormulaPage {...props} />
      ) : page.archetype === 'text-image' ? (
        <TextImagePage {...revealedProps} />
      ) : (
        <ChartFocusPage {...props} />
      )}
      <PageNumber arch={arch} pageIndex={pageIndex} />
    </div>
  )
}

function useStageScale() {
  const [scale, setScale] = useState(0)
  useEffect(() => {
    const fit = () => {
      setScale(Math.min(window.innerWidth / STAGE_W, window.innerHeight / STAGE_H))
    }
    fit()
    window.addEventListener('resize', fit)
    return () => window.removeEventListener('resize', fit)
  }, [])
  return scale
}

export default function App() {
  const steps = useMemo(() => buildSteps(deck.pages), [])
  const [step, setStep] = useState(() => {
    const fromHash = Number.parseInt(window.location.hash.slice(1), 10)
    if (Number.isInteger(fromHash) && fromHash >= 1 && fromHash <= deck.pages.length) {
      const first = firstStepOfPage(steps, fromHash - 1)
      if (first >= 0) return first
    }
    return 0
  })
  const scale = useStageScale()

  const advance = useCallback(() => setStep((s) => Math.min(s + 1, steps.length - 1)), [steps.length])
  const back = useCallback(() => setStep((s) => Math.max(s - 1, 0)), [])

  useEffect(() => {
    const onKey = (e) => {
      if (FORWARD_KEYS.has(e.key)) {
        e.preventDefault()
        advance()
      } else if (BACK_KEYS.has(e.key)) {
        e.preventDefault()
        back()
      } else if (e.key === 'Home') {
        e.preventDefault()
        setStep(0)
      } else if (e.key === 'End') {
        e.preventDefault()
        setStep(steps.length - 1)
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [advance, back, steps.length])

  const { page, revealed } = steps[step]

  useEffect(() => {
    window.history.replaceState(null, '', `#${page + 1}`) // deep link for QA / reload
    const titleBlock = deck.pages[0].blocks.find((b) => b.type === 'title')
    if (titleBlock) document.title = titleBlock.text
  }, [page])

  // external hash navigation (deep links, QA jumps) lands on the target
  // page's first step
  useEffect(() => {
    const onHashChange = () => {
      const n = Number.parseInt(window.location.hash.slice(1), 10)
      if (!Number.isInteger(n) || n < 1 || n > deck.pages.length) return
      const first = firstStepOfPage(steps, n - 1)
      if (first >= 0) setStep(first)
    }
    window.addEventListener('hashchange', onHashChange)
    return () => window.removeEventListener('hashchange', onHashChange)
  }, [steps])

  return (
    <div className="viewport" onClick={advance}>
      <div
        className="stage"
        style={{ transform: `scale(${scale})` }}
        role="application"
        aria-label={`演示 第 ${page + 1} 页，共 ${deck.pages.length} 页`}
      >
        <div key={page} className="slide-enter">
          <Slide page={deck.pages[page]} pageIndex={page} revealed={revealed} />
        </div>
      </div>
      <div className="controls" onClick={(e) => e.stopPropagation()}>
        <button type="button" aria-label="上一步" onClick={back}>
          ‹
        </button>
        <span className="progress">
          {page + 1} / {deck.pages.length}
        </span>
        <button type="button" aria-label="下一步" onClick={advance}>
          ›
        </button>
      </div>
    </div>
  )
}
